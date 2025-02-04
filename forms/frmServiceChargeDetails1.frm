VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceChargeDetails1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Analysis"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   Icon            =   "frmServiceChargeDetails1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   9045
   Begin VB.Frame Frame1 
      Height          =   7065
      Index           =   1
      Left            =   80
      TabIndex        =   6
      Top             =   0
      Width           =   7240
      Begin VB.CommandButton cmdSCDBdNew 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdSCDBdClose 
         Caption         =   "Do&ne"
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
         Left            =   5910
         TabIndex        =   3
         Top             =   6600
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
         Left            =   4515
         TabIndex        =   5
         Top             =   6600
         Width           =   1215
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
         Left            =   3120
         TabIndex        =   4
         Top             =   6600
         Width           =   1215
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
         Left            =   1320
         TabIndex        =   15
         Top             =   6600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtNName 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   3225
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
         Height          =   315
         Left            =   5520
         TabIndex        =   8
         Top             =   6135
         Width           =   1200
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtBudget 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   360
         Width           =   1200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetailsAnalysis 
         Height          =   5385
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   9499
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
         Left            =   2280
         TabIndex        =   13
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Index           =   12
         Left            =   5040
         TabIndex        =   12
         Top             =   6195
         Width           =   390
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
         TabIndex        =   11
         Top             =   120
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
         Left            =   5520
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin MSForms.ComboBox cboNCode 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3836;556"
         TextColumn      =   1
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1056;70555"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   7440
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
            Picture         =   "frmServiceChargeDetails1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceChargeDetails1.frx":0E64
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceChargeDetails1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private flgChange As Integer   '0 -> New, 1 -> Edit
Private bDiscard  As Boolean   'test 01072014
Public bNewBudget As Boolean

Private Sub cboNCode_Change()
   If Trim(cboNCode.text) = "" Then Exit Sub
   If cboNCode.ListIndex >= 0 Then txtNName.text = cboNCode.Column(1)
End Sub

Private Sub cmdSCDBdCancel_Click()
   If MsgBox("Do you wish to cancel the amendment of the budget analysis?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      txtSCDBudgetTotal.text = ""
      Unload Me
   End If
End Sub

Private Sub cmdSCDBdDelete_Click()
   If MsgBox("Do you wish to delete?", vbQuestion + vbYesNo, "Saving") = vbNo Then Exit Sub

   With flxSCBudgetDetailsAnalysis
      .TextMatrix(.row, 5) = "D"
      .RowHeight(.row) = 0
      txtSCDBudgetTotal.text = Format(Val(txtSCDBudgetTotal.text) - Val(.TextMatrix(.row, 4)), "0.00")
   End With
   flgChange = 1
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub cmdSCDBDEdit_Click()
   If cmdSCDBdEdit.Caption <> "&Edit" Then
      If txtNName.text = "" Or cboNCode.Value = "" Then
            ShowMsgInTaskBar "Please select the nominal name/code.", , "N"
            txtNName.SetFocus
            Exit Sub
      End If
      If txtBudget.text = "" Then
            ShowMsgInTaskBar "Please enter the total budget.", , "N"
            txtBudget.SetFocus
            Exit Sub
      End If

      updateGrid
      flgChange = 1
      SCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
   End If
End Sub

Private Sub cmdSCDBdClose_Click()
   bDiscard = False
   Me.Hide
   Unload Me
End Sub

Private Sub cmdSCDClose_Click()
   initialiseGrid
   Me.Hide
End Sub

Private Sub ConfigureFlxBRMain()
   Dim szFlxHeader As String

   flxSCBudgetDetailsAnalysis.Rows = 1
   flxSCBudgetDetailsAnalysis.RowHeight(0) = 0
   flxSCBudgetDetailsAnalysis.Clear
   flxSCBudgetDetailsAnalysis.Cols = 6
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF"
   flxSCBudgetDetailsAnalysis.FormatString = szFlxHeader$

   flxSCBudgetDetailsAnalysis.ColWidth(0) = 0
   flxSCBudgetDetailsAnalysis.ColWidth(1) = 0
   flxSCBudgetDetailsAnalysis.ColWidth(2) = lblRentCharges(1).Left - lblRentCharges(0).Left
   flxSCBudgetDetailsAnalysis.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(1).Left
   flxSCBudgetDetailsAnalysis.ColWidth(4) = flxSCBudgetDetailsAnalysis.Width - lblRentCharges(2).Left - 300
   flxSCBudgetDetailsAnalysis.ColWidth(5) = 0
End Sub

Private Sub LoadFlxRCMain()
   Dim rowIndex As Integer
   Dim col As Integer

   If frmMMain.IsRibbonVersion Then
   rowIndex = frmServiceCharge1.txtMatrixRow.text
   Else
   rowIndex = frmServiceCharge1.txtMatrixRow.text
   End If

   For col = 0 To 59
      If frmMMain.IsRibbonVersion Then
      If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
         Exit For
      Else
         addLine frmServiceCharge1.getDetailsFromMatrix(rowIndex, col)
         frmServiceCharge1.fillBufferMatrix frmServiceCharge1.getDetailsFromMatrix(rowIndex, col), col + 1
      End If
      Else
      If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
         Exit For
      Else
         addLine frmServiceCharge1.getDetailsFromMatrix(rowIndex, col)
         frmServiceCharge1.fillBufferMatrix frmServiceCharge1.getDetailsFromMatrix(rowIndex, col), col + 1
      End If
      End If
   Next col
End Sub

Private Sub SCSumTotal()
   Dim iRow As Integer

   txtSCDBudgetTotal.text = "0.00"

   For iRow = 1 To flxSCBudgetDetailsAnalysis.Rows - 1
      If flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 5) <> "D" Then _
         txtSCDBudgetTotal.text = Format(Val(txtSCDBudgetTotal.text) + _
                                         Val(flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 4)), "0.00")
   Next iRow
End Sub

Private Sub cmdSCDBdNew_Click()
   If txtNName.text = "" Or cboNCode.Value = "" Then
      ShowMsgInTaskBar "Please select the nominal name/code.", , "N"
      cboNCode.SetFocus
      Exit Sub
   End If
   If txtBudget.text = "" Then
      ShowMsgInTaskBar "Please enter the total budget.", , "N"
      txtBudget.SetFocus
      Exit Sub
   End If

   updateGrid
   SCSumTotal

   flgChange = 0
   ControlsModeRentBudgetDetails DefaultMode
   cmdSCDBdNew.Picture = ImageList1.ListImages.Item(1).Picture
End Sub

Private Sub flxSCBudgetDetailsAnalysis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub flxSCBudgetDetailsAnalysis_DblClick()
   If flxSCBudgetDetailsAnalysis.TextMatrix(1, 0) = "" Then Exit Sub

   On Error Resume Next

   cboNCode.Value = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 2)
   txtBudget.text = flxSCBudgetDetailsAnalysis.TextMatrix(flxSCBudgetDetailsAnalysis.row, 4)
   cmdSCDBdNew.Picture = ImageList1.ListImages.Item(2).Picture

   ControlsModeRentBudgetDetails EditMode
   cboNCode.SetFocus
   flgChange = 1
End Sub

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         cboNCode.text = ""
         txtNName.text = ""
         txtBudget.text = ""

         flxSCBudgetDetailsAnalysis.Enabled = True
         flxSCBudgetDetailsAnalysis.row = 0
         flxSCBudgetDetailsAnalysis.col = 0
'         frmServiceCharge1Details.Show
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
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 7515
   Me.Height = 7650
   Me.BackColor = MODULEBACKCOLOR
   Me.Refresh
   Frame1(1).BackColor = MODULEBACKCOLOR
   bDiscard = True

   initialiseGrid
   SCSumTotal

   cmdSCDBdNew.Picture = ImageList1.ListImages.Item(1).Picture
   Call WheelHook(Me.hWnd)
End Sub

Private Sub initialiseGrid()
   Call ConfigureFlxBRMain
   flgChange = 0

   LoadFlxRCMain

   Call LoadFund
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub LoadFund()
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      Dim count As Integer
      count = adoRst.RecordCount
      ReDim Data(2, count) As String
      ReDim Data2(2, count) As String
      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("Code").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("Name").Value)
         Data2(1, rRow) = Trim(adoRst.Fields.Item("Name").Value)
         Data2(0, rRow) = Trim(adoRst.Fields.Item("Code").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
    
      cboNCode.Clear
      cboNCode.Column() = Data2()
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub updateGrid()
   With flxSCBudgetDetailsAnalysis
      If flgChange = 0 Then
         .AddItem ""
         If .TextMatrix(.Rows - 1, 0) <> "" Then .AddItem ""
         .TextMatrix(.Rows - 1, 0) = UniqueID()
         If frmMMain.IsRibbonVersion Then
         .TextMatrix(.Rows - 1, 1) = frmServiceCharge1.txtBudgetId.text
         Else
         .TextMatrix(.Rows - 1, 1) = frmServiceCharge1.txtBudgetId.text
         End If
         .TextMatrix(.Rows - 1, 2) = cboNCode.text
         .TextMatrix(.Rows - 1, 3) = txtNName.text
         .TextMatrix(.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
         .TextMatrix(.Rows - 1, 5) = "N"        'New
      End If
      If flgChange = 1 Then
         .TextMatrix(.row, 2) = cboNCode.text
         .TextMatrix(.row, 3) = txtNName.text
         .TextMatrix(.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
         .TextMatrix(.row, 5) = "M"             'Modified
      End If
   End With
End Sub

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
   frmServiceCharge1.Enabled = True
   Else
   frmServiceCharge1.Enabled = True
   End If
   
   If Not bDiscard Then
      SaveBudgetAnalysis
      If frmMMain.IsRibbonVersion Then
      frmServiceCharge1.cmdDetails.Enabled = True
      frmServiceCharge1.txtBudget.Locked = True
      Else
      frmServiceCharge1.cmdDetails.Enabled = True
      frmServiceCharge1.txtBudget.Locked = True
      End If
   Else
      If frmMMain.IsRibbonVersion Then
      frmServiceCharge1.cmdDetails.Enabled = False
      frmServiceCharge1.txtBudget.Locked = False
      Else
      frmServiceCharge1.cmdDetails.Enabled = False
      frmServiceCharge1.txtBudget.Locked = False
      End If
   End If
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
   frmServiceCharge1.txtBudget.text = txtSCDBudgetTotal.text
   Else
   frmServiceCharge1.txtBudget.text = txtSCDBudgetTotal.text
   End If

   flgChange = 0
End Sub

Private Sub arraySave(ByVal Index As Integer)
   Dim rowIndex As Integer
   Dim col As Integer
   Dim found As Integer

   With flxSCBudgetDetailsAnalysis
      found = 0
      rowIndex = frmServiceCharge1.txtMatrixRow.text
      For col = 0 To 59
         If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = .TextMatrix(Index, 0) Then
            found = 1
            Exit For
         Else
            If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
               Exit For
            End If
         End If
      Next col

      If (found = 1) Then
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setFlgDel .TextMatrix(Index, 5)
      Else
         If col <> 59 And .RowHeight(Index) > 0 Then
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetId frmServiceCharge1.txtBudgetId.text
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetDetailId .TextMatrix(Index, 0)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
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
      rowIndex = frmServiceCharge1.txtMatrixRow.text
      For col = 0 To 59
         If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = .TextMatrix(Index, 0) Then
            found = 1
            Exit For
         Else
            If frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).getBudgetDetailID = "" Then
               Exit For
            End If
         End If
      Next col

      If (found = 1) Then
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setFlgDel .TextMatrix(Index, 5)
      Else
         If col <> 59 And .RowHeight(Index) > 0 Then
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetId frmServiceCharge1.txtBudgetId.text
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetDetailId .TextMatrix(Index, 0)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNCode .TextMatrix(Index, 2)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setNName .TextMatrix(Index, 3)
            frmServiceCharge1.getDetailsFromMatrix(rowIndex, col).setBudgetAmount .TextMatrix(Index, 4)
         Else
             ShowMsgInTaskBar "Please save service charge budget before proceeding", , "N"
         End If
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtBudget, KeyAscii
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
