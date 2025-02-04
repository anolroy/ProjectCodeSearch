VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmElectricityUsage 
   BackColor       =   &H00FFEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electricity Usage"
   ClientHeight    =   12090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmElectricityUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   12090
   ScaleWidth      =   11280
   Begin VB.TextBox txtTabbingControl 
      Height          =   495
      Left            =   8760
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmElectricityUsage.frx":0442
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -20
      ScaleHeight     =   825
      ScaleWidth      =   8400
      TabIndex        =   16
      Top             =   0
      Width           =   8435
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Form for Electricity Charges "
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   9000
      Width           =   8415
      Begin VB.TextBox txtDueDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   5640
         TabIndex        =   35
         Top             =   80
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtInputData 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   215
         Left            =   6720
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtInputDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   3680
         TabIndex        =   26
         Top             =   80
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.CheckBox chkSelectAllLeases 
         Caption         =   "Select all "
         Height          =   255
         Left            =   7395
         TabIndex        =   18
         Top             =   100
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeases 
         Height          =   1935
         Left            =   75
         TabIndex        =   28
         Top             =   765
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   720
         VariousPropertyBits=   276824083
         Caption         =   "Due Date:"
         Size            =   "1270;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   210
         Index           =   11
         Left            =   2200
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   345
         VariousPropertyBits=   8388627
         Caption         =   "TotalSelection"
         Size            =   "609;370"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   27
         Left            =   7400
         TabIndex        =   29
         Top             =   360
         Width           =   570
         VariousPropertyBits=   276824083
         Caption         =   "CurrentBalance"
         Size            =   "1005;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   2880
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   870
         VariousPropertyBits=   276824083
         Caption         =   "Input Date:"
         Size            =   "1535;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   23
         Left            =   4440
         TabIndex        =   25
         Top             =   435
         Width           =   360
         VariousPropertyBits=   276824083
         Caption         =   "Date"
         Size            =   "635;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   26
         Left            =   6700
         TabIndex        =   24
         Top             =   440
         Width           =   615
         VariousPropertyBits=   276824083
         Caption         =   "Amount"
         Size            =   "1085;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   25
         Left            =   6060
         TabIndex        =   23
         Top             =   440
         Width           =   525
         VariousPropertyBits=   276824083
         Caption         =   "Usages"
         Size            =   "926;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   24
         Left            =   5360
         TabIndex        =   22
         Top             =   360
         Width           =   585
         VariousPropertyBits=   276824083
         Caption         =   "Current Reading"
         Size            =   "1032;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   22
         Left            =   2520
         TabIndex        =   21
         Top             =   440
         Width           =   855
         VariousPropertyBits=   276824083
         Caption         =   "Description"
         Size            =   "1508;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   21
         Left            =   1020
         TabIndex        =   20
         Top             =   440
         Width           =   510
         VariousPropertyBits=   276824083
         Caption         =   "Lessee"
         Size            =   "900;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   20
         Left            =   160
         TabIndex        =   19
         Top             =   440
         Width           =   315
         VariousPropertyBits=   276824083
         Caption         =   "Unit"
         Size            =   "556;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   105
         Width           =   2100
         VariousPropertyBits=   276824083
         Caption         =   "Select Lease to input records:"
         Size            =   "3704;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   405
         Left            =   75
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   8415
      Begin VB.ListBox lstFund 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmElectricityUsage.frx":0485
         Left            =   7800
         List            =   "frmElectricityUsage.frx":048C
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lstDemandTypes 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmElectricityUsage.frx":0499
         Left            =   1440
         List            =   "frmElectricityUsage.frx":04A0
         TabIndex        =   33
         Top             =   600
         Width           =   5415
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   5
         Left            =   7800
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   885
         VariousPropertyBits=   276824083
         Caption         =   "Select Fund:"
         Size            =   "1561;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2280
         VariousPropertyBits=   276824083
         Caption         =   "Select Demand Types:"
         Size            =   "4022;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   8415
      Begin VB.ListBox lstProperties 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "frmElectricityUsage.frx":04B4
         Left            =   1440
         List            =   "frmElectricityUsage.frx":04BB
         TabIndex        =   32
         Top             =   600
         Width           =   5415
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1140
         VariousPropertyBits=   276824083
         Caption         =   "Select Property:"
         Size            =   "2011;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   8415
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   5580
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblFrameIndex 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "1"
         Height          =   195
         Left            =   8280
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         Index           =   1
         X1              =   0
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   8415
      Begin MSForms.ListBox lstClients 
         DataSource      =   "Adodc1"
         Height          =   1695
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   5415
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "9551;2990"
         ColumnCount     =   2
         cColumnInfo     =   1
         MatchEntry      =   0
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763"
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   960
         VariousPropertyBits=   276824083
         Caption         =   "Select Client:"
         Size            =   "1693;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmElectricityUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LD_ID = 9
Private Const UU_ID = 10

Private Const MAX_NUMBER_FRAME_INDEX = 7
Private bPreviewShowed As Boolean
Private szDemandCharges As String
Private iSelectedFlxRec As Integer
Private bDerivedFromSelAll As Boolean

Private bBackGotFocused As Boolean
Private bAllInputted As Boolean
Private iLeft As Integer, iTop As Integer, iCurRow As Integer
Private aListPropertyID() As String
Private aListDT_ID() As String

Private Sub chkSelectAllLeases_Click()
   Dim iRow As Integer

   If Not bDerivedFromSelAll Or Val(lblFrameIndex.Caption) > 4 Then Exit Sub

   If chkSelectAllLeases.Value Then
      iSelectedFlxRec = flxLeases.Rows - 1
   Else
      iSelectedFlxRec = 0
   End If

   For iRow = 1 To flxLeases.Rows - 1
      SelectRowFlxLease iRow, chkSelectAllLeases.Value
   Next iRow
End Sub

Private Sub chkSelectAllLeases_GotFocus()
   bDerivedFromSelAll = True
End Sub

Private Sub cmdBack_Click()
   lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
   If Val(lblFrameIndex.Caption) = 4 Then
      Label1(9).Visible = False
      txtInputDate.Visible = False
      Label1(4).Visible = False
      txtDueDate.Visible = False
      chkSelectAllLeases.Enabled = True
      txtTabbingControl.Visible = False
      txtInputData.Visible = False
      flxLeases.Enabled = True
      iCurRow = 0
   Else
      Frame1(Val(lblFrameIndex.Caption) - 1).ZOrder 0
      cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
      cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
      cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
   End If
End Sub

Private Sub cmdBack_GotFocus()
   bBackGotFocused = True
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Are you sure you want to cancel the routine?", vbQuestion + vbYesNo, "Global Lease Update") = vbNo Then Exit Sub
   Unload Me
End Sub

Private Sub cmdNext_Click()
   Dim adoConn As New ADODB.Connection

   If Val(lblFrameIndex.Caption) = 1 Then
      If lstClients.Value = "" Or IsNull(lstClients.Value) Then
         MsgBox "Please select a client from the list.", vbCritical + vbOKOnly, "Select Client"
         lstClients.SetFocus
         Exit Sub
      End If

      LoadProperties
      lstProperties.SetFocus
   End If

   If Val(lblFrameIndex.Caption) = 2 Then
      If lstProperties.text = "" Or IsNull(lstProperties.text) Then
         MsgBox "Please select a property from the list.", vbCritical + vbOKOnly, "Select Property"
         lstProperties.SetFocus
         Exit Sub
      End If

'     connect to database
      adoConn.Open getConnectionString

      LoadDemandType adoConn
'      LoadFund adoConn

      lstDemandTypes.SetFocus
      adoConn.Close
      Set adoConn = Nothing
   End If

   If Val(lblFrameIndex.Caption) = 3 Then
      If lstDemandTypes.text = "" Or IsNull(lstDemandTypes.text) Then
         MsgBox "Please select a demand type from the list.", vbCritical + vbOKOnly, "Select Demand Type"
         lstDemandTypes.SetFocus
         Exit Sub
      End If
'
'      If lstFund.text = "" Or IsNull(lstFund.text) Then
'         MsgBox "Please select a fund from the list.", vbCritical + vbOKOnly, "Select Fund"
'         lstFund.SetFocus
'         Exit Sub
'      End If

      LoadDataFlxLeases

      iLeft = flxLeases.Left + flxLeases.Width
      iTop = flxLeases.Top + flxLeases.Height
   End If

   If Val(lblFrameIndex.Caption) = 4 Then       'going to 5
      If iSelectedFlxRec = 0 Then
         MsgBox "Please select atleast a lease from the list.", vbCritical + vbOKOnly, "Select Lease"
         flxLeases.SetFocus
         Exit Sub
      End If
      Label1(9).Visible = True
      txtInputDate.text = Format(Date, "dd/mm/yyyy")
      txtInputDate.Visible = True
      Label1(4).Visible = True
      txtDueDate.text = Format(Date, "dd/mm/yyyy")
      txtDueDate.Visible = True
      chkSelectAllLeases.Enabled = False
      txtTabbingControl.Visible = True

      flxLeases.col = 5
      Call MoveDownPosition
      txtInputData.Left = iLeft
      txtInputData.Top = iTop
      txtInputData.Width = flxLeases.ColWidth(5)
      txtInputData.Visible = True
      txtInputData.SetFocus
      flxLeases.Enabled = False
      bAllInputted = False
   End If

   If Val(lblFrameIndex.Caption) = 5 Then
      If Not bAllInputted Then Exit Sub
   
      If MsgBox("Do you want to SAVE?", vbYesNo + vbQuestion, "Electricity Usage") = vbNo Then Exit Sub

      SaveEU
      ShowMsgInTaskBar "Data has been saved."
      cmdNext.Enabled = False
      cmdBack.Enabled = False
      cmdFinish.Enabled = True
      cmdCancel.Enabled = False
      Exit Sub
   End If

   If Val(lblFrameIndex.Caption) <> 4 Then
      Frame1(Val(lblFrameIndex.Caption)).Top = 860
      Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
      Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
   End If

   lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1
   cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
   cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
   cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
End Sub

Private Sub SaveEU()
'   szHeader$ = "|<Unit|<Tenant|<Decription|<Date|>CurrentReading|>Usage|>Amount|>Balance|<LeaseID|<UU_ID"
'               0  1     2       3           4     5               6      7       8        9        10
   Dim iRow As Integer
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM LUtilityUsage;"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   For iRow = 1 To flxLeases.Rows - 1
      If flxLeases.TextMatrix(iRow, 0) = "X" Then
         adoRST.AddNew
         adoRST.Fields.Item("UtilityUsage").Value = UniqueID()
         adoRST.Fields.Item("LeaseID").Value = flxLeases.TextMatrix(iRow, LD_ID)
         adoRST.Fields.Item("UtilityType").Value = "ELC"
         adoRST.Fields.Item("Description").Value = "ELECTRICITY USAGE"
         adoRST.Fields.Item("CurrentReading").Value = flxLeases.TextMatrix(iRow, 5)
         adoRST.Fields.Item("ReadingDate").Value = Format(txtInputDate.text, "dd mmmm yyyy")
         adoRST.Fields.Item("DueDate").Value = Format(txtDueDate.text, "dd mmmm yyyy")
         adoRST.Fields.Item("Usage").Value = flxLeases.TextMatrix(iRow, 6)
         adoRST.Fields.Item("Amount").Value = flxLeases.TextMatrix(iRow, 7)
         adoRST.Fields.Item("Balance").Value = Val(flxLeases.TextMatrix(iRow, 7)) + Val(flxLeases.TextMatrix(iRow, 8))
         adoRST.Fields.Item("IsGenerated").Value = False
         adoRST.Fields.Item("LatestReading").Value = True
         adoRST.Fields.Item("DemandType").Value = aListDT_ID(lstDemandTypes.ListIndex)
'         adoRST.Fields.Item("Fund").Value = lstFund.ItemData(lstFund.ListIndex)
         adoRST.Update
      End If
   Next iRow
   adoRST.Close

   For iRow = 1 To flxLeases.Rows - 1
      If flxLeases.TextMatrix(iRow, 0) = "X" And flxLeases.TextMatrix(iRow, UU_ID) <> "" Then
         szSQL = "SELECT * FROM LUtilityUsage " & _
                 "WHERE UtilityUsage = '" & flxLeases.TextMatrix(iRow, UU_ID) & "';"
         adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

         If Not adoRST.Fields.Item("IsGenerated").Value Then adoRST.Fields.Item("BF").Value = True
         adoRST.Fields.Item("LatestReading").Value = False
         adoRST.Update
         adoRST.Close
      End If
      If flxLeases.TextMatrix(iRow, 0) = "X" And flxLeases.TextMatrix(iRow, LD_ID) <> "" Then
         szSQL = "UPDATE LeaseDetails " & _
                 "SET UtilityUsage = 'Y' " & _
                 "WHERE LeaseID = '" & flxLeases.TextMatrix(iRow, LD_ID) & "';"
         adoConn.Execute szSQL
      End If
   Next iRow

   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function MoveRightPosition() As Boolean
   flxLeases.col = flxLeases.col + 1
   If flxLeases.col > 7 Then
      MoveRightPosition = False
      Exit Function
   End If
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveRightPosition = True
End Function

Private Function MoveLeftPosition() As Boolean
   flxLeases.col = flxLeases.col - 1
   If flxLeases.col < 5 Then
      MoveLeftPosition = False
      Exit Function
   End If
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveLeftPosition = True
End Function

Private Function MoveUpPosition() As Boolean
   Dim iRow As Integer

   For iRow = iCurRow - 1 To 1 Step -1
      If flxLeases.TextMatrix(iRow, 0) = "X" Then
         flxLeases.row = iRow
         iCurRow = iRow
         Exit For
      End If
   Next iRow
   If iRow = 0 Then
      iCurRow = 0
      MoveUpPosition = False
      Exit Function
   End If

   If flxLeases.col < 5 Then flxLeases.col = 5
   If flxLeases.col > 7 Then flxLeases.col = 7
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveUpPosition = True
End Function

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer

   For iRow = iCurRow + 1 To flxLeases.Rows - 1
      If flxLeases.TextMatrix(iRow, 0) = "X" Then
         flxLeases.row = iRow
         iCurRow = iRow
         Exit For
      End If
   Next iRow
   If iRow = flxLeases.Rows Then
      iCurRow = 0
      MoveDownPosition = False
      Exit Function
   End If

   If flxLeases.col < 5 Then flxLeases.col = 5
   If flxLeases.col > 7 Then flxLeases.col = 7
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveDownPosition = True
End Function

Private Sub LoadDataFlxLeases()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT LUU.*, Units.UnitNumber, Tenants.Name " & _
           "FROM LeaseDetails, Units, Tenants, " & _
             "(SELECT LeaseDetails.LeaseID, UU.ReadingDate AS ReadingDate, " & _
                  "UU.Description, UU.CurrentReading, UU.ReadingDate, " & _
                  "UU.Usage , UU.Amount, UU.Balance, UU.IsGenerated, " & _
                  "UU.UtilityUsage " & _
             " FROM LeaseDetails LEFT OUTER JOIN " & _
                  "(SELECT * FROM LUtilityUsage WHERE UtilityType = 'ELC' AND " & _
                   "LatestReading = TRUE) AS UU " & _
                  "ON LeaseDetails.LeaseID = UU.LeaseID" & _
             ") AS LUU " & _
           "WHERE Units.PropertyID = '" & aListPropertyID(lstProperties.ListIndex) & "' AND " & _
               "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber AND " & _
               "LeaseDetails.Status = TRUE AND " & _
               "LeaseDetails.LeaseID = LUU.LeaseID;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'   szHeader$ = "|<Unit|<Tenant|<Decription|<Date|>CurrentReading|>Usage|>Amount|>Balance|<LeaseID|<UU_ID"
   flxLeases.Clear
   flxLeases.Rows = 2
   r = 1
   While Not adoRST.EOF
      flxLeases.TextMatrix(r, 0) = ""
      flxLeases.TextMatrix(r, 1) = adoRST.Fields.Item("UnitNumber").Value
      flxLeases.TextMatrix(r, 2) = adoRST.Fields.Item("Name").Value
      flxLeases.TextMatrix(r, 3) = IIf(IsNull(adoRST.Fields.Item("Description").Value), "", _
                                       adoRST.Fields.Item("Description").Value)
      If IsNull(adoRST.Fields.Item("ReadingDate").Value) Then
         flxLeases.TextMatrix(r, 4) = ""
      Else
         flxLeases.TextMatrix(r, 4) = Format(CDate(adoRST.Fields.Item("ReadingDate").Value), "DD/MM/YYYY")
      End If
      flxLeases.TextMatrix(r, 5) = IIf(IsNull(adoRST.Fields.Item("CurrentReading").Value), "", _
                                       adoRST.Fields.Item("CurrentReading").Value)
      flxLeases.TextMatrix(r, 6) = IIf(IsNull(adoRST.Fields.Item("Usage").Value), "", _
                                       adoRST.Fields.Item("Usage").Value)
      If IsNull(adoRST.Fields.Item("Amount").Value) Then
         flxLeases.TextMatrix(r, 7) = ""
      Else
         flxLeases.TextMatrix(r, 7) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
      End If
      If Not adoRST.Fields.Item("IsGenerated").Value Then
         If IsNull(adoRST.Fields.Item("Balance").Value) Then
            flxLeases.TextMatrix(r, 8) = ""
         Else
            flxLeases.TextMatrix(r, 8) = Format(adoRST.Fields.Item("Balance").Value, "0.00")
         End If
      Else
         flxLeases.TextMatrix(r, 8) = ""
      End If
      flxLeases.TextMatrix(r, LD_ID) = IIf(IsNull(adoRST.Fields.Item("LeaseID").Value), "", _
                                       adoRST.Fields.Item("LeaseID").Value)
      flxLeases.TextMatrix(r, UU_ID) = IIf(IsNull(adoRST.Fields.Item("UtilityUsage").Value), "", _
                                       adoRST.Fields.Item("UtilityUsage").Value)
      adoRST.MoveNext
      If Not adoRST.EOF Then flxLeases.AddItem ""
      r = r + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   iSelectedFlxRec = 0
End Sub

Private Sub LoadDemandType(ByVal adoConn As ADODB.Connection)
   Dim szSQL As String, r As Integer
   Dim adoRST As New ADODB.Recordset

   lstDemandTypes.Clear

   szSQL = "SELECT ID, TYPE " & _
           "FROM DEMANDTYPES " & _
           "WHERE PropertyID = '" & aListPropertyID(lstProperties.ListIndex) & "' OR PropertyID = 'ALL';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim aListDT_ID(adoRST.RecordCount - 1) As String

   r = 0
   While Not adoRST.EOF
      If InStr(UCase(adoRST.Fields.Item("TYPE").Value), "ELECTRICITY") > 0 Then
         aListDT_ID(r) = adoRST.Fields.Item("ID").Value
         lstDemandTypes.List(r) = adoRST.Fields.Item("TYPE").Value
         r = r + 1
      End If
      adoRST.MoveNext
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub LoadFund(ByVal adoConn As ADODB.Connection)
   ' Error Handler
   On Error GoTo Error_Handler

   Dim rRow As Integer, iRec As Integer
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
      rRow = 0
      lstFund.Clear
      While Not adoRST.EOF
         lstFund.AddItem adoRST.Fields.Item("FundName").Value
         lstFund.ItemData(rRow) = adoRST.Fields.Item("FundID").Value
         rRow = rRow + 1
         adoRST.MoveNext
      Wend
   End If

   ' Destroy Objects
   Set adoRST = Nothing
   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "Error in Loading fund.", vbExclamation, "Loading Fund"
   ' Destroy Objects
   Set adoRST = Nothing
End Sub

Private Sub LoadProperties()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   lstProperties.Clear

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & lstClients.Value & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   ReDim aListPropertyID(adoRST.RecordCount - 1) As String

   r = 0
   While Not adoRST.EOF
      aListPropertyID(r) = adoRST.Fields.Item("PROPERTYID").Value
      lstProperties.List(r) = adoRST.Fields.Item("PROPERTYNAME").Value
      r = r + 1
      adoRST.MoveNext
   Wend

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdFinish_Click()
   Form_Unload 0
End Sub

Private Sub flxLeases_Click()
   If flxLeases.row > 0 And Val(lblFrameIndex.Caption) < 5 Then
      If flxLeases.TextMatrix(flxLeases.row, 0) = "X" Then
         iSelectedFlxRec = iSelectedFlxRec - 1

         SelectRowFlxLease flxLeases.row, False
         If chkSelectAllLeases.Value Then chkSelectAllLeases.Value = False
      Else
         iSelectedFlxRec = iSelectedFlxRec + 1

         SelectRowFlxLease flxLeases.row, True
      End If
   End If
End Sub

Private Sub SelectRowFlxLease(iRow As Integer, bSel As Boolean)
   Dim c As Integer

   flxLeases.TextMatrix(iRow, 0) = IIf(bSel, "X", "")
   flxLeases.row = iRow
   For c = 0 To flxLeases.Cols - 1
      flxLeases.col = c
      flxLeases.CellBackColor = IIf(bSel, &HE0E0E0, vbWhite)
   Next c
   Label1(11).Visible = IIf(iSelectedFlxRec > 0, True, False)
   Label1(11).Caption = iSelectedFlxRec
End Sub

Private Sub flxLeases_GotFocus()
   bDerivedFromSelAll = False
End Sub

Private Sub Form_Load()
   Me.Width = 8490
   Me.Height = 5040
   Me.Top = frmMMain.Height / 2 - Me.Height / 2
   Me.Left = frmMMain.Width / 2 - Me.Width / 2

   Me.BackColor = MODULEBACKCOLOR
   Frame1(0).BackColor = Me.BackColor
   Frame1(1).BackColor = Me.BackColor
   Frame1(2).BackColor = Me.BackColor
   Frame1(3).BackColor = Me.BackColor
   Frame1(8).BackColor = Me.BackColor
   chkSelectAllLeases.BackColor = Me.BackColor

   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT;"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   While Not adoRST.EOF
      lstClients.AddItem adoRST.Fields.Item("CLIENTID").Value
      lstClients.List(r, 1) = adoRST.Fields.Item("CLIENTNAME").Value
      r = r + 1
      adoRST.MoveNext
   Wend

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   bPreviewShowed = False
   bDerivedFromSelAll = False
   bBackGotFocused = False
   txtInputDate.text = Format(Date, "dd/mm/yyyy")
   txtDueDate.text = Format(Date, "dd/mm/yyyy")
   ConfigureFlxLeases
End Sub

Private Sub ConfigureFlxLeases()
   Dim szHeader As String, iCol As Integer

   flxLeases.Clear
   flxLeases.Cols = 11
   flxLeases.Rows = 2
   flxLeases.RowHeight(0) = 0
   szHeader$ = "|<Unit|<Tenant|<Decription|<Date|>CurrentReading|>Usage|>Amount|>Balance|<LeaseID|<UU_ID"
   flxLeases.FormatString = szHeader$

   flxLeases.ColWidth(0) = 0
   For iCol = 21 To flxLeases.Cols + 20 - 4
      flxLeases.ColWidth(iCol - 20) = Label1(iCol).Left - Label1(iCol - 1).Left
   Next iCol
   flxLeases.ColWidth(iCol - 20 + 1) = 0
   flxLeases.ColWidth(iCol - 20 + 2) = 0

   flxLeases.row = 0
   flxLeases.col = 0

   iSelectedFlxRec = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub lstClients_GotFocus()
   bBackGotFocused = False
End Sub

Private Sub lstClients_LostFocus()
   cmdNext.SetFocus
End Sub

Private Sub lstDemandTypes_DblClick()
   If lstDemandTypes.ListCount = 0 Then MsgBox "There are no demand type ELECTRICITY.", vbInformation + vbOKOnly, "Electricity Usage"
End Sub

Private Sub lstDemandTypes_GotFocus()
   bBackGotFocused = False
End Sub

Private Sub lstDemandTypes_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 9 Then cmdNext.SetFocus
   If lstDemandTypes.ListCount = 0 Then MsgBox "There are no demand type ELECTRICITY.", vbInformation + vbOKOnly, "Electricity Usage"
End Sub

Private Sub lstProperties_GotFocus()
   bBackGotFocused = False
End Sub

Private Sub lstProperties_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 9 Then cmdNext.SetFocus
End Sub

Private Sub txtDueDate_Change()
   TextBoxChangeDate txtDueDate
End Sub

Private Sub txtDueDate_GotFocus()
   SelTxtInCtrl txtDueDate
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDueDate, KeyAscii
End Sub

Private Sub txtDueDate_LostFocus()
   TextBoxFormatDate txtDueDate
End Sub

Private Sub txtInputData_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      Call DigitTextKeyPress(txtInputData, KeyAscii)
   End If
End Sub

Private Sub txtInputData_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      flxLeases.TextMatrix(flxLeases.row, flxLeases.col) = txtInputData.text

      If txtInputData.text = "" Then
         MsgBox "Please type the new " & Label1(flxLeases.col + 19).Caption & ".", vbCritical & vbOKOnly, "Electricity Usage"
         txtInputData.SetFocus
         Exit Sub
      End If

      If MoveRightPosition Then
         txtInputData.Left = iLeft
         txtInputData.Top = iTop
         txtInputData.Width = flxLeases.ColWidth(flxLeases.col)
         txtInputData.text = ""
         txtInputData.Visible = True
         txtInputData.SetFocus
      Else
         flxLeases.col = 5
         If Not MoveDownPosition Then
            txtInputData.Visible = False
            cmdNext.SetFocus
            bAllInputted = True
            Exit Sub
         End If
         txtInputData.Left = iLeft
         txtInputData.Top = iTop
         txtInputData.Width = flxLeases.ColWidth(flxLeases.col)
         txtInputData.text = ""
         txtInputData.Visible = True
         txtInputData.SetFocus
      End If
   End If

   If KeyCode > 36 And KeyCode < 41 Then
      flxLeases.TextMatrix(flxLeases.row, flxLeases.col) = txtInputData.text

      If KeyCode = 37 Then _
         MoveLeftPosition                    'Left Key

      If KeyCode = 38 Then _
         MoveUpPosition                      'Up Key

      If KeyCode = 39 Then _
         MoveRightPosition                   'Right Key

      If KeyCode = 40 Then _
         MoveDownPosition                    'Down key

      txtInputData.Left = iLeft
      txtInputData.Top = iTop
      txtInputData.Width = flxLeases.ColWidth(flxLeases.col)
      txtInputData.text = ""
      txtInputData.Visible = True
      txtInputData.SetFocus
   End If
End Sub
'
'Private Sub txtInputData_LostFocus()
'   If txtInputData.text <> "" Then
'      flxLeases.TextMatrix(flxLeases.Row, flxLeases.Col) = txtInputData.text
'   End If
'End Sub

Private Sub txtInputDate_Change()
   TextBoxChangeDate txtInputDate
End Sub

Private Sub txtInputDate_GotFocus()
   SelTxtInCtrl txtInputDate
End Sub

Private Sub txtInputDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtInputDate, KeyAscii
End Sub

Private Sub txtInputDate_LostFocus()
   TextBoxFormatDate txtInputDate
End Sub

Private Sub txtTabbingControl_GotFocus()
   If txtInputData.text = "" Then
      MsgBox "Please type the new " & Label1(flxLeases.col + 19).Caption & ".", vbCritical & vbOKOnly, "Electricity Usage"
      txtInputData.SetFocus
      Exit Sub
   End If
   If MoveRightPosition Then
      txtInputData.Left = iLeft
      txtInputData.Top = iTop
      txtInputData.Width = flxLeases.ColWidth(flxLeases.col)
      txtInputData.text = ""
      txtInputData.Visible = True
      txtInputData.SetFocus
   Else
      flxLeases.col = 5
      If Not MoveDownPosition Then
         txtInputData.Visible = False
         cmdNext.SetFocus
         bAllInputted = True
         Exit Sub
      End If
      txtInputData.Left = iLeft
      txtInputData.Top = iTop
      txtInputData.Width = flxLeases.ColWidth(flxLeases.col)
      txtInputData.text = ""
      txtInputData.Visible = True
      txtInputData.SetFocus
   End If
End Sub
