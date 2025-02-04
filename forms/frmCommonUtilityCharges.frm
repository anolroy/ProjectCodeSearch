VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCommonUtilityCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Common Utility Charges"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   3750
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommonUtilityCharges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11865
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "&Exclude"
      Height          =   285
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Copy from lease"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtAmtIC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtPcgIC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtAmtSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtPcgSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtAmtRC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtPcgRC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Canc&el"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Copy from lease"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Copy from lease"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame fraHeader 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   4
         Top             =   525
         Width           =   1455
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9720
         TabIndex        =   3
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txtDateIssue 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Top             =   165
         Width           =   1455
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox chkPro 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11445
         TabIndex        =   8
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtCurrReading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtPrevReading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtUsage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtTotalCharge 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8280
         MaxLength       =   70
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Charge"
         Height          =   255
         Left            =   7320
         TabIndex        =   52
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Usage"
         Height          =   255
         Left            =   5640
         TabIndex        =   51
         Top             =   1440
         Width           =   495
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   7560
         TabIndex        =   50
         Top             =   165
         Width           =   225
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "397;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cboClientList 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   165
         Width           =   3615
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6376;556"
         TextColumn      =   2
         ColumnCount     =   8
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin MSForms.ComboBox cboPropertyList 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   525
         Width           =   3615
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6376;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   40
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   39
         Top             =   165
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date:"
         Height          =   195
         Index           =   8
         Left            =   8805
         TabIndex        =   38
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date:"
         Height          =   195
         Index           =   7
         Left            =   5205
         TabIndex        =   37
         Top             =   165
         Width           =   900
      End
      Begin MSForms.ComboBox cboFund 
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   960
         Width           =   2655
         VariousPropertyBits=   1820346395
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4683;556"
         TextColumn      =   2
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;70555"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   9
         Left            =   4200
         TabIndex        =   36
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   7320
         TabIndex        =   35
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Type"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   34
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Prorata"
         Height          =   255
         Left            =   10800
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Reading"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Reading"
         Height          =   255
         Left            =   165
         TabIndex        =   31
         Top             =   1440
         Width           =   1215
      End
      Begin MSForms.ComboBox cmbDT 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   960
         Width           =   3015
         VariousPropertyBits=   1820346395
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5318;556"
         TextColumn      =   2
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;70555"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date:"
         Height          =   195
         Index           =   19
         Left            =   8805
         TabIndex        =   30
         Top             =   165
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   780
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   11535
      End
   End
   Begin VB.TextBox txtInputGrid 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdPreviewDemands 
      Caption         =   "Preview Batch"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Copy from lease"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdBatchClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.OptionButton optSngBatchDemand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Single demand"
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   7845
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optConBatchDemand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consolidate demand"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   7845
      Width           =   1815
   End
   Begin VB.CommandButton cmdBatchDemand 
      Caption         =   "Generate Demands"
      Height          =   375
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Copy from lease"
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "Apply %"
      Height          =   255
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Apply percentage from lease"
      Top             =   2520
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemands 
      Height          =   4425
      Left            =   120
      TabIndex        =   27
      Top             =   2880
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7805
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "%"
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   53
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   41
      Top             =   7440
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   2
      Left            =   120
      Top             =   7800
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee ID"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   26
      Top             =   2520
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   -120
      X2              =   13080
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount"
      Height          =   255
      Index           =   6
      Left            =   10320
      TabIndex        =   23
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usage"
      Height          =   195
      Index           =   4
      Left            =   8280
      TabIndex        =   24
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   21
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      Index           =   1
      X1              =   -120
      X2              =   13080
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "frmCommonUtilityCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iLeft As Integer, iTop As Integer
Private szLeaseID_RC As String, szLeaseID_SC As String, szLeaseID_IC As String
Private szAmt_RC As String, szAmt_SC As String, szAmt_IC As String
Private szLease As String
Private iCurRow As Integer, iCurCol As Integer
Private szIC As String

Private Sub FillCboType(conConnection As ADODB.Connection)
   Dim adoRst     As New ADODB.Recordset
   Dim szSQLStr   As String
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim c          As Integer
   Dim b          As Boolean

   szSQLStr = "SELECT ID, Type, CategoryCode " & _
              "FROM DemandTypes;"
   adoRst.Open szSQLStr, conConnection, adOpenStatic, adLockReadOnly
   
   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
'                                                     Rent Charge
      ReDim Data(1, adoRst.RecordCount) As String

      For i = 0 To adoRst.RecordCount - 1
         b = False
         For j = 0 To 1
            If adoRst.Fields(2).Value = 1 Then
               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
               b = True
            End If
         Next j
         If b Then c = c + 1

         adoRst.MoveNext
         If adoRst.EOF Then Exit For
      Next i

      ReDim Preserve Data(1, c) As String

      cmbDT.Column() = Data()
'
''                                                     Service Charge
'      ReDim Data(1, adoRst.RecordCount) As String
'      adoRst.MoveFirst
'      c = 0
'
'      For i = 0 To adoRst.RecordCount - 1
'         b = False
'         For j = 0 To 1
'            If adoRst.Fields(2).Value = 2 Then
'               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'               b = True
'            End If
'         Next j
'         If b Then c = c + 1
'
'         adoRst.MoveNext
'         If adoRst.EOF Then Exit For
'      Next i
'
'      ReDim Preserve Data(1, c) As String
'
'      cmbDTSC.Column() = Data()
'
''                                                     Insurance Charge
'      ReDim Data(1, adoRst.RecordCount) As String
'      adoRst.MoveFirst
'      c = 0
'
'      For i = 0 To adoRst.RecordCount - 1
'         b = False
'         For j = 0 To 1
'            If adoRst.Fields(2).Value = 3 Then
'               Data(j, c) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'               b = True
'            End If
'         Next j
'         If b Then c = c + 1
'
'         adoRst.MoveNext
'         If adoRst.EOF Then Exit For
'      Next i
'
'      ReDim Preserve Data(1, c) As String
'
'      cmbDTIC.Column() = Data()
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub cboFundIC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboFundSC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cboPropertyList_Click()          'LoadFlxDemands
   Dim adoConn As New ADODB.Connection
   Dim adoRstLeaseDtl As New ADODB.Recordset ', adoRstSplitDemand As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   ConfigureFlxDemands

   adoConn.Open getConnectionString

   szSQLStr = "SELECT L.LeaseID, " & _
                  "L.SageAccountNumber, " & _
                  "L.CompanyName " & _
              "FROM LeaseDetails AS L, Units AS U " & _
              "WHERE L.Status = TRUE AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & cboPropertyList.Column(0) & "' " & _
              "ORDER BY L.SageAccountNumber;"

   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   iRow = 1
   While Not adoRstLeaseDtl.EOF
      flxDemands.TextMatrix(iRow, 0) = adoRstLeaseDtl.Fields.Item("SageAccountNumber").Value
      flxDemands.TextMatrix(iRow, 1) = adoRstLeaseDtl.Fields.Item("CompanyName").Value
      flxDemands.TextMatrix(iRow, 8) = adoRstLeaseDtl.Fields.Item("LeaseID").Value

      iRow = iRow + 1
      adoRstLeaseDtl.MoveNext
      If Not adoRstLeaseDtl.EOF Then flxDemands.AddItem ""
   Wend
   adoRstLeaseDtl.Close

   Set adoRstLeaseDtl = Nothing

   adoConn.Close
   Set adoConn = Nothing
   txtDateIssue.SetFocus
End Sub
'
'Private Sub chkIC_Click()
'   If chkIC.Value = 0 Then
'      cmbDTIC.text = ""
'      cboFundIC.text = ""
'      txtUsage.text = ""
'      chkProIC.Value = 0
'      txtDescIC.text = ""
'   Else
'      Label7.ForeColor = vbBlack
'   End If
'End Sub

Private Sub chkProIC_GotFocus()
   If txtUsage.text = "" Or Val(txtUsage) < 0 Then txtUsage.SetFocus
End Sub

Private Sub chkPro_GotFocus()
'   If txtCurrReading.text = "" Or Val(txtCurrReading.text) < 0 Then txtCurrReading.SetFocus
   If txtCurrReading.text = "" Then txtCurrReading.SetFocus
End Sub

Private Sub chkProSC_GotFocus()
   If txtPrevReading.text = "" Or Val(txtPrevReading.text) < 0 Then txtPrevReading.SetFocus
End Sub

Private Sub GenConBtDmds()
   Dim BRcount As Integer, SCcount As Integer, szSQLStr As String
   Dim iSerial As Integer, lDemand As Long, ICcount As Integer, iVATCode As Integer
   Dim cAmount As Currency, sChargingFig As Single, Msg As String

   Dim adoRstDemandRec As New ADODB.Recordset, adoDmdTypRC As ADODB.Recordset
   Dim adoRstLeaseDtl As New ADODB.Recordset, adoRstSplitDemand As New ADODB.Recordset
   Dim adoDmdTypSC As ADODB.Recordset, adoDmdTypIC As ADODB.Recordset

   If MsgBox("  Are you sure you wish to generate batch demands?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Generate Batch Demands") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

'   Connect to Demands table to add new demands.
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

'   If chkRC.Value = 1 Then
'      Set adoDmdTypRC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDT.Column(0)
'      adoDmdTypRC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If
'   If chkSC.Value = 1 Then
'      Set adoDmdTypSC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDTSC.Column(0)
'      adoDmdTypSC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If
'   If chkIC.Value = 1 Then
'      Set adoDmdTypIC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDTIC.Column(0)
'      adoDmdTypIC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If
'
'   ListOfAllLessee

   szSQLStr = "SELECT L.*, Units.PropertyID " & _
              "FROM LeaseDetails AS L, Units " & _
              "WHERE L.UnitNumber = Units.UnitNumber AND " & _
                    "L.LeaseID IN (" & szLease & ");"
'Debug.Print szSQLStr
   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   iSerial = 1
   
   If adoRstLeaseDtl.EOF Then
      adoRstLeaseDtl.Close
      Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF
         iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn, sChargingFig)

'*********************************************************************************************************
'         Rent Charges Demands
'*********************************************************************************************************
'**** Insert the Header info in the DemandRecPreview table
         lDemand = NextRef(adoConn, "DEMAND_REF")
         With adoRstDemandRec
            .AddNew
            .Fields.Item("DemandID").Value = lDemand
            .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
            .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
            .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
            .Fields.Item("Source").Value = 1
            .Fields.Item("TransactionType").Value = 1
            .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
            .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
            .Fields.Item("IsPrinted").Value = False
            .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoConn)
            .Fields.Item("Spare1").Value = adoDmdTypRC.Fields.Item("spare1").Value
            .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
            .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
            .Update
         End With

         If IsRCSelected(adoRstLeaseDtl!LeaseID, cAmount) Then
            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !SplitID = 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypRC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypRC.Fields.Item("NominalNameforAmount").Value
               !NominalCodeForVAT = adoDmdTypRC.Fields.Item("NominalCodeForVAT").Value
               !NominalNameforVAT = adoDmdTypRC.Fields.Item("NominalNameforVAT").Value
               !NominalCodeForTotal = adoDmdTypRC.Fields.Item("NominalCodeForTotal").Value
               !NominalNameforTotal = adoDmdTypRC.Fields.Item("NominalNameforTotal").Value
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sChargingFig / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
               !SageRef = adoDmdTypRC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !TypeOfDemand = cmbDT.Column(0)
               !description = cmbDT.Column(1)
               !DateFrom = CDate(txtDateFrom.text)
               !DateTo = CDate(txtDateTo.text)
               !SageDepartment = cboFund.Column(0)

               .Update
            End With

            BRcount = BRcount + 1
            iSerial = iSerial + 1
         End If
''************************************************************************************************
''         Service Charge demands
''************************************************************************************************
'         If IsSCSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
'            With adoRstSplitDemand
'               .AddNew
'               !DSR = UniqueID()
'               !SplitID = 1
'               !DemandId = lDemand
'               !A_M = "B"
'               !NominalCodeforAmount = adoDmdTypSC.Fields.Item("NominalCodeforAmount").Value
'               !NominalNameforAmount = adoDmdTypSC.Fields.Item("NominalNameforAmount").Value
'               !NominalCodeForVAT = adoDmdTypSC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypSC.Fields.Item("NominalNameforVAT").Value
'               !NominalCodeForTotal = adoDmdTypSC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypSC.Fields.Item("NominalNameforTotal").Value
'               !Amount = cAmount
'               !VAT_CODE = iVATCode
'               !VATAmount = !Amount * sChargingFig / 100
'               !TotalAmount = CCur(!Amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypSC.Fields.Item("Prefix").Value
'               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
'               !VATMonth = Month(!DueDate)
'               !TypeOfDemand = cmbDTSC.Column(0)
'               !description = cmbDTSC.Column(1)
'               !DateFrom = txtDateFrom.text
'               !DateTo = txtDateTo.text
'
'               !SageDepartment = cboFundSC.Column(0)
'               !ChargingFigure = sChargingFig
'               .Update
'            End With
'
'            SCcount = SCcount + 1
'            iSerial = iSerial + 1
'         End If
'
''************************************************************************************************
''   Insurance Charge demands
''************************************************************************************************
'         If IsICSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
'            With adoRstSplitDemand
'               .AddNew
'               !DSR = UniqueID()
'               !SplitID = 1
'               !DemandId = lDemand
'               !A_M = "B"
'               !NominalCodeforAmount = adoDmdTypIC.Fields.Item("NominalCodeforAmount").Value
'               !NominalNameforAmount = adoDmdTypIC.Fields.Item("NominalNameforAmount").Value
'               !NominalCodeForVAT = adoDmdTypIC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypIC.Fields.Item("NominalNameforVAT").Value
'               !NominalCodeForTotal = adoDmdTypIC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypIC.Fields.Item("NominalNameforTotal").Value
'               !Amount = cAmount
'               !VAT_CODE = iVATCode
'               !VATAmount = !Amount * sChargingFig / 100
'               !TotalAmount = CCur(!Amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypIC.Fields.Item("Prefix").Value
'               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
'               !VATMonth = Month(!DueDate)
'               !TypeOfDemand = cmbDTIC.Column(0)
'               !description = cmbDTIC.Column(1)
'               !DateFrom = txtDateFrom.text
'               !DateTo = txtDateTo.text
'
'               !SageDepartment = cboFundIC.Column(0)
'               !ChargingFigure = sChargingFig
'               .Update
'            End With
'
'            ICcount = ICcount + 1
'            iSerial = iSerial + 1
'         End If

         adoRstLeaseDtl.MoveNext
      Wend

'      If chkRC.Value = 1 Then
'         adoDmdTypRC.Close
'         Set adoDmdTypRC = Nothing
'         Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
'      End If
'      If chkSC.Value = 1 Then
'         adoDmdTypSC.Close
'         Set adoDmdTypSC = Nothing
'         Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
'      End If
'      If chkIC.Value = 1 Then
'         adoDmdTypIC.Close
'         Set adoDmdTypIC = Nothing
'         Msg = Msg & ICcount & " Demands for Insurance Charge were generated." & Chr(13)
'      End If

      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close

      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If

   MousePointer = vbDefault

   Msg = Msg & "A total of " & BRcount + SCcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Batch Demand Generated"

'  Bring all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoConn

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrH:
'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
   MsgBox ERR.Number & " - (pcm_001)" & ERR.description, vbOKOnly, "Error"

   Set adoConn = Nothing
End Sub

Private Sub ListOfLeaseID()
   Dim iRow As Integer

   For iRow = 1 To flxDemands.Rows - 1
      If Val(flxDemands.TextMatrix(iRow, 3)) > 0 Then
         If szLeaseID_RC = "" Then
            szLeaseID_RC = flxDemands.TextMatrix(iRow, 8)
            szAmt_RC = flxDemands.TextMatrix(iRow, 3)
         Else
            szLeaseID_RC = szLeaseID_RC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_RC = szAmt_RC & ", " & flxDemands.TextMatrix(iRow, 3)
         End If
      End If
      
      If Val(flxDemands.TextMatrix(iRow, 5)) > 0 Then
         If szLeaseID_SC = "" Then
            szLeaseID_SC = flxDemands.TextMatrix(iRow, 8)
            szAmt_SC = flxDemands.TextMatrix(iRow, 5)
         Else
            szLeaseID_SC = szLeaseID_SC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_SC = szAmt_SC & ", " & flxDemands.TextMatrix(iRow, 5)
         End If
      End If
      
      If Val(flxDemands.TextMatrix(iRow, 7)) > 0 Then
         If szLeaseID_IC = "" Then
            szLeaseID_IC = flxDemands.TextMatrix(iRow, 8)
            szAmt_IC = flxDemands.TextMatrix(iRow, 7)
         Else
            szLeaseID_IC = szLeaseID_IC & ", " & flxDemands.TextMatrix(iRow, 8)
            szAmt_IC = szAmt_IC & ", " & flxDemands.TextMatrix(iRow, 7)
         End If
      End If
   Next iRow

   szLease = ""
   For iRow = 1 To flxDemands.Rows - 1
      If (Val(flxDemands.TextMatrix(iRow, 3)) > 0 Or _
          Val(flxDemands.TextMatrix(iRow, 5)) > 0 Or _
          Val(flxDemands.TextMatrix(iRow, 7)) > 0) And _
          InStr(szLease, Len(flxDemands.TextMatrix(iRow, 8))) = 0 Then

         szLease = szLease + IIf(szLease = "", "", ", ") + "'" & flxDemands.TextMatrix(iRow, 8) & "'"
      End If
   Next iRow
End Sub
'
'Private Sub chkRC_Click()
'   If chkRC.Value = 0 Then
'      cmbDT.text = ""
'      cboFund.text = ""
'      txtCurrReading.text = ""
'      chkPro.Value = 0
'      txtDesc.text = ""
'   Else
'      Label5.ForeColor = vbBlack
'   End If
'End Sub
'
'Private Sub chkSC_Click()
'   If chkSC.Value = 0 Then
'      cmbDTSC.text = ""
'      cboFundSC.text = ""
'      txtPrevReading.text = ""
'      chkProSC.Value = 0
'      txtTotalCharge.text = ""
'   Else
'      Label6.ForeColor = vbBlack
'   End If
'End Sub
'
'Private Sub cmbDTIC_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   KeyAscii = 0
'End Sub

Private Sub cmbDT_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmbDTSC_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdBatchClose_Click()
   Unload Me
End Sub

Private Sub cmdBatchDemand_Click()
   Dim iRow As Integer, iCol As Integer

   If cmdOK.Enabled Then
      cmdOK.SetFocus
      Exit Sub
   End If

   ListOfLeaseID

   If optSngBatchDemand.Value Then
      GenSngBtDmds
   Else
      GenConBtDmds
   End If

   szLeaseID_RC = ""
   szLeaseID_SC = ""
   szLeaseID_IC = ""

   For iRow = 1 To flxDemands.Rows - 1
      For iCol = 2 To 7
         flxDemands.TextMatrix(iRow, iCol) = ""
      Next iCol
   Next iRow
   cmdCancel_Click
End Sub

Private Sub GenSngBtDmds()
   Dim BRcount As Integer, SCcount As Integer, szSQLStr As String
   Dim iVATCode As Integer
   Dim iSerial As Integer, lDemand As Long, ICcount As Integer
   Dim cAmount As Currency, sChargingFig As Single, Msg As String, sVatRate As Single

   Dim adoRstDemandRec As New ADODB.Recordset, adoDmdTypRC As ADODB.Recordset
   Dim adoRstLeaseDtl As New ADODB.Recordset, adoRstSplitDemand As New ADODB.Recordset
   Dim adoDmdTypSC As ADODB.Recordset, adoDmdTypIC As ADODB.Recordset

   If MsgBox("  Are you sure you wish to generate batch demands?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Generate Batch Demands") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

'   Connect to Demands table to add new demands.
   szSQLStr = "SELECT * FROM DemandRecords"
   adoRstDemandRec.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   szSQLStr = "SELECT * FROM DemandSplitRecords"
   adoRstSplitDemand.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

'   If chkRC.Value = 1 Then
'      Set adoDmdTypRC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDT.Column(0)
'      adoDmdTypRC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If
'   If chkSC.Value = 1 Then
'      Set adoDmdTypSC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDTSC.Column(0)
'      adoDmdTypSC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If
'   If chkIC.Value = 1 Then
'      Set adoDmdTypIC = New ADODB.Recordset
'      szSQLStr = "SELECT * FROM DemandTypes WHERE ID = " & cmbDTIC.Column(0)
'      adoDmdTypIC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'   End If

   szSQLStr = "SELECT LeaseDetails.*, Units.PropertyID " & _
              "FROM LeaseDetails, Units " & _
              "WHERE LeaseDetails.UnitNumber = Units.UnitNumber;"

   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic

   iSerial = 1

   If adoRstLeaseDtl.EOF Then
      adoRstLeaseDtl.Close
      Set adoRstLeaseDtl = Nothing
   Else
      While Not adoRstLeaseDtl.EOF
'*********************************************************************************************************
'         Rent Charges Demands
'*********************************************************************************************************
         If IsRCSelected(adoRstLeaseDtl!LeaseID, cAmount) Then
            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn, sVatRate)

'**** Insert the Header info in the DemandRecPreview table
            lDemand = NextRef(adoConn, "DEMAND_REF")
            With adoRstDemandRec
               .AddNew
               .Fields.Item("DemandID").Value = lDemand
               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
               .Fields.Item("Source").Value = 1
               .Fields.Item("TransactionType").Value = 1
               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
               .Fields.Item("IsPrinted").Value = False
               .Fields.Item("Spare1").Value = adoDmdTypRC.Fields.Item("spare1").Value
               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoConn)
               .Fields.Item("Details").Value = "Rent"
               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Update
            End With

            With adoRstSplitDemand
               .AddNew
               !DSR = UniqueID()
               !SplitID = 1
               !DEMANDID = lDemand
               !A_M = "B"
               !NominalCodeforAmount = adoDmdTypRC.Fields.Item("NominalCodeforAmount").Value
               !NominalNameforAmount = adoDmdTypRC.Fields.Item("NominalNameforAmount").Value
               !NominalCodeForVAT = adoDmdTypRC.Fields.Item("NominalCodeForVAT").Value
               !NominalNameforVAT = adoDmdTypRC.Fields.Item("NominalNameforVAT").Value
               !NominalCodeForTotal = adoDmdTypRC.Fields.Item("NominalCodeForTotal").Value
               !NominalNameforTotal = adoDmdTypRC.Fields.Item("NominalNameforTotal").Value
               !amount = cAmount
               !VAT_CODE = iVATCode
               !VATAmount = !amount * sVatRate / 100
               !TotalAmount = CCur(!amount) + CCur(!VATAmount)
               !SageRef = adoDmdTypRC.Fields.Item("Prefix").Value
               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
               !VATMonth = Month(!DueDate)
               !TypeOfDemand = cmbDT.Column(0)
               !DateFrom = CDate(txtDateFrom.text)
               !DateTo = CDate(txtDateTo.text)
               !SageDepartment = cboFund.Column(0)
               !description = txtDesc.text

               .Update
            End With

            BRcount = BRcount + 1
            iSerial = iSerial + 1
         End If
''************************************************************************************************
''         Service Charge demands
''************************************************************************************************
'         If IsSCSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
'            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn, sVatRate)
''**** Insert the Header info in the DemandRecords table
'            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE NEXT DEMAND ID
'            With adoRstDemandRec
'               .AddNew
'               .Fields.Item("DemandID").Value = lDemand
'               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'               .Fields.Item("Source").Value = 1
'               .Fields.Item("TransactionType").Value = 1
'               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
'               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'               .Fields.Item("IsPrinted").Value = False
'               .Fields.Item("Spare1").Value = adoDmdTypSC.Fields.Item("spare1").Value
'               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoConn)
'               .Fields.Item("Details").Value = "Service Charge"
'               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
'               .Update
'            End With
'
'            With adoRstSplitDemand
'               .AddNew
'               !DSR = UniqueID()
'               !SplitID = 1
'               !DemandId = lDemand
'               !A_M = "B"
'               !NominalCodeforAmount = adoDmdTypSC.Fields.Item("NominalCodeforAmount").Value
'               !NominalNameforAmount = adoDmdTypSC.Fields.Item("NominalNameforAmount").Value
'               !NominalCodeForVAT = adoDmdTypSC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypSC.Fields.Item("NominalNameforVAT").Value
'               !NominalCodeForTotal = adoDmdTypSC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypSC.Fields.Item("NominalNameforTotal").Value
'               !Amount = cAmount
'               !VAT_CODE = iVATCode
'               !VATAmount = !Amount * sVatRate / 100
'               !TotalAmount = CCur(!Amount) + CCur(!VATAmount)
'               !SageRef = adoDmdTypSC.Fields.Item("Prefix").Value
'               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
'               !VATMonth = Month(!DueDate)
'               !TypeOfDemand = cmbDTSC.Column(0)
'               !DateFrom = txtDateFrom.text
'               !DateTo = txtDateTo.text
'               !description = txtTotalCharge.text
'               !SageDepartment = cboFundSC.Column(0)
'               !ChargingFigure = sChargingFig
'               .Update
'            End With
'
'            SCcount = SCcount + 1
'            iSerial = iSerial + 1
'         End If
'
''************************************************************************************************
''   Insurance Charge demands
''************************************************************************************************
'         If IsICSelected(adoRstLeaseDtl!LeaseID, cAmount, sChargingFig) Then
'            iVATCode = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn, sVatRate)
'   '**** Insert the Header info in the DemandRecords table
'            lDemand = NextRef(adoConn, "DEMAND_REF")        'GET THE NEXT DEMAND ID
'            With adoRstDemandRec
'               .AddNew
'               .Fields.Item("DemandID").Value = lDemand
'               .Fields.Item("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'               .Fields.Item("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'               .Fields.Item("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'               .Fields.Item("Source").Value = 1
'               .Fields.Item("TransactionType").Value = 1
'               .Fields.Item("IssueDate").Value = Format(txtDateIssue.text, "dd/mm/yyyy")
'               .Fields.Item("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'               .Fields.Item("IsPrinted").Value = False
'               .Fields.Item("DmdSlNo").Value = SlNumber("SI", "DemandRecords", adoConn)
'               .Fields.Item("Spare1").Value = adoDmdTypIC.Fields.Item("spare1").Value
'               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'               .Fields.Item("Details").Value = "Insurance"
'               .Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
'               .Update
'            End With
'
'            With adoRstSplitDemand
'               .AddNew
'               !DSR = UniqueID()
'               !SplitID = 1
'               !DemandId = lDemand
'               !A_M = "B"
'               !NominalCodeforAmount = adoDmdTypIC.Fields.Item("NominalCodeforAmount").Value
'               !NominalNameforAmount = adoDmdTypIC.Fields.Item("NominalNameforAmount").Value
'               !NominalCodeForVAT = adoDmdTypIC.Fields.Item("NominalCodeForVAT").Value
'               !NominalNameforVAT = adoDmdTypIC.Fields.Item("NominalNameforVAT").Value
'               !NominalCodeForTotal = adoDmdTypIC.Fields.Item("NominalCodeForTotal").Value
'               !NominalNameforTotal = adoDmdTypIC.Fields.Item("NominalNameforTotal").Value
'               !Amount = cAmount
'               !VAT_CODE = iVATCode
'               !VATAmount = !Amount * sVatRate / 100
'               !TotalAmount = CCur(!Amount) + CCur(!VATAmount)
'               !TotalAmount = !Amount
'               !SageRef = adoDmdTypIC.Fields.Item("Prefix").Value
'               !DueDate = Format(txtDateIssue.text, "dd/mm/yyyy")
'               !VATMonth = Month(!DueDate)
'               !TypeOfDemand = cmbDTIC.Column(0)
'               !description = txtDescIC.text
'               !VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn)
'               !DateFrom = txtDateFrom.text
'               !DateTo = txtDateTo.text
'
'               !SageDepartment = cboFundIC.Column(0)
'               !ChargingFigure = sChargingFig
'               .Update
'            End With
'
'            ICcount = ICcount + 1
'            iSerial = iSerial + 1
'         End If

         adoRstLeaseDtl.MoveNext
      Wend

'      If chkRC.Value = 1 Then
'         adoDmdTypRC.Close
'         Set adoDmdTypRC = Nothing
'         Msg = Msg & BRcount & " Demands for Rent were generated." & Chr(13)
'      End If
'      If chkSC.Value = 1 Then
'         adoDmdTypSC.Close
'         Set adoDmdTypSC = Nothing
'         Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
'      End If
'      If chkIC.Value = 1 Then
'         adoDmdTypIC.Close
'         Set adoDmdTypIC = Nothing
'         Msg = Msg & ICcount & " Demands for Insurance Charge were generated." & Chr(13)
'      End If

      adoRstLeaseDtl.Close
      adoRstDemandRec.Close
      adoRstSplitDemand.Close

      Set adoRstLeaseDtl = Nothing
      Set adoRstDemandRec = Nothing
      Set adoRstSplitDemand = Nothing
   End If

   MousePointer = vbDefault

   Msg = Msg & "A total of " & BRcount + SCcount + ICcount & " demands were generated."

   MsgBox Msg, vbOKOnly + vbInformation, "Batch Demand Generated"

'  Bring all Invoices or Demands into tlbReceipt table *********************************************
   MigrateInvIntoReceipt adoConn

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

ErrH:
'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
   MsgBox ERR.Number & " - (pcm_001)" & ERR.description, vbOKOnly, "Error"

   Set adoConn = Nothing
End Sub

Private Function IsRCSelected(ByVal szLeaseID As String, ByRef cAmount As Currency) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String

   szaLeaseID = Split(szLeaseID_RC, ", ")
   szaAmount = Split(szAmt_RC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsRCSelected = True
         cAmount = Val(szaAmount(i))
         Exit Function
      End If
   Next i

   IsRCSelected = False
   cAmount = 0
End Function

Private Function IsSCSelected(ByVal szLeaseID As String, ByRef cAmount As Currency, ByRef sChargingFig As Single) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String, j As Integer

   szaLeaseID = Split(szLeaseID_SC, ", ")
   szaAmount = Split(szAmt_SC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsSCSelected = True
         cAmount = Val(szaAmount(i))
         
         For j = 1 To flxDemands.Rows - 1
            If flxDemands.TextMatrix(j, 8) = szLeaseID Then
               sChargingFig = IIf(IsNull(flxDemands.TextMatrix(j, 4)), Null, Val(flxDemands.TextMatrix(j, 4)))
               Exit For
            End If
         Next j
         
         Exit Function
      End If
   Next i

   IsSCSelected = False
   cAmount = 0
End Function

Private Function IsICSelected(ByVal szLeaseID As String, ByRef cAmount As Currency, ByRef sChargingFig As Single) As Boolean
   Dim szaLeaseID() As String, i As Integer, szaAmount() As String, j As Integer

   szaLeaseID = Split(szLeaseID_IC, ", ")
   szaAmount = Split(szAmt_IC, ", ")

   For i = 0 To UBound(szaLeaseID)
      If szaLeaseID(i) = szLeaseID Then
         IsICSelected = True
         cAmount = Val(szaAmount(i))
         
         For j = 1 To flxDemands.Rows - 1
            If flxDemands.TextMatrix(j, 8) = szLeaseID Then
               sChargingFig = IIf(IsNull(flxDemands.TextMatrix(j, 6)), Null, Val(flxDemands.TextMatrix(j, 6)))
               Exit For
            End If
         Next j
         
         Exit Function
      End If
   Next i

   IsICSelected = False
   cAmount = 0
End Function

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRst.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
'   cboProperty.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub ConfigureFlxDemands()
   Dim szHeader As String, i As Integer

   flxDemands.Clear
   flxDemands.Cols = 9
   flxDemands.Rows = 2

   szHeader$ = "<UnitID|<LesseeID|<Lessee|<Description|>Usage" & _
               "|>Percentage|>Amt||"
   flxDemands.FormatString = szHeader$

   For i = 1 To flxDemands.Cols - 4
      flxDemands.ColWidth(i - 1) = Label3(i).Left - Label3(i - 1).Left
   Next i
   flxDemands.ColWidth(i - 1) = flxDemands.Left + flxDemands.Width - Label3(i - 1).Left - 300
   flxDemands.ColWidth(i) = 0
   flxDemands.ColWidth(i + 1) = 0
   flxDemands.ColWidth(i + 2) = 0

   flxDemands.RowHeight(0) = 0
End Sub

Private Sub cmdCancel_Click()
   cmdOK.Enabled = True
   cmdCancel.Enabled = False
   fraHeader.Enabled = True
'   chkRC.Value = 0
'   chkSC.Value = 0
'   chkIC.Value = 0

   ClearGrid
End Sub

Private Sub ClearGrid()
   Dim iRow As Integer, iCol As Integer

   For iRow = 1 To flxDemands.Rows - 1
      For iCol = 2 To 7
         flxDemands.TextMatrix(iRow, iCol) = ""
      Next iCol
   Next iRow
   CalAllTotal
   szLeaseID_RC = ""
   szLeaseID_SC = ""
   szLeaseID_IC = ""
End Sub

Private Sub cmdClear_Click()
   If MsgBox("Do you want to clear the entries of the grid?", vbQuestion + vbYesNo, "Batch Demands") = vbNo Then Exit Sub

   ClearGrid
End Sub

Private Sub cmdExclude_Click()
   If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Batch Demands") = vbNo Then Exit Sub

   Dim iRow As Integer, iCurRow As Integer, iCurCol As Integer

   iCurRow = flxDemands.row
   iCurCol = flxDemands.col
   
   flxDemands.col = 1
   iRow = flxDemands.Rows - 1
   Do
      flxDemands.row = iRow
      If flxDemands.CellBackColor = RGB(233, 232, 155) Then
         flxDemands.RemoveItem iRow
         If iRow = iCurRow Then iCurRow = 0
      End If

      iRow = iRow - 1
   Loop While iRow > 0

   flxDemands.row = iCurRow
   flxDemands.col = iCurCol
   CalAllTotal
End Sub
'
'Private Sub cmdLookupIC_Click()
'   If cboPropertyList.text = "" Then Exit Sub
'
'   If chkIC.Value = 0 Then
'      MsgBox "Insurance charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
'      Label7.ForeColor = vbRed
'      Exit Sub
'   End If
'
'   If Val(txtUsage.text) < 0 Or txtUsage.text = "" Then
'      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"
'
'      fraHeader.Enabled = True
'      cmdOK.Enabled = True
'      cmdCancel.Enabled = False
'
'      txtUsage.SetFocus
'      Exit Sub
'   End If
'
'   Dim adoConn As New ADODB.Connection
'   Dim adoRstRC As New ADODB.Recordset
'   Dim szSQLStr As String, iRow As Integer
'
'   adoConn.Open getConnectionString
'
'   szSQLStr = "SELECT L.SageAccountNumber, R.ChargingFigure  " & _
'              "FROM LeaseDetails AS L, Units AS U, LInsuranceCharges AS R " & _
'              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.ChargingType = 1 AND " & _
'                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
'                  "L.UnitNumber = U.UnitNumber AND " & _
'                  "U.PropertyID = '" & cboPropertyList.Column(0) & "';"
'
'   adoRstRC.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
''Debug.Print szSQLStr
'   While Not adoRstRC.EOF
'      For iRow = 1 To flxDemands.Rows - 1
'         If flxDemands.TextMatrix(iRow, 0) = adoRstRC.Fields.Item("SageAccountNumber").Value Then
'            flxDemands.TextMatrix(iRow, 6) = adoRstRC.Fields.Item("ChargingFigure").Value
'            flxDemands.TextMatrix(iRow, 7) = _
'                  Format(Val(txtUsage.text) * (Val(flxDemands.TextMatrix(iRow, 6)) / 100), "0.00")
'         End If
'      Next iRow
'      adoRstRC.MoveNext
'   Wend
'   adoRstRC.Close
'   Set adoRstRC = Nothing
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub cmdLookup_Click()
   If cboPropertyList.text = "" Then Exit Sub

'   If chkRC.Value = 0 Then
'      MsgBox "Rent charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
'      Label5.ForeColor = vbRed
'      Exit Sub
'   End If

   If Val(txtCurrReading.text) < 0 Or txtCurrReading.text = "" Then
      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"

      fraHeader.Enabled = True
      cmdOK.Enabled = True
      cmdCancel.Enabled = False

      txtCurrReading.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRstRC As New ADODB.Recordset
   Dim szSQLStr As String, iRow As Integer

   adoConn.Open getConnectionString

   szSQLStr = "SELECT L.SageAccountNumber, R.spare2  " & _
              "FROM LeaseDetails AS L, Units AS U, LRentCharges AS R " & _
              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.spare1 = '2' AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & cboPropertyList.Column(0) & "';"

   adoRstRC.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQLStr
   While Not adoRstRC.EOF
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = adoRstRC.Fields.Item("SageAccountNumber").Value Then
            flxDemands.TextMatrix(iRow, 2) = adoRstRC.Fields.Item("spare2").Value
            flxDemands.TextMatrix(iRow, 3) = _
                  Format(Val(txtCurrReading.text) * (Val(flxDemands.TextMatrix(iRow, 2)) / 100), "0.00")
         End If
      Next iRow
      adoRstRC.MoveNext
   Wend
   adoRstRC.Close
   Set adoRstRC = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Sub
'
'Private Sub cmdLookupSC_Click()
'   If cboPropertyList.text = "" Then Exit Sub
'
'   If chkSC.Value = 0 Then
'      MsgBox "Service charge option is not selected.", vbInformation + vbOKOnly, "Batch Demands"
'      Label6.ForeColor = vbRed
'      Exit Sub
'   End If
'
'   If Val(txtPrevReading.text) < 0 Or txtPrevReading.text = "" Then
'      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"
'
'      fraHeader.Enabled = True
'      cmdOK.Enabled = True
'      cmdCancel.Enabled = False
'
'      txtPrevReading.SetFocus
'      Exit Sub
'   End If
'
'   Dim adoConn As New ADODB.Connection
'   Dim adoRstSC As New ADODB.Recordset
'   Dim szSQLStr As String, iRow As Integer
'
'   adoConn.Open getConnectionString
'
'   szSQLStr = "SELECT L.SageAccountNumber, R.CMFigure  " & _
'              "FROM LeaseDetails AS L, Units AS U, LServiceCharges AS R " & _
'              "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.ChargingMethod = 2 AND " & _
'                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
'                  "R.ServiceChargeDept = '" & CStr(cboFundSC.Column(0)) & "' AND " & _
'                  "L.UnitNumber = U.UnitNumber AND " & _
'                  "U.PropertyID = '" & cboPropertyList.Column(0) & "';"
'
'   adoRstSC.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
''Debug.Print szSQLStr
'   While Not adoRstSC.EOF
'      For iRow = 1 To flxDemands.Rows - 1
'         If flxDemands.TextMatrix(iRow, 0) = adoRstSC.Fields.Item("SageAccountNumber").Value Then
'            flxDemands.TextMatrix(iRow, 4) = adoRstSC.Fields.Item("CMFigure").Value
'            flxDemands.TextMatrix(iRow, 5) = _
'                  Format(Val(txtPrevReading.text) * (Val(flxDemands.TextMatrix(iRow, 4)) / 100), "0.00")
'         End If
'      Next iRow
'      adoRstSC.MoveNext
'   Wend
'   adoRstSC.Close
'   Set adoRstSC = Nothing
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub cmdOK_Click()
   If Not CheckInput Then Exit Sub

   Dim iRow As Integer, dProrata As Double, dPercentage As Double

'   If chkRC.Value = 1 Then
'      If chkPro.Value = 1 Then
'         dProrata = Val(txtCurrReading.text) / (flxDemands.Rows - 1)
'         dPercentage = 100 / (flxDemands.Rows - 1)
'
'         For iRow = 1 To flxDemands.Rows - 1
'            flxDemands.TextMatrix(iRow, 2) = Format(dPercentage, "0.0000")
'            flxDemands.TextMatrix(iRow, 3) = Format(dProrata, "0.00")
'         Next iRow
'      End If
'   End If

   If frmMMain.IsRibbonVersion And lblPostingDate.ToolTipText <> "" Then
      Dim adoConn As New ADODB.Connection
      Dim szSQL As String

      adoConn.Open getConnectionString

'      If IsPeriodStatus(lblPostingDate.ToolTipText, cboClientList.Value, adoConn) <> 1 Then
'         ShowMsgInTaskBar "The transaction date falls within a closed period", "Y", "N"
'         txtDateIssue.SetFocus
'         adoConn.Close
'         Set adoConn = Nothing
'         Exit Sub
'      End If
      adoConn.Close
      Set adoConn = Nothing
   End If

'   If chkSC.Value = 1 Then
'      If chkProSC.Value = 1 Then
''         dProrata = RoundingNumber(Val(txtPrevReading.text) / (flxDemands.Rows - 1), 2)
'         dPercentage = RoundingNumber(100 / (flxDemands.Rows - 1), 4)
'         dProrata = RoundingNumber(Val(txtPrevReading.text) * (dPercentage / 100), 2)
''Debug.Print RoundingNumber(Val(txtPrevReading.text) * (dPercentage / 100), 2)
'         For iRow = 1 To flxDemands.Rows - 1
'            flxDemands.TextMatrix(iRow, 4) = Format(dPercentage, "0.0000")
'            flxDemands.TextMatrix(iRow, 5) = Format(dProrata, "0.00")
'         Next iRow
'      End If
'   End If
'
'   If chkIC.Value = 1 Then
'      If chkProIC.Value = 1 Then
'         dProrata = Val(txtUsage.text) / (flxDemands.Rows - 1)
'         dPercentage = 100 / (flxDemands.Rows - 1)
'
'         For iRow = 1 To flxDemands.Rows - 1
'            flxDemands.TextMatrix(iRow, 6) = Format(dPercentage, "0.0000")
'            flxDemands.TextMatrix(iRow, 7) = Format(dProrata, "0.00")
'         Next iRow
'      End If
'   End If

   fraHeader.Enabled = False
   cmdOK.Enabled = False
   cmdCancel.Enabled = True

   CalAllTotal
End Sub

Private Sub flxDemands_Click()
   If cmdOK.Enabled Then
      cmdOK.SetFocus
      Exit Sub
   End If

   If flxDemands.col = 0 Or flxDemands.col = 1 Then
      UMarkRowFlxGrid flxDemands, flxDemands.row
   End If

   Dim i As Integer, iFlxSPayCol As Integer

   If flxDemands.col < 2 Then Exit Sub
   If flxDemands.TextMatrix(flxDemands.row, 0) = "" Then Exit Sub

'   If (flxDemands.col = 2 Or flxDemands.col = 3) And chkRC.Value = 0 Then
'      MsgBox "Rent Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
'      Exit Sub
'   End If
'   If (flxDemands.col = 2 And (txtCurrReading.text = "" Or Val(txtCurrReading.text) < 0)) And chkRC.Value = 0 Then
'      MsgBox "Please input the budget amount for rent charge.", vbInformation + vbOKOnly, "Batch Payment"
'      txtCurrReading.SetFocus
'      Exit Sub
'   End If
'
'   If (flxDemands.col = 4 Or flxDemands.col = 5) And chkSC.Value = 0 Then
'      MsgBox "Service Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
'      Exit Sub
'   End If
'   If (flxDemands.col = 4 And (txtPrevReading.text = "" Or Val(txtPrevReading.text) < 0)) And chkSC.Value = 0 Then
'      MsgBox "Please input the budget amount for service charge.", vbInformation + vbOKOnly, "Batch Payment"
'      txtPrevReading.SetFocus
'      Exit Sub
'   End If
'
'   If (flxDemands.col = 6 Or flxDemands.col = 7) And chkIC.Value = 0 Then
'      MsgBox "Insurance Charge demand option has not been selected.", vbCritical + vbOKOnly, "Batch Demands"
'      Exit Sub
'   End If
'   If (flxDemands.col = 6 And (txtUsage.text = "" Or Val(txtUsage.text) < 0)) And chkIC.Value = 0 Then
'      MsgBox "Please input the budget amount for insurance charge.", vbInformation + vbOKOnly, "Batch Payment"
'      txtUsage.SetFocus
'      Exit Sub
'   End If

   If flxDemands.col = 2 Or flxDemands.col = 3 Then txtInputGrid.BackColor = Label3(77).BackColor
'   If flxDemands.col = 4 Or flxDemands.col = 5 Then txtInputGrid.BackColor = Label4.BackColor
   If flxDemands.col = 6 Or flxDemands.col = 7 Then txtInputGrid.BackColor = Label3(8).BackColor

   txtInputGrid.Top = flxDemands.CellTop + flxDemands.Top
   txtInputGrid.Left = flxDemands.CellLeft + flxDemands.Left
   txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
   txtInputGrid.Height = flxDemands.RowHeight(flxDemands.row) - 15
   txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
   txtInputGrid.Visible = True
   txtInputGrid.SetFocus
   SelTxtInCtrl txtInputGrid

   iCurRow = flxDemands.row
   iCurCol = flxDemands.col
End Sub

Private Sub flxDemands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Load()
   frmDemands3.Hide

   Me.Height = 8865
   Me.Width = 11865
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fraHeader.BackColor = MODULEBACKCOLOR
'   chkRC.BackColor = MODULEBACKCOLOR
'   chkSC.BackColor = MODULEBACKCOLOR
'   chkIC.BackColor = MODULEBACKCOLOR
   chkPro.BackColor = MODULEBACKCOLOR
'   chkProSC.BackColor = MODULEBACKCOLOR
'   chkProIC.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   PrepareList adoConn, cboClientList, cboPropertyList
   ConfigureFlxDemands

   FillCboType adoConn
   LoadDept adoConn

   adoConn.Close
   Set adoConn = Nothing

   txtDateFrom.text = Format(Now, "dd/mm/yyyy")
   txtDateIssue.text = Format(Now, "dd/mm/yyyy")
   txtDateTo.text = Format(Now, "dd/mm/yyyy")

   txtPcgRC.Left = cmdLookup.Left
   txtPcgRC.Width = flxDemands.ColWidth(2)
   txtAmtRC.Left = txtPcgRC.Left + txtPcgRC.Width + 20
   txtAmtRC.Width = flxDemands.ColWidth(3)

'   txtPcgSC.Left = cmdLookupSC.Left
'   txtPcgSC.Width = flxDemands.ColWidth(4)
'   txtAmtSC.Left = txtPcgSC.Left + txtPcgSC.Width + 20
'   txtAmtSC.Width = flxDemands.ColWidth(5)
'
'   txtPcgIC.Left = cmdLookupIC.Left
'   txtPcgIC.Width = flxDemands.ColWidth(6)
'   txtAmtIC.Left = txtPcgIC.Left + txtPcgIC.Width + 20
'   txtAmtIC.Width = flxDemands.ColWidth(7)

   BATCH_DEMAND_PROCESS = True
   If UCase(SystemUser) <> "SAMRAT" And UCase(WS_Name) <> "WS1" Then
      Call WheelHook(Me.hWnd)
   End If
End Sub

Private Sub CalAllTotal()
   Dim iRow As Integer, cTotal As Currency

   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 2))
   Next iRow
   txtPcgRC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 3))
   Next iRow
   txtAmtRC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 4))
   Next iRow
   txtPcgSC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 5))
   Next iRow
   txtAmtSC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 6))
   Next iRow
   txtPcgIC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 7))
   Next iRow
   txtAmtIC.text = Format(cTotal, "0.00")
End Sub

Private Sub LoadDept(adoConn As ADODB.Connection)
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName FROM Fund;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
      ReDim Data(2, adoRst.RecordCount) As String

      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
         Data(1, rRow) = adoRst.Fields.Item("FundName").Value
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
      cboFund.Clear
      cboFund.Column() = Data()
'      cboFundSC.Clear
'      cboFundSC.Column() = Data()
'      cboFundIC.Clear
'      cboFundIC.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Exit Sub

   ' Error Handling Code
Error_Handler:

   ' Destroy Objects
   Set adoRst = Nothing
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub fraHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDateIssue.text, cboClientList.Value
End Sub
'
'Private Sub txtUsage_GotFocus()
'   If chkIC.Value = 0 Then
'      chkIC.SetFocus
'   Else
'      SelTxtInCtrl txtUsage
'   End If
'End Sub

Private Sub txtUsage_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtUsage, KeyAscii
End Sub

Private Sub txtUsage_LostFocus()
   txtUsage.text = Format(txtUsage.text, "0.00")
End Sub
'
'Private Sub txtCurrReading_GotFocus()
'   If chkRC.Value = 0 Then
'      chkRC.SetFocus
'   Else
'      SelTxtInCtrl txtCurrReading
'   End If
'End Sub

Private Sub txtCurrReading_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtCurrReading, KeyAscii
End Sub

Private Sub txtCurrReading_LostFocus()
   txtUsage.text = Val(txtCurrReading.text) - Val(txtPrevReading.text)
End Sub
'
'Private Sub txtPrevReading_GotFocus()
'   If chkSC.Value = 0 Then
'      chkSC.SetFocus
'   Else
'      SelTxtInCtrl txtPrevReading
'   End If
'End Sub

Private Sub txtPrevReading_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtPrevReading, KeyAscii
End Sub

Private Sub txtPrevReading_LostFocus()
   txtPrevReading.text = Format(txtPrevReading.text, "0.00")
End Sub

Private Sub txtDateFrom_Change()
   TextBoxChangeDate txtDateFrom
End Sub

Private Sub txtDateFrom_GotFocus()
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
   TextBoxFormatDate txtDateFrom
End Sub

Private Sub txtDateIssue_Change()
   TextBoxChangeDate txtDateIssue
End Sub

Private Sub txtDateIssue_GotFocus()
   SelTxtInCtrl txtDateIssue
End Sub

Private Sub txtDateIssue_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateIssue, KeyAscii
End Sub

Private Sub txtDateIssue_LostFocus()
   If txtDateIssue.text <> "" Then
      If TextBoxFormatDate(txtDateIssue) Then
         lblPostingDate.ToolTipText = txtDateIssue.text
      End If
   End If
End Sub

Private Sub txtDateTo_Change()
   TextBoxChangeDate txtDateTo
End Sub

Private Sub txtDateTo_GotFocus()
   SelTxtInCtrl txtDateTo
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateTo, KeyAscii
End Sub

Private Sub txtDateTo_LostFocus()
   TextBoxFormatDate txtDateTo
End Sub
'
'Private Sub txtDescIC_GotFocus()
'   If chkIC.Value = 0 Then chkIC.SetFocus
'End Sub
'
'Private Sub txtDesc_GotFocus()
'   If chkRC.Value = 0 Then chkRC.SetFocus
'End Sub
'
'Private Sub txtTotalCharge_GotFocus()
'   If chkSC.Value = 0 Then chkSC.SetFocus
'End Sub
'
'Private Sub cboFundIC_GotFocus()
'   If chkIC.Value = 0 Then chkIC.Value = 1
'End Sub
'
'Private Sub cboFund_GotFocus()
'   If chkRC.Value = 0 Then chkRC.Value = 1
'End Sub
'
'Private Sub cboFundSC_GotFocus()
'   If chkSC.Value = 0 Then chkSC.Value = 1
'End Sub
'
'Private Sub cmbDTIC_GotFocus()
'   If chkIC.Value = 0 Then chkIC.Value = 1
'End Sub
'
'Private Sub cmbDT_GotFocus()
'   If chkRC.Value = 0 Then chkRC.Value = 1
'End Sub
'
'Private Sub cmbDTSC_GotFocus()
'   If chkSC.Value = 0 Then chkSC.Value = 1
'End Sub

Private Function CheckInput() As Boolean
   CheckInput = False

   If cboPropertyList.text = "" Then
      MsgBox "Please select the property.", vbInformation + vbOKOnly, "Batch Demands"
      cboPropertyList.SetFocus
      Exit Function
   End If
   If txtDateIssue.text = "" Then
      MsgBox "Please issue the from date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateIssue.SetFocus
      Exit Function
   End If
   If txtDateFrom.text = "" Then
      MsgBox "Please input the from date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateFrom.SetFocus
      Exit Function
   End If
   If txtDateTo.text = "" Then
      MsgBox "Please input the to date.", vbInformation + vbOKOnly, "Batch Demands"
      txtDateTo.SetFocus
      Exit Function
   End If

'   If (chkRC.Value = 0) And (chkSC.Value = 0) And (chkIC.Value = 0) Then
'      MsgBox "Please select atleast one charge.", vbInformation + vbOKOnly, "Batch Demands"
'      chkRC.SetFocus
'      Exit Function
'   End If
'   If chkRC.Value And cmbDT.text = "" Then
'      MsgBox "Please select the demand type for rent charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cmbDT.SetFocus
'      Exit Function
'   End If
'   If chkSC.Value And cmbDTSC.text = "" Then
'      MsgBox "Please select the demand type for service charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cmbDTSC.SetFocus
'      Exit Function
'   End If
'   If chkIC.Value And cmbDTIC.text = "" Then
'      MsgBox "Please select the demand type for insurance charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cmbDTIC.SetFocus
'      Exit Function
'   End If
'   If chkRC.Value And cboFund.text = "" Then
'      MsgBox "Please select the fund for rent charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cboFund.SetFocus
'      Exit Function
'   End If
'   If chkSC.Value And cboFundSC.text = "" Then
'      MsgBox "Please select the fund for service charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cboFundSC.SetFocus
'      Exit Function
'   End If
'   If chkIC.Value And cboFundIC.text = "" Then
'      MsgBox "Please select the fund for insurance charge.", vbInformation + vbOKOnly, "Batch Demands"
'      cboFundIC.SetFocus
'      Exit Function
'   End If

   CheckInput = True
End Function

Private Sub txtInputGrid_LostFocus()
   flxDemands.TextMatrix(iCurRow, iCurCol) = txtInputGrid.text

   CalAllTotal
End Sub

Private Sub txtInputGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      flxDemands.TextMatrix(flxDemands.row, flxDemands.col) = txtInputGrid.text

'      InputGridLostFocus

      If MoveDownPosition Then
         txtInputGrid.Left = iLeft
         txtInputGrid.Top = iTop
         txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
         txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
         txtInputGrid.Visible = True
         SelTxtInCtrl txtInputGrid
      End If
   End If

   If KeyCode > 36 And KeyCode < 41 Then
      flxDemands.TextMatrix(flxDemands.row, flxDemands.col) = txtInputGrid.text
'      InputGridLostFocus

      If KeyCode = 38 Then
         If MoveUpPosition Then                     'Up Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 40 Then
         If MoveDownPosition Then                   'Down key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 39 Then
         If MoveRightPosition Then                  'Right Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      If KeyCode = 37 Then
         If MoveLeftPosition Then                  'Left Key
            txtInputGrid.Left = iLeft
            txtInputGrid.Top = iTop
            txtInputGrid.Width = flxDemands.ColWidth(flxDemands.col)
            txtInputGrid.text = flxDemands.TextMatrix(flxDemands.row, flxDemands.col)
            txtInputGrid.Visible = True
         End If
      End If

      SelTxtInCtrl txtInputGrid
   End If

   CalAllTotal
End Sub

Private Function MoveLeftPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.col Mod 2 = 1 Then
      flxDemands.col = flxDemands.col - 1
      iCurCol = flxDemands.col
   Else
      txtInputGrid.Visible = False
      MoveLeftPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveLeftPosition = True
End Function

Private Function MoveRightPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.col Mod 2 = 0 Then
      flxDemands.col = flxDemands.col + 1
      iCurCol = flxDemands.col
   Else
      txtInputGrid.Visible = False
      MoveRightPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveRightPosition = True
End Function

Private Function MoveUpPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.row > 1 Then
      flxDemands.row = flxDemands.row - 1
      iCurRow = flxDemands.row
   Else
      txtInputGrid.Visible = False
      MoveUpPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveUpPosition = True
End Function

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer

   If flxDemands.row < flxDemands.Rows - 1 Then
      flxDemands.row = flxDemands.row + 1
      iCurRow = flxDemands.row
   Else
      txtInputGrid.Visible = False
      MoveDownPosition = False
      Exit Function
   End If

   iLeft = flxDemands.CellLeft + flxDemands.Left
   iTop = flxDemands.CellTop + flxDemands.Top
   MoveDownPosition = True
End Function
'
'Private Sub InputGridLostFocus()
'   If chkRC.Value = 1 And flxDemands.col = 2 And (txtCurrReading.text <> "" And Val(txtCurrReading.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 3) = _
'                 Format(Val(txtCurrReading.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 2)) / 100), "0.00")
'   End If
'   If chkRC.Value = 1 And flxDemands.col = 3 And (txtCurrReading.text <> "" And Val(txtCurrReading.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 2) = _
'                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 3)) / Val(txtCurrReading.text) * 100, "0.0000")
'   End If
'
'   If chkSC.Value = 1 And flxDemands.col = 4 And (txtPrevReading.text <> "" And Val(txtPrevReading.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 5) = _
'                 Format(Val(txtPrevReading.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 4)) / 100), "0.00")
'   End If
'   If chkSC.Value = 1 And flxDemands.col = 5 And (txtPrevReading.text <> "" And Val(txtPrevReading.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 4) = _
'                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 5)) / Val(txtPrevReading.text) * 100, "0.0000")
'   End If
'
'   If chkIC.Value = 1 And flxDemands.col = 6 And (txtUsage.text <> "" And Val(txtUsage.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 7) = _
'                 Format(Val(txtUsage.text) * (Val(flxDemands.TextMatrix(flxDemands.row, 6)) / 100), "0.00")
'   End If
'   If chkIC.Value = 1 And flxDemands.col = 7 And (txtUsage.text <> "" And Val(txtUsage.text) > 0) Then
'      flxDemands.TextMatrix(flxDemands.row, 6) = _
'                 Format(Val(flxDemands.TextMatrix(flxDemands.row, 7)) / Val(txtUsage.text) * 100, "0.00")
'   End If
'End Sub
'
'Public Sub TestingCommand()
'   cboPropertyList.ListIndex = 2
'   chkRC.Value = 1
'   cmbDT.ListIndex = 0
'   cboFund.ListIndex = 0
'End Sub

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
