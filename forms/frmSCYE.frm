VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSCYE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Charge Yearend"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSCYE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBilled 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3300
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1875
      Width           =   1005
   End
   Begin VB.CommandButton cmdGenDmdNow 
      Caption         =   "Generate &Demand Now"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Copy from lease"
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Service Charge Period:"
      Height          =   975
      Left            =   120
      TabIndex        =   40
      Top             =   495
      Width           =   4455
      Begin VB.OptionButton optPeriods 
         Caption         =   "Date Range"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optPeriods 
         Caption         =   "Financial Year"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtDateTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   6
         Top             =   585
         Width           =   1095
      End
      Begin VB.TextBox txtDateFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "01/01/2000"
         Top             =   600
         Width           =   1095
      End
      Begin MSForms.ComboBox cmbFinancialYear 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   2760
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4868;503"
         TextColumn      =   2
         ColumnCount     =   6
         ListRows        =   20
         cColumnInfo     =   6
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0;0;0"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   42
         Top             =   620
         Width           =   180
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   1440
         TabIndex        =   41
         Top             =   620
         Width           =   360
      End
   End
   Begin VB.TextBox txtDueDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1875
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Copy from lease"
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtActualExp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1875
      Width           =   1000
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "C&alculate Service Charge Adjustment"
      Height          =   375
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Generate Payment later"
      Top             =   1095
      Width           =   3135
   End
   Begin VB.TextBox txtGlobalSCBudget 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2310
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1875
      Width           =   1005
   End
   Begin VB.TextBox txtBudAmtSC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4290
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1875
      Width           =   1005
   End
   Begin VB.TextBox txtDescSC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2610
      Width           =   3975
   End
   Begin VB.CommandButton cmdLookupSC 
      Caption         =   "A&pply %"
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Apply percentage from lease"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdExclude 
      Caption         =   "E&xclude"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Copy from lease"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox txtAmtSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox txtPcgSC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSCYearend 
      Caption         =   "&Generate Service Charge Yearend Lease"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Copy from lease"
      Top             =   8640
      Width           =   3615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemands 
      Height          =   4455
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CheckBox chkProSC 
      Appearance      =   0  'Flat
      Caption         =   "&Prorata"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   1880
      Width           =   855
   End
   Begin MSForms.Label lblPostingDate 
      Height          =   285
      Left            =   8280
      TabIndex        =   45
      Top             =   1875
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billed"
      Height          =   195
      Index           =   1
      Left            =   3900
      TabIndex        =   44
      Top             =   1635
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      Height          =   195
      Index           =   8
      Left            =   2805
      TabIndex        =   43
      Top             =   1635
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
      Height          =   195
      Index           =   7
      Left            =   7080
      TabIndex        =   39
      Top             =   1635
      Width           =   705
   End
   Begin VB.Label lblSC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Type:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   2235
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   32
      Top             =   1875
      Width           =   1110
   End
   Begin MSForms.ComboBox cboSCDemandType 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   2235
      Width           =   3975
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7011;503"
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "%"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.ComboBox cboProperty 
      Height          =   285
      Left            =   5355
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5530;503"
      BoundColumn     =   0
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
      Object.Width           =   "1411;-1;0;0"
   End
   Begin MSForms.ComboBox cboClient 
      Height          =   285
      Left            =   675
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6800;503"
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property"
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   34
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
      Height          =   195
      Index           =   2
      Left            =   1870
      TabIndex        =   33
      Top             =   1635
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   9000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment"
      Height          =   195
      Index           =   4
      Left            =   4470
      TabIndex        =   31
      Top             =   1635
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " Amount"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   29
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   26
      Top             =   8160
      Width           =   390
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "SC Fund"
      Height          =   195
      Index           =   9
      Left            =   4680
      TabIndex        =   25
      Top             =   720
      Width           =   585
   End
   Begin MSForms.ComboBox cboFund 
      Height          =   285
      Left            =   5355
      TabIndex        =   7
      Top             =   720
      Width           =   3135
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5530;503"
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   9000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Service Charge"
      Height          =   375
      Left            =   5760
      TabIndex        =   36
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmSCYE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szaProperty()   As String
Private FY_Data()       As String

Private Sub cboClient_Click()
   Dim i As Integer
   Dim c As Integer

   c = 0
   For i = 0 To UBound(szaProperty, 2)
      If cboClient.Column(0) = szaProperty(3, i) Then
         c = c + 1
      End If
   Next i

   Dim szaTempProp() As String
   ReDim szaTempProp(3, c) As String

   c = 0
   For i = 0 To UBound(szaProperty, 2)
      If cboClient.Column(0) = szaProperty(3, i) Then
         szaTempProp(0, c) = szaProperty(0, i)
         szaTempProp(1, c) = szaProperty(1, i)
         szaTempProp(2, c) = szaProperty(2, i)
         szaTempProp(3, c) = szaProperty(3, i)
         c = c + 1
      End If
   Next i

   cboProperty.Column() = szaTempProp()

   LoadCmbFinancialYear
End Sub

Private Sub cmdClose_Click()
   Unload Me
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

Private Sub CalAllTotal()
   Dim iRow As Integer, cTotal As Currency

   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 2))
   Next iRow
   txtPcgSC.text = Format(cTotal, "0.00")
   
   cTotal = 0
   For iRow = 1 To flxDemands.Rows - 1
      cTotal = cTotal + Val(flxDemands.TextMatrix(iRow, 3))
   Next iRow
   txtAmtSC.text = Format(cTotal, "0.00")
End Sub

Private Sub cmdGenDmdNow_Click()
   ShowMsgInTaskBar "Please take a backup of your data", "Y", "P"

   If MsgBox("Have you taken backup of your data?", vbQuestion + vbYesNo, "Service Charge Demand") = vbNo Then Exit Sub

   If txtDescSC.text = "" Then
      MsgBox "Please type a description.", vbCritical + vbOKOnly, "SC Yearend"
      txtDescSC.SetFocus
      Exit Sub
   End If
   If cboSCDemandType.text = "" Then
      MsgBox "Please select demand type.", vbCritical + vbOKOnly, "SC Yearend"
      cboSCDemandType.SetFocus
      Exit Sub
   End If

   Dim iRow As Integer, dProrata As Double, dPercentage As Double

   If chkProSC.Value = 1 Then
      dPercentage = Round(100 / (flxDemands.Rows - 1), 4)
      dProrata = Round(Val(txtBudAmtSC.text) * (dPercentage / 100), 2)
      For iRow = 1 To flxDemands.Rows - 1
         flxDemands.TextMatrix(iRow, 2) = Format(dPercentage, "0.0000")
         flxDemands.TextMatrix(iRow, 3) = Format(dProrata, "0.00")
      Next iRow
   End If

   CalAllTotal
   cmdSCYearend.Enabled = True
   
End Sub

Private Sub cmdLookupSC_Click()
   If txtBudAmtSC.text = "" Then
      MsgBox "Please input the budget amount.", vbInformation + vbOKOnly, "Batch Demands"
      txtBudAmtSC.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRstSC As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT L.SageAccountNumber, R.CMFigure  " & _
           "FROM LeaseDetails AS L, Units AS U, LServiceCharges AS R " & _
           "WHERE L.Status = TRUE AND R.LeaseID = L.LeaseID AND R.ChargingMethod = 2 AND " & _
                "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                 "R.ServiceChargeDept = '" & CStr(cboFund.Column(0)) & "' AND " & _
                 "L.UnitNumber = U.UnitNumber AND " & _
                 "U.PropertyID = '" & cboProperty.Column(0) & "';"

   adoRstSC.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQL
   While Not adoRstSC.EOF
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 0) = adoRstSC.Fields.Item("SageAccountNumber").Value Then
            flxDemands.TextMatrix(iRow, 2) = adoRstSC.Fields.Item("CMFigure").Value
            flxDemands.TextMatrix(iRow, 3) = _
                  Format(Val(txtBudAmtSC.text) * (Val(flxDemands.TextMatrix(iRow, 2)) / 100), "0.00")
         End If
      Next iRow
      adoRstSC.MoveNext
   Wend
   adoRstSC.Close
   Set adoRstSC = Nothing

   adoConn.Close
   Set adoConn = Nothing

   CalAllTotal
End Sub

Private Sub cmdSCYearend_Click()
   Dim adoConn    As New ADODB.Connection
   Dim adoRstSC   As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer

   adoConn.Open getConnectionString

   With adoRstSC
      For iRow = 1 To flxDemands.Rows - 1
         If flxDemands.TextMatrix(iRow, 2) <> "" Then
            szSQL = "SELECT * " & _
                    "FROM LServiceCharges " & _
                    "WHERE SCYE = TRUE AND " & _
                          "LeaseID = '" & flxDemands.TextMatrix(iRow, 4) & "';"
            .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

            If .EOF Then
               .AddNew
               !LeaseID = flxDemands.TextMatrix(iRow, 4)
               !ServiceCharge = UniqueID()
               !SCYE = True
               !SCFrequency = 11                             'Yearly in Advance
               !ChargingMethod = 3                           'Annual
            End If

            !SCDemandType = cboSCDemandType.Column(0)
            !SCPayableFrom = CDate(txtDateFrom.text)
            !SCNextDueDate = txtDueDate.text
            !ServiceChargeDept = cboFund.Column(0)
            !CMFigure = CDbl(flxDemands.TextMatrix(iRow, 2))
            !SCTotal = flxDemands.TextMatrix(iRow, 3)
            !SCAmount = !SCTotal

            !SCDesc = txtDescSC.text
            .Update
            .Close
         End If
      Next iRow
   End With

   adoConn.Close
   Set adoConn = Nothing

   cmdClose.SetFocus
   cmdSCYearend.Enabled = False
   ShowMsgInTaskBar "Service Charge Leases have been generated.", "Y", "P"
End Sub

Private Sub cmdSort_Click()
   Dim szStartDate   As String
   Dim szEndDate     As String

   If cboClient.text = "" Then
      MsgBox "Please select a client from the dropdown list.", vbCritical + vbOKOnly, "SC Yearend"
      cboClient.SetFocus
      Exit Sub
   End If

   If cboProperty.text = "" Then
      MsgBox "Please select a property from the dropdown list.", vbCritical + vbOKOnly, "SC Yearend"
      cboProperty.SetFocus
      Exit Sub
   End If

   If optPeriods(1).Value Then
      If txtDateFrom.text = "" Then
         MsgBox "Please enter the From date.", vbCritical + vbOKOnly, "SC Yearend"
         txtDateFrom.SetFocus
         Exit Sub
      End If

      If txtDateTo.text = "" Then
         MsgBox "Please enter the To date.", vbCritical + vbOKOnly, "SC Yearend"
         txtDateTo.SetFocus
         Exit Sub
      End If
      szStartDate = txtDateFrom.text
      szEndDate = txtDateTo.text
   Else
      If cmbFinancialYear.text = "" Then
         MsgBox "Please select a financial year from the dropdown list.", vbCritical + vbOKOnly, "SC Yearend"
         cmbFinancialYear.SetFocus
         Exit Sub
      End If
      szStartDate = cmbFinancialYear.Column(3)
      szEndDate = cmbFinancialYear.Column(5)
   End If

   If cboFund.text = "" Then
      MsgBox "Please select a fund from the dropdown list.", vbCritical + vbOKOnly, "SC Yearend"
      cboFund.SetFocus
      Exit Sub
   End If

   Dim adoConn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String
   Dim iRow       As Integer
   Dim szaData()  As String
   Dim curPI      As Currency
   Dim curB       As Currency

   adoConn.Open getConnectionString

'~~~~~ SERVICE CHARGE ACTUAL        #####################################################################

   szSQL = "SELECT (IIF(ISNULL(PI_Q.PI), 0, PI_Q.PI) - IIF(ISNULL(PC_Q.PC), 0, PC_Q.PC)) AS A " & _
           "FROM (" & _
               "SELECT SUM(S.TOTAL_AMOUNT) AS PI " & _
               "FROM   tblPurInv AS P INNER JOIN tblPurInvSRec AS S ON P.MY_ID = S.ParentID " & _
               "WHERE  P.PropertyID = '" & cboProperty.Column(0) & "' AND P.TRAN_DATE >= #" & _
                      Format(szStartDate, "dd mmmm yyyy") & "# AND P.TRAN_DATE <= #" & _
                      Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                     "S.DEPT_ID = " & Val(cboFund.Value) & " AND P.TransactionType = 6 " & _
           ") AS PI_Q, (" & _
               "SELECT SUM(S.TOTAL_AMOUNT) AS PC " & _
               "FROM   tblPurInv AS P INNER JOIN tblPurInvSRec AS S ON P.MY_ID = S.ParentID " & _
               "WHERE  P.PropertyID = '" & cboProperty.Column(0) & "' AND P.TRAN_DATE >= #" & _
                     Format(szStartDate, "dd mmmm yyyy") & "# AND P.TRAN_DATE <= #" & _
                     Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                    "S.DEPT_ID = " & Val(cboFund.Value) & " AND P.TransactionType = 7" & _
           ") AS PC_Q;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   curPI = 0
   If Not adoRst.EOF Then
      curPI = IIf(IsNull(adoRst.Fields.Item("A").Value), 0, adoRst.Fields.Item("A").Value)
   End If
   adoRst.Close

   szSQL = "SELECT (IIF(ISNULL(BP_Q.P), 0, BP_Q.P) - IIF(ISNULL(BR_Q.R), 0, BR_Q.R)) AS A " & _
           "FROM (" & _
               "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS P " & _
               "FROM   tlbBankPayment AS BP, NominalLedger AS N " & _
               "WHERE  BP.NOMINAL_CODE = N.Code AND N.ClientID = '" & cboClient.Value & "' AND " & _
                    "N.SubType = 'EX1' AND BP.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
                    "BP.TRAN_DATE >= #" & Format(szStartDate, "dd mmmm yyyy") & "# AND " & _
                    "BP.TRAN_DATE <= #" & Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                    "BP.DEPT_ID = '" & cboFund.Value & "' AND " & _
                    "BP.TransactionType = 11" & _
           ") AS BP_Q, (" & _
               "SELECT SUM(BR.NET_AMOUNT + BR.VAT) AS R " & _
               "FROM   tlbBankPayment AS BR, NominalLedger AS N " & _
               "WHERE  BR.NOMINAL_CODE = N.Code AND N.ClientID = '" & cboClient.Value & "' AND " & _
                    "N.SubType = 'EX1' AND BR.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
                    "BR.TRAN_DATE >= #" & Format(szStartDate, "dd mmmm yyyy") & "# AND " & _
                    "BR.TRAN_DATE <= #" & Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                    "BR.DEPT_ID = '" & cboFund.Value & "' AND " & _
                    "BR.TransactionType = 12) AS BR_Q;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   curB = 0
   If Not adoRst.EOF Then
      curB = IIf(IsNull(adoRst.Fields.Item("A").Value), 0, adoRst.Fields.Item("A").Value)
   End If
   adoRst.Close

   txtActualExp.text = Format(curPI - curB, "0.00")
'------------------------------------------------------------------------------------------------------

'~~~~~~     SERVICE CHARGE BUDGET      ################################################################

   szSQL = "SELECT SUM(G.TotalBudget) AS B " & _
           "FROM   GlobalSC AS G, Fund AS F " & _
           "WHERE  G.Fund = F.FundID AND F.CategoryCode = 2  AND " & _
               "   G.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
               "   F.FundID = " & Val(cboFund.Value) & ";"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      txtGlobalSCBudget.text = Format(adoRst.Fields.Item("B").Value, "0.00")
   Else
      txtGlobalSCBudget.text = "0.00"
   End If
   adoRst.Close
'------------------------------------------------------------------------------------------------------

'~~~~~      BILLED         ############################################################################

   szSQL = "SELECT (IIF(ISNULL(BP_Q.P), 0, BP_Q.P) - IIF(ISNULL(BR_Q.R), 0, BR_Q.R)) AS A " & _
           "FROM (" & _
               "SELECT SUM(S.TotalAmount) AS P " & _
               "FROM   DemandRecords AS D, DemandSplitRecords AS S, Units AS U " & _
               "WHERE  D.DemandID = S.DemandID AND D.UnitNumber = U.UnitNumber AND " & _
                     " U.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
                     " S.SageDepartment = " & Val(cboFund.Value) & " AND " & _
                     " D.IssueDate >= #" & Format(szStartDate, "dd mmmm yyyy") & "# AND " & _
                     " D.IssueDate <= #" & Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                     " D.TransactionType = 1) AS BP_Q, (" & _
               "SELECT SUM(S.TotalAmount) AS R " & _
               "FROM   DemandRecords AS D, DemandSplitRecords AS S, Units AS U, Property AS P " & _
               "WHERE  D.DemandID = S.DemandID AND D.UnitNumber = U.UnitNumber AND " & _
                     " U.PropertyID = '" & cboProperty.Column(0) & "' AND " & _
                     " S.SageDepartment = " & Val(cboFund.Value) & " AND " & _
                     " D.IssueDate >= #" & Format(szStartDate, "dd mmmm yyyy") & "# AND " & _
                     " D.IssueDate <= #" & Format(szEndDate, "dd mmmm yyyy") & "# AND " & _
                     " D.TransactionType = 2) AS BR_Q;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      txtBilled.text = Format(adoRst.Fields.Item("A").Value, "0.00")
   Else
      txtBilled.text = "0.00"
   End If
   adoRst.Close
'------------------------------------------------------------------------------------------------------

'~~~~~      Adjustment         ########################################################################

   txtBudAmtSC.text = Format(Val(txtActualExp.text) - Val(txtGlobalSCBudget.text), "0.00")
'------------------------------------------------------------------------------------------------------

   If Val(txtBudAmtSC.text) < 0 Then
      
   End If

'  Load Lessee
   szSQL = "SELECT L.*, U.PropertyID " & _
              "FROM LeaseDetails AS L, Units AS U " & _
              "WHERE L.Status = TRUE AND " & _
                  "(L.OLED = TRUE OR DATEDIFF('D', NOW, L.ENDDATE) >= 0) AND " & _
                  "L.UnitNumber = U.UnitNumber AND " & _
                  "U.PropertyID = '" & cboProperty.Column(0) & "';"

   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   flxDemands.Rows = 2
   iRow = 1
   While Not adoRst.EOF
      flxDemands.TextMatrix(iRow, 0) = adoRst.Fields.Item("SageAccountNumber").Value
      flxDemands.TextMatrix(iRow, 1) = adoRst.Fields.Item("CompanyName").Value
      flxDemands.TextMatrix(iRow, 4) = adoRst.Fields.Item("LeaseID").Value

      iRow = iRow + 1
      adoRst.MoveNext
      If Not adoRst.EOF Then flxDemands.AddItem ""
   Wend
   adoRst.Close
   
   szSQL = "SELECT ID, Type FROM DemandTypes " & _
             "WHERE (CategoryCode = 2 OR CategoryCode = 5) AND " & _
                   "(PropertyID = '" & cboProperty.Column(0) & "' OR PropertyID = 'ALL');"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iRow = adoRst.RecordCount
   ReDim szaData(1, iRow) As String

   iRow = 0
   If Not adoRst.EOF Then
      While Not adoRst.EOF
         szaData(0, iRow) = adoRst!Id
         szaData(1, iRow) = adoRst!Type
         iRow = iRow + 1
         adoRst.MoveNext
      Wend
   End If
   adoRst.Close
   cboSCDemandType.Clear
   cboSCDemandType.Column() = szaData()

   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   chkProSC.SetFocus
End Sub

Private Sub LoadFlxDemands()
'  Loading flxDemands in cmdSort_Click method which is Calculate button
End Sub

Private Sub flxDemands_Click()
   If flxDemands.col = 0 Or flxDemands.col = 1 Then
      UMarkRowFlxGrid flxDemands, flxDemands.row
   End If
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.Width = 8715
   Me.Height = 9765
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   optPeriods(0).BackColor = MODULEBACKCOLOR
   optPeriods(1).BackColor = MODULEBACKCOLOR
   chkProSC.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   PrepareList adoConn, cboClient, cboProperty
   LoadFund adoConn
   LoadCmbFY_Data adoConn

   adoConn.Close
   Set adoConn = Nothing

   ConfigureFlxDemands
   txtDueDate.text = Format(Now, "dd/mm/yyyy")
   lblPostingDate.ToolTipText = txtDueDate.text
   optPeriods(0).Value = True

'   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigureFlxDemands()
   Dim szHeader As String, i As Integer

   flxDemands.Clear
   flxDemands.Cols = 5
   flxDemands.Rows = 2

   szHeader$ = "<ID|<Lessee|>SC Per|>SC Amt|LeaseID"
   flxDemands.FormatString = szHeader$

   For i = 1 To flxDemands.Cols - 2
      flxDemands.ColWidth(i - 1) = Label3(i).Left - Label3(i - 1).Left
   Next i
   flxDemands.ColWidth(i - 1) = flxDemands.Left + flxDemands.Width - Label3(i - 1).Left - 300
   flxDemands.ColWidth(i) = 0

   flxDemands.RowHeight(0) = 0
   
   txtPcgSC.Left = Label3(2).Left
   txtPcgSC.Width = flxDemands.ColWidth(2)
   txtAmtSC.Left = Label3(3).Left
   txtAmtSC.Width = flxDemands.ColWidth(3)
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

   adoRst.Close
'*************************************** PROPERTY ******************************************
   If cboC.text <> "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProPostCode, ClientID " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboC.Column(0) & "' " & _
              "ORDER BY PropertyID;"
   Else
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProPostCode, ClientID " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   End If
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1

   ReDim szaProperty(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           szaProperty(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cboP.Column() = szaProperty()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadFund(adoConn As ADODB.Connection)
   Dim rRow As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName FROM Fund WHERE CategoryCode = 2;"

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
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Exit Sub

   ' Error Handling Code
Error_Handler:

   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDueDate.text, cboClient.Column(0)
End Sub

Private Sub optPeriods_Click(Index As Integer)
   If optPeriods(0).Value Then
      txtDateFrom.text = ""
      txtDateFrom.Enabled = False
      txtDateTo.Enabled = False
      cmbFinancialYear.Enabled = True
   Else
      txtDateFrom.text = "01/01/2000"
      txtDateFrom.Enabled = True
      txtDateTo.Enabled = True
      cmbFinancialYear.Enabled = False
   End If
End Sub

Private Sub txtDateFrom_LostFocus()
   TextBoxFormatDate txtDateFrom
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

   On Error GoTo ErrHanlder

   If txtDateTo.text = "" Then Exit Sub

   If DateDiff("D", CDate(txtDateFrom.text), CDate(txtDateTo.text)) < 0 Then
      MsgBox "The 'To Date' cannot be before the 'From Date'.", vbCritical + vbOKOnly, "Date Warning"
      txtDateTo.SetFocus
      SelTxtInCtrl txtDateTo
   End If

   Exit Sub

ErrHanlder:
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

Private Sub txtDueDate_Change()
   TextBoxChangeDate txtDueDate
End Sub

Private Sub txtDueDate_GotFocus()
   SelTxtInCtrl txtDueDate
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDueDate, KeyAscii
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

Private Sub LoadCmbFinancialYear()
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer

   K = 0
   For i = 0 To UBound(FY_Data, 2)
      If FY_Data(2, i) = cboClient.Value Then
         K = K + 1
      End If
   Next i
   
   ReDim Data(5, K - 1) As String

   K = 0
   For i = 0 To UBound(FY_Data, 2)
      If FY_Data(2, i) = cboClient.Value Then
         For j = 0 To 5
            Data(j, K) = FY_Data(j, i)
         Next j
         K = K + 1
      End If
   Next i
   cmbFinancialYear.Column() = Data()
   cmbFinancialYear.ListIndex = -1
End Sub

Private Sub LoadCmbFY_Data(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status, FY_EndDate " & _
           "FROM   FinancialYear " & _
           "ORDER BY FY_StDate;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1
   ReDim FY_Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         FY_Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Financial year has not been created for the clients", "Y", "N"
   Exit Sub
End Sub

Private Sub txtDueDate_LostFocus()
   If txtDueDate.text <> "" Then
      lblPostingDate.ToolTipText = txtDueDate.text
   End If
End Sub
