VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLeaseViewSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summary view of Leases Details"
   ClientHeight    =   11505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13605
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeaseViewSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11505
   ScaleWidth      =   13605
   Begin VB.Frame fraBody 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8295
      Left            =   0
      TabIndex        =   15
      Top             =   3120
      Width           =   11535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentCharges 
         Height          =   2445
         Left            =   30
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4313
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9474192
         ForeColorFixed  =   16777215
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSC 
         Height          =   2445
         Left            =   30
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4313
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9474192
         ForeColorFixed  =   16777215
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxIns 
         Height          =   2445
         Left            =   30
         TabIndex        =   18
         Top             =   5760
         Visible         =   0   'False
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4313
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   9474192
         ForeColorFixed  =   16777215
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
      Begin VB.Label lblInsuranceCharges 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Insurance Charges"
         Height          =   195
         Left            =   30
         TabIndex        =   21
         Top             =   5520
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblServiceCharges 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Service Charges"
         Height          =   195
         Left            =   30
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblRC 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rent Charges"
         Height          =   195
         Left            =   30
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   11640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   720
      Top             =   11640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Main"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTenantLookup 
      Height          =   1965
      Left            =   30
      TabIndex        =   0
      Top             =   1050
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
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
   Begin MSForms.TextBox txtSearchLEDate 
      Height          =   315
      Left            =   9090
      TabIndex        =   14
      Top             =   705
      Visible         =   0   'False
      Width           =   1800
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3175;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtSearchUnitName 
      Height          =   315
      Left            =   6120
      TabIndex        =   13
      Top             =   705
      Width           =   2955
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "5221;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTenantSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lease End Date"
      Height          =   195
      Index           =   4
      Left            =   9090
      TabIndex        =   12
      Top             =   480
      Width           =   1080
   End
   Begin MSForms.TextBox txtSearchUnitNum 
      Height          =   315
      Left            =   3840
      TabIndex        =   11
      Top             =   705
      Width           =   2280
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4022;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTenantSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Name"
      Height          =   195
      Index           =   3
      Left            =   6120
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin MSForms.TextBox txtSearchName 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   705
      Width           =   2865
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "5054;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboPropertyList 
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   120
      Width           =   2925
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5159;503"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   3
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboClientList 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   2925
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5159;503"
      BoundColumn     =   0
      TextColumn      =   1
      ColumnCount     =   8
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   13
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   14
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   645
   End
   Begin MSForms.TextBox txtSearchTenant 
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   705
      Width           =   900
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "1587;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblTenantSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Number"
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   495
      Width           =   900
   End
   Begin VB.Label lblTenantSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   495
      Width           =   405
   End
   Begin VB.Label lblTenantSort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee ID"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   495
      Width           =   690
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   30
      Top             =   495
      Width           =   11340
   End
End
Attribute VB_Name = "frmLeaseViewSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OriginalSituation()
   Dim iRow As Integer

   For iRow = 1 To flxRentCharges.Rows - 2
      flxRentCharges.RowHeight(iRow) = 0
   Next iRow
   For iRow = 1 To flxSC.Rows - 2
      flxSC.RowHeight(iRow) = 0
   Next iRow
   For iRow = 1 To flxIns.Rows - 2
      flxIns.RowHeight(iRow) = 0
   Next iRow

   lblServiceCharges.Top = 2760
   lblInsuranceCharges.Top = 5520
   flxSC.Top = 3000
   flxIns.Top = 5760
   flxRentCharges.Height = 2445
   flxSC.Height = 2445
   flxIns.Height = 2445
End Sub

Private Sub cboClientList_Change()
'Resolved by BOSL
'Issue No: 0000467
'Added By: Asif. 04 Sep 2014
If cboClientList.ListCount > 0 Then
    LoadPropertyDropDown
End If
''''''''''''''''''''''
End Sub

Private Sub cboPropertyList_Change()
    'Resolved by BOSL
    'Issue No: 0000467
    'Should be filtered according to the property selected
    'Added By: Asif. 19 Sep 2014
    
    If (cboPropertyList.ListCount > 0) Then
         Dim adoConn As New ADODB.Connection
         Dim szSQL As String
         adoConn.Open getConnectionString
         
         szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitNumber, " & _
                       "IQ.UnitName, IQ.EndDate, IQ.PropertyID " & _
             "FROM Tenants AS T LEFT JOIN " & _
                 "[" & _
                 "SELECT U.UnitNumber, U.UnitName, U.PropertyID, L.SageAccountNumber, " & _
                       "L.EndDate " & _
                 "From Units AS U INNER JOIN LeaseDetails AS L ON " & _
                       "U.UnitNumber = L.UnitNumber " & _
                 "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
             "WHERE IQ.PropertyID = '" & cboPropertyList.Column(0) & "' AND (ISNULL(T.Comments) OR T.Comments='') " & _
             "ORDER BY TenantID;"
         
'         Debug.Print szSQL
         PopulateTenantLookup szSQL, adoConn
         
         'Restore to original position
         Me.Height = fraBody.Top + 400
         Me.Width = 11865
        
        adoConn.Close
        Set adoConn = Nothing
   End If
''''''''''''''''''''''''' end of modification'''''''''''''''''
End Sub

Private Sub flxTenantLookup_RowColChange()
   Dim iRow As Integer, i As Integer, j As Integer, K As Integer

   lblRC.Visible = False
   flxRentCharges.Visible = False
   lblServiceCharges.Visible = False
   flxSC.Visible = False
   lblInsuranceCharges.Visible = False
   flxIns.Visible = False

   OriginalSituation

'  Count Rent Charges in the rent charge grid
   i = 0
   For iRow = 1 To flxRentCharges.Rows - 1
      If flxRentCharges.TextMatrix(iRow, 0) = flxTenantLookup.TextMatrix(flxTenantLookup.row, 0) Then
         i = i + 1
         flxRentCharges.RowHeight(iRow) = 240
      End If
   Next iRow
'  Count Service Charges in the service charge grid
   j = 0
   For iRow = 1 To flxSC.Rows - 1
      If flxSC.TextMatrix(iRow, 0) = flxTenantLookup.TextMatrix(flxTenantLookup.row, 0) Then
         j = j + 1
         flxSC.RowHeight(iRow) = 240
      End If
   Next iRow
'  Count Insurance Charges in the insurance charge grid
   K = 0
   For iRow = 1 To flxIns.Rows - 1
      If flxIns.TextMatrix(iRow, 0) = flxTenantLookup.TextMatrix(flxTenantLookup.row, 0) Then
         K = K + 1
         flxIns.RowHeight(iRow) = 240
      End If
   Next iRow

   If i > 0 Then
      lblRC.Visible = True
      flxRentCharges.Visible = True

      If (j > 0 And K = 0) Or (j = 0 And K > 0) Then
         flxRentCharges.Height = 3765
      End If
      If j = 0 And K = 0 Then
         flxRentCharges.Height = 8085
      End If
      Me.Height = 11985
   End If

   If j > 0 Then
      lblServiceCharges.Visible = True
      flxSC.Visible = True

      If i = 0 And K = 0 Then
         lblServiceCharges.Top = lblRC.Top
         flxSC.Top = flxRentCharges.Top
         flxSC.Height = 8085
      End If
      If i = 0 And K > 0 Then
         lblServiceCharges.Top = lblRC.Top
         flxSC.Top = flxRentCharges.Top
         flxSC.Height = 3765
      End If
      If i > 0 And K = 0 Then
         lblServiceCharges.Top = flxRentCharges.Top + flxRentCharges.Height + 40
         flxSC.Top = lblServiceCharges.Top + 200
         flxSC.Height = 3765
      End If
      Me.Height = 11985
   End If

   If K > 0 Then
      lblInsuranceCharges.Visible = True
      flxIns.Visible = True

      If i = 0 And j = 0 Then
         lblInsuranceCharges.Top = lblRC.Top
         flxIns.Top = flxRentCharges.Top
         flxIns.Height = 8085
      End If
      If i = 0 And j > 0 Then
         lblInsuranceCharges.Top = flxSC.Top + flxSC.Height + 40
         flxIns.Top = lblInsuranceCharges.Top + 200
         flxIns.Height = 3765
      End If
      If i > 0 And j = 0 Then
         lblInsuranceCharges.Top = flxRentCharges.Top + flxRentCharges.Height + 40
         flxIns.Top = lblInsuranceCharges.Top + 200
         flxIns.Height = 3765
      End If
      Me.Height = 11985
   End If

'   If i = 0 And j = 0 And K = 0 Then Me.Height = fraBody.Top + 550
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.Height = fraBody.Top + 550
   Me.Width = 11865
   Me.BackColor = MODULEBACKCOLOR
   fraBody.BackColor = MODULEBACKCOLOR

   Dim szHeader As String, szSQL As String
   Dim adoConn As New ADODB.Connection

   szHeader$ = "<SageAccountNumber|<Type|<Frequency|<BRNextDueDate|>BRTotal|>BRAmount"
   ConfigurFlxCharges flxRentCharges, szHeader
   szHeader$ = "<SageAccountNumber|<Type|<Frequency|<SCNextDueDate|>SCTotal|>SCAmount"
   ConfigurFlxCharges flxSC, szHeader
   szHeader$ = "<SageAccountNumber|<Type|<Frequency|<InsuranceNextDueDate|>TotalYearlyInsurance|>ChargingFigure"
   ConfigurFlxCharges flxIns, szHeader

   adoConn.Open getConnectionString

   PrepareList adoConn             'prepare the list of clients and properties in dwopdown comboes

'Resolved by BOSL
'Issue No: 0000467
'Should be filtered according to the property selected
'Modified By: Asif. 19 Sep 2014

'   szSQL = "SELECT T.SageAccountNumber, T.CompanyName, IQ.UnitNumber, " & _
'                     "IQ.UnitName, IQ.EndDate " & _
'           "FROM Tenants AS T LEFT JOIN " & _
'               "[" & _
'               "SELECT U.UnitNumber, U.UnitName, L.SageAccountNumber, " & _
'                     "L.EndDate " & _
'               "From Units AS U INNER JOIN LeaseDetails AS L ON " & _
'                     "U.UnitNumber = L.UnitNumber " & _
'               "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'           "WHERE ISNULL(T.Comments) OR T.Comments='' " & _
'           "ORDER BY TenantID;"
'
'   Debug.Print szSQL
   
'   PopulateTenantLookup szSQL, adoConn

   flxTenantLookup.row = 0
   flxTenantLookup.col = 0

'  Load Rent Charges
   szSQL = "SELECT L.SageAccountNumber, D.Type, F.Frequency, " & _
               "R.BRNextDueDate, R.BRTotal, R.BRAmount " & _
           "FROM LeaseDetails AS L, LRentCharges AS R, " & _
               "DemandTypes AS D, Frequencies AS F " & _
           "WHERE L.LeaseID = R.LeaseID AND R.BRDemandType = D.ID AND " & _
               "R.BRFrequency = F.ID"
   PopulateCharges flxRentCharges, szSQL, adoConn

'  Load Service Charges
   szSQL = "SELECT L.SageAccountNumber, D.Type, F.Frequency, " & _
               "S.SCNextDueDate, S.SCTotal, S.SCAmount " & _
           "FROM LeaseDetails AS L, LServiceCharges AS S, " & _
               "DemandTypes AS D, Frequencies AS F " & _
           "WHERE L.LeaseID = S.LeaseID AND S.SCDemandType = D.ID AND " & _
               "S.SCFrequency = F.ID"
   PopulateCharges flxSC, szSQL, adoConn

'  Load Insurance Charges
   szSQL = "SELECT L.SageAccountNumber, D.Type, F.Frequency, " & _
               "I.InsuranceNextDueDate, I.TotalYearlyInsurance, I.ChargingFigure " & _
           "FROM LeaseDetails AS L, LInsuranceCharges AS I, " & _
               "DemandTypes AS D, Frequencies AS F " & _
           "WHERE L.LeaseID = I.LeaseID AND I.InsuranceDemandType = D.ID AND " & _
               "I.InsuranceFrequency = F.ID"
   PopulateCharges flxIns, szSQL, adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub PopulateCharges(flxControl As MSHFlexGrid, szSQL As String, adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, r As Integer
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 1
   While Not adoRst.EOF
      For i = 0 To flxControl.Cols - 1
         flxControl.TextMatrix(r, i) = adoRst.Fields.Item(flxControl.TextMatrix(0, i)).Value
      Next i
      flxControl.RowHeight(r) = 0
      r = r + 1
      flxControl.AddItem ""
      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub

'Resolved by BOSL
'Issue No: 0000467
'Added By: Asif. 04 Sep 2014

Private Sub LoadPropertyDropDown()

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler
   
   adoConn.Open getConnectionString
   
   cboPropertyList.Clear
'*************************************** PROPERTY ******************************************
   If cboClientList.Column(0) <> "ALL" Then
   
        szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE Property.ClientID = '" & cboClientList.Column(0) & "' " & _
           "ORDER BY PropertyID;"
   Else
    
        szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"
           
   End If
   
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboPropertyList.Column() = Data()
   
   If cboPropertyList.ListCount > 0 Then
        cboPropertyList.ListIndex = 0
   End If
   
NoRes:
   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT ********************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   Dim Data() As String

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
   cboClientList.Column() = Data()
   cboClientList.ListIndex = 0
   
'Resolved by BOSL
'Issue No: 0000467
'Loading property list not required here. Property list will be loaded when a client is selected.
'Modified By: Asif. 04 Sep 2014

'   adoRst.Close
''*************************************** PROPERTY ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "ORDER BY PropertyID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Properties"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboPropertyList.Column() = Data()
'   cboPropertyList.ListIndex = 0
NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Public Function PopulateTenantLookup(ByVal sSQLQuery_ As String, adoConn As ADODB.Connection)
   Dim iRow As Integer

   iRow = 1

   ConfigurFlexGrid
   populateGrid adoConn, sSQLQuery_, flxTenantLookup
End Function

Private Sub ConfigurFlxCharges(flxControl As MSHFlexGrid, szHd As String)
   Dim i As Integer

   flxControl.Rows = 2
   flxControl.Cols = 6
   flxControl.FormatString = szHd$
   flxControl.ColWidth(0) = 0

   For i = 1 To flxControl.Cols - 1
      flxControl.ColWidth(i) = flxControl.Width / 5 - 60
   Next i
End Sub

Private Sub ConfigurFlexGrid()
   Dim szHeader As String

'Resolved by BOSL
'Issue No: 0000467
'Added a column property in the grid
'Modified By: Asif. 19 Sep 2014
   
   szHeader$ = "<SageAccountNumber|<CompanyName|<UnitNumber|<UnitName|<EndDate"

   flxTenantLookup.Clear
   flxTenantLookup.Rows = 2
   'flxTenantLookup.Cols = 5
   flxTenantLookup.Cols = 6
'''''''''''''End of Modification

   flxTenantLookup.FormatString = szHeader$
   flxTenantLookup.RowHeight(0) = 350
   flxTenantLookup.row = 0

   Dim i As Integer

   For i = 0 To flxTenantLookup.Cols - 1
        flxTenantLookup.col = i
        flxTenantLookup.CellFontBold = True
   Next i

   flxTenantLookup.ColWidth(0) = txtSearchTenant.Width
   flxTenantLookup.ColWidth(1) = txtSearchName.Width
   flxTenantLookup.ColWidth(2) = txtSearchUnitNum.Width
   flxTenantLookup.ColWidth(3) = txtSearchUnitName.Width
   flxTenantLookup.ColWidth(4) = txtSearchLEDate.Width
   
'Resolved by BOSL
'Issue No: 0000467
'Setting width of column property in the grid
'Modified By: Asif. 19 Sep 2014
   flxTenantLookup.ColWidth(5) = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtSearchName_Change()
   Dim i As Integer

   If Len(txtSearchName.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchUnitNum.text = ""
      txtSearchUnitName.text = ""
   End If

   For i = 1 To flxTenantLookup.Rows - 1
      flxTenantLookup.RowHeight(i) = 240
      If UCase(Left(flxTenantLookup.TextMatrix(i, 1), Len(txtSearchName.text))) <> UCase(txtSearchName.text) Then
         flxTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchName_GotFocus()
   SelTxtInCtrl txtSearchName
End Sub

Private Sub txtSearchTenant_Change()
   Dim i As Integer

   If Len(txtSearchTenant.text) > 0 Then
      txtSearchName.text = ""
      txtSearchUnitNum.text = ""
      txtSearchUnitName.text = ""
   End If

   For i = 1 To flxTenantLookup.Rows - 1
      flxTenantLookup.RowHeight(i) = 240
      If UCase(Left(flxTenantLookup.TextMatrix(i, 0), Len(txtSearchTenant.text))) <> UCase(txtSearchTenant.text) Then
         flxTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchTenant_GotFocus()
   SelTxtInCtrl txtSearchTenant
End Sub

Private Sub txtSearchUnitName_Change()
   Dim i As Integer

   If Len(txtSearchUnitNum.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
      txtSearchUnitNum.text = ""
   End If

   For i = 1 To flxTenantLookup.Rows - 1
      flxTenantLookup.RowHeight(i) = 240
      If UCase(Left(flxTenantLookup.TextMatrix(i, 3), Len(txtSearchUnitName.text))) <> UCase(txtSearchUnitName.text) Then
         flxTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchUnitName_GotFocus()
   SelTxtInCtrl txtSearchUnitName
End Sub

Private Sub txtSearchUnitNum_Change()
   Dim i As Integer

   If Len(txtSearchUnitNum.text) > 0 Then
      txtSearchTenant.text = ""
      txtSearchName.text = ""
      txtSearchUnitName.text = ""
   End If

   For i = 1 To flxTenantLookup.Rows - 1
      flxTenantLookup.RowHeight(i) = 240
      If UCase(Left(flxTenantLookup.TextMatrix(i, 2), Len(txtSearchUnitNum.text))) <> UCase(txtSearchUnitNum.text) Then
         flxTenantLookup.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtSearchUnitNum_GotFocus()
   SelTxtInCtrl txtSearchUnitNum
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
