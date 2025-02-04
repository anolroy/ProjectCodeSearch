VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmChangeLesseeID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Lessee ID"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "frmChangeLesseeID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9120
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   9405
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   13
      Top             =   495
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
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3525
         Left            =   45
         TabIndex        =   15
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6218
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   21
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   20
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   18
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7590
      Left            =   80
      TabIndex        =   5
      Top             =   0
      Width           =   8865
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
         Left            =   7695
         TabIndex        =   23
         Top             =   165
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
         Left            =   4050
         TabIndex        =   22
         Top             =   135
         Width           =   300
      End
      Begin VB.Frame fraGrid 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7335
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   8700
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTenantLookup 
            Height          =   6765
            Left            =   0
            TabIndex        =   9
            Top             =   225
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   11933
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   13553358
            ForeColorFixed  =   16777215
            BackColorSel    =   12648447
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   12
            Top             =   0
            Width           =   285
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   2
            Left            =   1050
            TabIndex        =   11
            Top             =   0
            Width           =   405
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   3
            Left            =   3690
            TabIndex        =   10
            Top             =   0
            Width           =   585
         End
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   675
         TabIndex        =   25
         Top             =   135
         Width           =   3330
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5874;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   5265
         TabIndex        =   24
         Top             =   135
         Width           =   2430
         VariousPropertyBits=   746604571
         Size            =   "4286;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   1
         Left            =   4515
         TabIndex        =   7
         Top             =   165
         Width           =   645
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   6
         Top             =   165
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7065
      TabIndex        =   3
      Top             =   7875
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   7875
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5715
      TabIndex        =   2
      Top             =   7875
      Width           =   1215
   End
   Begin MSForms.TextBox txtTenantID 
      Height          =   315
      Left            =   855
      TabIndex        =   0
      Top             =   7830
      Width           =   2985
      VariousPropertyBits=   746604575
      BackColor       =   15858158
      MaxLength       =   20
      Size            =   "5265;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblIDName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   7830
      Width           =   615
   End
End
Attribute VB_Name = "frmChangeLesseeID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WhoAmI As String
Dim sTextBox As String
'Private TENANTID As String
'Keyword for this form Change lease ID
Private Sub PrepareLesseeList()
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientList.Column() = Data()
'   cboClientList.ListIndex = 0
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
'NoRes:
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
End Sub

Private Sub cboClientList_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

'   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

'   If txtClientList.Tag <> "ALL" Then
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'              "ORDER BY PropertyID;"
'   Else
'      szSQL = "SELECT PropertyID, PropertyName, " & _
'                  "ProAddressLine1, ProPostCode " & _
'              "FROM Property " & _
'              "ORDER BY PropertyID;"
'   End If
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count

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

   If WhoAmI = "LesseeID" Then                     '  Current tenants Only
      szSQL = "SELECT " _
            & " Tenants.TenantID, Name, " _
            & " iif(isnull(HOAddressLine1), '', HOAddressLine1) + ' ' + " _
            & " iif(isnull(HOAddressLine2), '', HOAddressLine2) + ' ' + " _
            & " iif(isnull(HOAddressLine3), '', HOAddressLine3) as Address, " _
            & " HOPostCode , HOTelephone, iif(isnull(Comments),'CURRENT','DELETED') as Notes " _
            & " From Tenants, LeaseDetails, Units, Property "
      szSQL = szSQL & _
              " WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " _
                 & "LeaseDetails.UnitNumber = Units.UnitNumber AND " _
                 & "Units.PropertyID = Property.PropertyID AND " _
                 & " " _
                 & "Property.ClientID = '" & txtClientList.Tag & "' " & _
                 "ORDER BY TenantID;"
   End If
   'LeaseDetails.Status = TRUE AND has been removed by anol 20230728
   If WhoAmI = "UnitID" Then
      szSQL = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
        "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE " & _
        "FROM Units, Property " & _
        "WHERE Units.PropertyID = Property.PropertyID AND " & _
        "Property.ClientID = '" & txtClientList.Tag & "' "
   End If
   If WhoAmI = "PropertyID" Then
            If txtClientList.Tag <> "ALL" Then
                szSQL = "SELECT PropertyID, PropertyName, " & _
                    "ProAddressLine1, ProPostCode " & _
                "FROM Property " & _
                "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                "ORDER BY PropertyID;"
            Else
                szSQL = "SELECT PropertyID, PropertyName, " & _
                    "ProAddressLine1, ProPostCode " & _
                "FROM Property " & _
                "ORDER BY PropertyID;"
            End If
   End If
   If szSQL = "" Then Exit Sub
   PopulateTenantPropertyWise szSQL, adoConn

'**************************************************************************************************

NoRes:
   'adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   'adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Public Function PopulateTenantPropertyWise(ByVal sSQLQuery_ As String, ByVal adoConn As ADODB.Connection)
   Dim iRow As Integer
   iRow = 1

   gridTenantLookup.Clear
   gridTenantLookup.Rows = 2
   gridTenantLookup.Cols = 6
   ConfigurFlexGrid

   populateGrid adoConn, sSQLQuery_, gridTenantLookup
   gridTenantLookup.RowHeight(0) = 350
End Function

Private Sub ConfigurFlexGrid()
   gridTenantLookup.RowHeight(0) = 350
   gridTenantLookup.row = 0
   Dim i As Integer
   For i = 0 To gridTenantLookup.Cols - 1
        gridTenantLookup.col = i
        gridTenantLookup.CellFontBold = True
   Next i

    gridTenantLookup.ColWidth(0) = 1000
    gridTenantLookup.ColAlignment(0) = vbLeftJustify
    gridTenantLookup.ColAlignment(1) = vbLeftJustify
    gridTenantLookup.ColAlignment(2) = vbLeftJustify
    gridTenantLookup.ColAlignment(3) = vbLeftJustify
    gridTenantLookup.ColAlignment(5) = vbLeftJustify
   
    gridTenantLookup.TextMatrix(0, 0) = "Sage A/C"
    
    gridTenantLookup.ColWidth(1) = 2900
    gridTenantLookup.TextMatrix(0, 1) = "Name"
    
    gridTenantLookup.ColWidth(2) = 2500
    gridTenantLookup.TextMatrix(0, 2) = "Address"
    
    gridTenantLookup.ColWidth(3) = 1100
    gridTenantLookup.TextMatrix(0, 3) = "Post Code"
    
    gridTenantLookup.ColWidth(4) = 600
    gridTenantLookup.TextMatrix(0, 4) = "Telephone"
    
    gridTenantLookup.ColWidth(5) = 0
    gridTenantLookup.TextMatrix(0, 5) = "Note"
End Sub

Private Sub cboPropertyList_Click()
   Dim adoConn As New ADODB.Connection
   Dim sSQLQuery_ As String

   adoConn.Open getConnectionString

   If WhoAmI = "LesseeID" Then                     '  Current tenants Only
      sSQLQuery_ = "SELECT " _
            & "Tenants.TenantID, Name, " _
            & "iif(isnull(HOAddressLine1),'',HOAddressLine1) + ' ' + " _
            & "iif(isnull(HOAddressLine2),'',HOAddressLine2) + ' ' + " _
            & "iif(isnull(HOAddressLine3),'',HOAddressLine3) as Address, " _
            & "HOPostCode , HOTelephone, iif(isnull(Comments),'CURRENT','DELETED') as Notes " _
            & "From Tenants, LeaseDetails, Units, Property "
      sSQLQuery_ = sSQLQuery_ & _
              "WHERE Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " _
                 & "LeaseDetails.UnitNumber = Units.UnitNumber AND " _
                 & "Units.PropertyID = Property.PropertyID AND " _
                 & "LeaseDetails.Status = TRUE "
      If txtClientList.Tag <> "ALL" Then
         If txtPropertyName.Tag <> "ALL" Then
            sSQLQuery_ = sSQLQuery_ & "AND Property.PropertyID = '" & txtPropertyName.Tag & "' " & _
                        "AND Property.ClientID = '" & txtClientList.Tag & "' " & _
                     "ORDER BY TenantID;"
         Else
            sSQLQuery_ = sSQLQuery_ & _
                        "AND Property.ClientID = '" & txtClientList.Tag & "' " & _
                     "ORDER BY TenantID;"
         End If
      Else
         If txtPropertyName.Tag <> "ALL" Then
            sSQLQuery_ = sSQLQuery_ & "AND Property.PropertyID = '" & txtPropertyName.Tag & "' " & _
                     "ORDER BY TenantID;"
         Else
            sSQLQuery_ = sSQLQuery_ & " ORDER BY TenantID;"
         End If
      End If
   End If
   If WhoAmI = "UnitID" Then
      If txtPropertyName.Tag = "ALL" Then
         If txtClientList.Tag = "ALL" Then
            sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
                           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE " & _
                         "FROM UNITS;"
         Else
            sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
                              "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE " & _
                         "FROM Units, Property " & _
                         "WHERE Units.PropertyID = Property.PropertyID AND " & _
                             "Property.ClientID = '" & txtClientList.Tag & "' "
         End If
      Else
         sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
                           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE " & _
                      "FROM Units " & _
                      "WHERE Units.PropertyID = '" & txtPropertyName.Tag & "' "
      End If
   End If
    If WhoAmI = "SupplierID" Then
      
         If txtClientList.Tag = "ALL" Then
            sSQLQuery_ = "SELECT " _
                & "SupplierID, SupplierName, " _
                & "iif(isnull(SupplierAddressLine1),'',SupplierAddressLine1) + ' ' + " _
                & "iif(isnull(SupplierAddressLine2),'',SupplierAddressLine2) + ' ' + " _
                & "iif(isnull(SupplierAddressLine3),'',SupplierAddressLine3) + ' ' + " _
                & "iif(isnull(SupplierAddressLine4),'',SupplierAddressLine4) as Address, " _
                & "SupplierPostCode , SupplierOfficeTel " _
                & "From Supplier Where TYPE='SUPPLIER' " _
                & "ORDER BY SupplierID;"
         Else
              sSQLQuery_ = "SELECT " _
                & "SupplierID, SupplierName, " _
                & "iif(isnull(SupplierAddressLine1),'',SupplierAddressLine1) + ' ' + " _
                & "iif(isnull(SupplierAddressLine2),'',SupplierAddressLine2) + ' ' + " _
                & "iif(isnull(SupplierAddressLine3),'',SupplierAddressLine3) + ' ' + " _
                & "iif(isnull(SupplierAddressLine4),'',SupplierAddressLine4) as Address, " _
                & "SupplierPostCode , SupplierOfficeTel " _
                & "From Supplier Where TYPE='SUPPLIER' AND SupplierID= '" & txtClientList.Tag & "'" _
                & "ORDER BY SupplierID;"
         End If
      
   End If
'Debug.Print sSQLQuery_
   If sSQLQuery_ <> "" Then
        PopulateTenantPropertyWise sSQLQuery_, adoConn
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub


Private Sub cmdClientList_Click()
    If WhoAmI = "SupplierID" Then
        sTextBox = "3"
        picClient.Left = 915
        picClient.Top = 70
        picClient.Visible = True
        LoadflxSupplier
        
        fraGrid.Enabled = False
        txtSearchClientID.SetFocus
    Else
        sTextBox = "1"
        picClient.Left = 915
        picClient.Top = 70
        picClient.Visible = True
        LoadflxClient
        
        fraGrid.Enabled = False
        txtSearchClientID.SetFocus
     End If
End Sub

Private Sub cmdClose_Click()
   Form_Unload 1
End Sub

Private Sub cmdEdit_Click()
    If MsgBox("Please ensure that there are no other users logged into Prestige before proceeding." & Chr(13) & _
              "Do you wish to continue?", vbYesNo, "Confirm Edit") = vbYes Then
        cmdEdit.Enabled = False
        txtTenantID.Locked = False
        cmdSave.Enabled = True
        txtTenantID.SetFocus
        Frame1.Enabled = False
    End If
End Sub

Private Sub UpdateSupplierID()
   Dim adoConn As New ADODB.Connection
   Dim Rst As New ADODB.Recordset
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
    sSQLQuery = "SELECT SupplierID " & _
               "FROM Supplier " & _
               "Where SupplierID = '" & txtTenantID.text & "'"
   
  Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
  If Not Rst.EOF Then
        MsgBox "This supplier ID already exists in the database. Please enter correct supplier ID", vbInformation
        Rst.Close
        Exit Sub
  End If
  Rst.Close
  

   sSQLQuery = "SELECT SupplierID " & _
               "FROM Supplier " & _
               "Where SupplierID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"

   Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   If Not Rst.EOF Then
       Rst!SupplierID = txtTenantID.text
       Rst.Update
   End If
   Rst.Close

   sSQLQuery = "UPDATE tblBatchPayment " & _
               "SET SupplierID ='" & txtTenantID.text & "' " & _
               "WHERE SupplierID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tblBatchTransaction " & _
               "SET SupplierID ='" & txtTenantID.text & "' " & _
               "WHERE SupplierID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tblPurInv " & _
               "SET SUPP_AC ='" & txtTenantID.text & "' " & _
               "WHERE SUPP_AC = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbPayment " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & txtTenantID.text & "' " & _
               "WHERE OwnerID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   cmdSave.Enabled = False
   txtTenantID.Locked = True

   adoConn.Close
   Set adoConn = Nothing
   
    MsgBox "SupplierID has been updated.", vbInformation + vbOKOnly, "Edit ID"
End Sub

Private Sub UpdateUnitID()
   Dim adoConn As New ADODB.Connection
   Dim Rst As New ADODB.Recordset
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
    'added by anol 20170207 issue 187
    
   sSQLQuery = "SELECT UnitNumber " & _
               "FROM Units " & _
               "Where UnitNumber = '" & txtTenantID.text & "'"
  Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
  If Not Rst.EOF Then
        MsgBox "This Unit ID already exists in the database. Please enter correct Unit ID", vbInformation
        Rst.Close
        Exit Sub
  End If
  Rst.Close
  
   sSQLQuery = "SELECT UnitNumber " & _
               "FROM Units " & _
               "Where UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"

   Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   If Not Rst.EOF Then
      Rst!UnitNumber = txtTenantID.text
      Rst.Update
   Else
      Rst.Close

      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If
   Rst.Close

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & txtTenantID.text & "' " & _
               "WHERE OwnerID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE DemandRecords " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE LeaseDetails " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tblBtRptTran " & _
               "SET UnitID ='" & txtTenantID.text & "' " & _
               "WHERE UnitID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tblPoA " & _
               "SET UnitID ='" & txtTenantID.text & "' " & _
               "WHERE UnitID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tblPrevGLU " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tblPurInvSRec " & _
               "SET UNIT_ID ='" & txtTenantID.text & "' " & _
               "WHERE UNIT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE TemplateUnitSelection " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbChildDemandRecord " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbCreditNote " & _
               "SET UNIT_ID ='" & txtTenantID.text & "' " & _
               "WHERE UNIT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbFloor " & _
               "SET UNIT_ID ='" & txtTenantID.text & "' " & _
               "WHERE UNIT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbImages " & _
               "SET PREMISIS_ID ='" & txtTenantID.text & "' " & _
               "WHERE PREMISIS_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbLetterReports " & _
               "SET UnitNo ='" & txtTenantID.text & "' " & _
               "WHERE UnitNo = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbReceipt " & _
               "SET UnitID ='" & txtTenantID.text & "' " & _
               "WHERE UnitID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbRecharged " & _
               "SET UNIT_ID ='" & txtTenantID.text & "' " & _
               "WHERE UNIT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE tlbRechargePre " & _
               "SET UNIT_ID ='" & txtTenantID.text & "' " & _
               "WHERE UNIT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE UnitSafety " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   sSQLQuery = "UPDATE UnitUtilities " & _
               "SET UnitNumber ='" & txtTenantID.text & "' " & _
               "WHERE UnitNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   cmdSave.Enabled = False
   txtTenantID.Locked = True

   adoConn.Close
   Set adoConn = Nothing
    MsgBox "UnitID has been updated.", vbInformation + vbOKOnly, "Edit ID"
End Sub

Private Sub UpdatePropertyID()
   Dim adoConn As New ADODB.Connection
   Dim Rst As New ADODB.Recordset
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
'added by anol 20170207 issue 187
   sSQLQuery = "SELECT PropertyID " & _
               "FROM Property " & _
               "Where PropertyID = '" & txtTenantID.text & "'"
   
  Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
  If Not Rst.EOF Then
        MsgBox "This Property ID already exists in the database. Please enter correct Property ID", vbInformation
        Rst.Close
        Exit Sub
  End If
  Rst.Close
   sSQLQuery = "SELECT PropertyID " & _
               "FROM Property " & _
               "Where PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"

   Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   If Not Rst.EOF Then
      Rst!propertyID = txtTenantID.text
      Rst.Update
   Else
      Rst.Close

      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If
   Rst.Close

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & txtTenantID.text & "' " & _
               "WHERE OwnerID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientGlobalData " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientProAgr " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE DemandTypes " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalData " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalInsurance " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalRC " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalSC " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE InterestRates " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE NLPosting " & _
               "SET PROPERTY_ID ='" & txtTenantID.text & "' " & _
               "WHERE PROPERTY_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyAnalysis " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyInsurance " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyLandlord " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyMaintHistory " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertySafety " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyUtilities " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchPayment " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchReceipt " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchTransaction " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblPurInv " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tlbBankPayment " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tlbPayment " & _
               "SET UnitID ='" & txtTenantID.text & "' " & _
               "WHERE UnitID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbPaymentSplit " & _
               "SET TRANS ='" & txtTenantID.text & "' " & _
               "WHERE TRANS = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
    sSQLQuery = "UPDATE UNITS " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
      '*******************************************************************
   'code added by anol 01/08/2023
   sSQLQuery = "UPDATE FundMatrix " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE ChargeTypes " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   'code added by anol 10/08/2023
   sSQLQuery = "UPDATE tlbReceiptSplit " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
  
   
   '*************************************************
   
   cmdSave.Enabled = False
   txtTenantID.Locked = True

   adoConn.Close
   Set adoConn = Nothing
        MsgBox "PropertyID has been updated.", vbInformation + vbOKOnly, "Edit ID"
End Sub

Private Sub UpdateClientID()
   Dim adoConn As New ADODB.Connection
   Dim Rst As New ADODB.Recordset
   adoConn.Open getConnectionString

   ' Event Type
   Dim sSQLQuery As String
   'added by anol 20170207 issue 187
   
   sSQLQuery = "SELECT ClientID " & _
               "FROM Client " & _
               "Where ClientID = '" & txtTenantID.text & "'"
   
  Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
  If Not Rst.EOF Then
        MsgBox "This client ID already exists in the database. Please enter correct client ID", vbInformation
        Rst.Close
        Exit Sub
  End If
  Rst.Close
   sSQLQuery = "SELECT ClientID " & _
               "FROM Client " & _
               "Where ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"

   Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   If Not Rst.EOF Then
       Rst!ClientID = txtTenantID.text
       Rst.Update
   End If
   Rst.Close

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & txtTenantID.text & "' " & _
               "WHERE OwnerID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE Client " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientGlobalData " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientProAgr " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE Property " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tblBatchPayment " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbBankPayment " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbClientBanks " & _
               "SET CLIENT_ID ='" & txtTenantID.text & "' " & _
               "WHERE CLIENT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbRecharged " & _
               "SET CLIENT_ID ='" & txtTenantID.text & "' " & _
               "WHERE CLIENT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE Units " & _
               "SET PropertyID ='" & txtTenantID.text & "' " & _
               "WHERE PropertyID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   'added by anol 20161110
   'Financial year is not updating the clientID
   sSQLQuery = "UPDATE FinancialYear " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   'NominalLedger
   sSQLQuery = "UPDATE NominalLedger " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   'NJ_Header
    sSQLQuery = "UPDATE NJ_Header " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
        adoConn.Execute (sSQLQuery)
   'NLPosting
    sSQLQuery = "UPDATE NLPosting " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
        adoConn.Execute (sSQLQuery)
  
   'tblBatchReceipt
   sSQLQuery = "UPDATE tblBatchReceipt " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
        adoConn.Execute (sSQLQuery)
     'tblPurInv
     sSQLQuery = "UPDATE tblPurInv " & _
               "SET CL_ID ='" & txtTenantID.text & "' " & _
               "WHERE CL_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
        adoConn.Execute (sSQLQuery)
    'tlbBankReconcilation
     sSQLQuery = "UPDATE tlbBankReconcilation " & _
           "SET ClientID ='" & txtTenantID.text & "' " & _
           "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    'tlbBankReconClosingBal
    
     sSQLQuery = "UPDATE tlbBankReconClosingBal " & _
           "SET ClientID ='" & txtTenantID.text & "' " & _
           "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    'tlbClientBanks
    sSQLQuery = "UPDATE tlbClientBanks " & _
           "SET CLIENT_ID ='" & txtTenantID.text & "' " & _
           "WHERE CLIENT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    'tlbPayment
    sSQLQuery = "UPDATE tlbPayment " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   'tlbReceipt
   sSQLQuery = "UPDATE tlbReceipt " & _
               "SET ClientID ='" & txtTenantID.text & "' " & _
               "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   'tlbRecharged
    sSQLQuery = "UPDATE tlbRecharged " & _
           "SET CLIENT_ID ='" & txtTenantID.text & "' " & _
           "WHERE CLIENT_ID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    
    'tlbPayable
    sSQLQuery = "UPDATE tlbPayable " & _
           "SET clientLandlordID ='" & txtTenantID.text & "' " & _
           "WHERE clientLandlordID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    
    'tlbPayable
    sSQLQuery = "UPDATE FundMatrix " & _
           "SET ClientID ='" & txtTenantID.text & "' " & _
           "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    
    'RentSummaryStatement
     sSQLQuery = "UPDATE RentSummaryStatement " & _
           "SET ClientIDLandlordID ='" & txtTenantID.text & "' " & _
           "WHERE ClientIDLandlordID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    'RentSummaryStatementdetails
     sSQLQuery = "UPDATE RentSummaryStatementdetails " & _
           "SET ClientID ='" & txtTenantID.text & "' " & _
           "WHERE ClientID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
    adoConn.Execute (sSQLQuery)
    
   'End of addition

   cmdSave.Enabled = False
   txtTenantID.Locked = True

   adoConn.Close
   Set adoConn = Nothing
        MsgBox "ClientID has been updated.", vbInformation + vbOKOnly, "Edit ID"
End Sub

Private Sub UpdateLesseeID()
   Dim adoConn As New ADODB.Connection
   Dim Rst As New ADODB.Recordset
   adoConn.Open getConnectionString
  
   ' Event Type
   Dim sSQLQuery As String
   'issue 187 by anol 20170208  validation for existing ID

   sSQLQuery = "SELECT  SageAccountNumber, TenantID " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & txtTenantID.text & "'"
  Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic
  If Not Rst.EOF Then
        If MsgBox("This Tenant ID already exists in the database. Do you still want to update this ID?", vbYesNo, "Please confirm?") = vbNo Then
            Rst.Close
            Exit Sub
        End If
  End If
  Rst.Close
   sSQLQuery = "SELECT  SageAccountNumber, TenantID " & _
               "FROM TENANTS " & _
               "Where TenantID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"

   Rst.Open sSQLQuery, adoConn, adOpenDynamic, adLockOptimistic

   If Not Rst.EOF Then
       Rst!SageAccountNumber = txtTenantID.text
       Rst!TenantID = txtTenantID.text
       Rst.Update
   End If
   Rst.Close

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & txtTenantID.text & "' " & _
               "WHERE OwnerID = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE LeaseDetails " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE DemandRecords " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE DemandRecPreview " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbChildDemandRecord " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbDRCurrentPrint " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbLetterReports " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE tlbReceipt " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE Units " & _
               "SET SageAccountNumber ='" & txtTenantID.text & "' " & _
               "WHERE SageAccountNumber = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE PropertyMaintHistory " & _
               "SET ReportedBy ='" & txtTenantID.text & "' " & _
               "WHERE ReportedBy = '" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "';"
   adoConn.Execute (sSQLQuery)
     'added by anol 20170207 issue 187 was not updating tenanat ID on the  NL posting correctly
   sSQLQuery = "Update NLPosting SET ACCOUNT_NUMBER='" & txtTenantID.text & "' where   ACCOUNT_NUMBER='" & gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) & "'"
   adoConn.Execute (sSQLQuery)
   
   gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) = txtTenantID.text
   cmdSave.Enabled = False
   txtTenantID.Locked = True

   adoConn.Close
   Set adoConn = Nothing
   
    MsgBox "LesseeID has been updated.", vbInformation + vbOKOnly, "Edit ID"
End Sub

Private Sub PrepareSupplierList()
   Dim adoConn As New ADODB.Connection
   Dim sSQLQuery_ As String

   adoConn.Open getConnectionString
'  Current tenants Only

   sSQLQuery_ = "SELECT " _
         & "SupplierID, SupplierName, " _
         & "iif(isnull(SupplierAddressLine1),'',SupplierAddressLine1) + ' ' + " _
         & "iif(isnull(SupplierAddressLine2),'',SupplierAddressLine2) + ' ' + " _
         & "iif(isnull(SupplierAddressLine3),'',SupplierAddressLine3) + ' ' + " _
         & "iif(isnull(SupplierAddressLine4),'',SupplierAddressLine4) as Address, " _
         & "SupplierPostCode , SupplierOfficeTel " _
         & "From Supplier Where TYPE='SUPPLIER' " _
         & "ORDER BY SupplierID;"
'Debug.Print sSQLQuery_
      PopulateTenantPropertyWise sSQLQuery_, adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub PreparePropertyList()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim Data() As String

   adoConn.Open getConnectionString
'  Current tenants Only

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   ReDim Data(TotalCol, TotalRow) As String

'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientList.Column() = Data()
'   cboClientList.ListIndex = 0
   adoRst.Close
   Set adoRst = Nothing

   szSQL = "SELECT " _
                  & "PropertyID, PropertyName, " _
                  & "iif(isnull(ProAddressLine1),'', ProAddressLine1) + ' ' + " _
                  & "iif(isnull(ProAddressLine2),'', ProAddressLine2) + ' ' + " _
                  & "iif(isnull(ProAddressLine3),'', ProAddressLine3) + ' ' + " _
                  & "iif(isnull(ProPostCode), '', ProPostCode) AS Address " _
              & "From Property " _
              & "ORDER BY PropertyID;"
'Debug.Print sSQLQuery_
''adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''   Dim iRow As Integer
''   iRow = 1
''
''   gridTenantLookup.Clear
''   gridTenantLookup.Rows = 2
''   gridTenantLookup.Cols = 6
''   ConfigurFlexGrid
''
''
''
''            iRow = 1
''            gridTenantLookup.AddItem ""
''            gridTenantLookup.TextMatrix(iRow, 1) = "ALL"
''            gridTenantLookup.TextMatrix(iRow, 2) = "ALL Clients"
''            gridTenantLookup.RowHeight(iRow) = 240
''
''           iRow = 2
''           While Not adoRst.EOF
''               gridTenantLookup.row = 1
''               gridTenantLookup.RowSel = 1
''               gridTenantLookup.ColSel = 1
''               gridTenantLookup.TextMatrix(iRow, 0) = ""
''               gridTenantLookup.TextMatrix(iRow, 1) = adoRst.Fields.Item(0).Value
''               gridTenantLookup.TextMatrix(iRow, 2) = adoRst.Fields.Item(1).Value
''               gridTenantLookup.TextMatrix(iRow, 3) = adoRst.Fields.Item(2).Value
''               gridTenantLookup.RowHeight(iRow) = 240
''               adoRst.MoveNext
''               If Not adoRst.EOF Then gridTenantLookup.AddItem ""
''               iRow = iRow + 1
''            Wend
''
''   adoRst.Close
   
   
   
      PopulateTenantPropertyWise szSQL, adoConn

   adoConn.Close
   Set adoConn = Nothing
   Exit Sub

NoRes:
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
   Exit Sub
End Sub

Private Sub PrepareClientList()
   Dim adoConn As New ADODB.Connection
   Dim sSQLQuery_ As String

   adoConn.Open getConnectionString
'  Current tenants Only

   sSQLQuery_ = "SELECT " _
                  & "ClientID, ClientName, " _
                  & "iif(isnull(ClientAddressLine1),'', ClientAddressLine1) + ' ' + " _
                  & "iif(isnull(ClientAddressLine2),'', ClientAddressLine2) + ' ' + " _
                  & "iif(isnull(ClientAddressLine2),'', ClientAddressLine2) + ' ' + " _
                  & "ClientPostCode , ClientOfficeTel " _
              & "From Client " _
              & "ORDER BY ClientID;"
'Debug.Print sSQLQuery_
      PopulateTenantPropertyWise sSQLQuery_, adoConn

   adoConn.Close
   Set adoConn = Nothing
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
   flxClient.ColWidth(0) = 80
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
           flxClient.TextMatrix(rRow, 2) = "ALL"
           flxClient.RowHeight(rRow) = 240
           flxClient.AddItem ""
           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    fraGrid.Enabled = True
End Sub

Private Sub cmdProperty_Click()
    sTextBox = "2"
     picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadPropertyList
    fraGrid.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxSupplier()
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
   lblClientID.Caption = "Supplier ID"
   lblClientName.Caption = "Supplier Name"
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
   szSQL = "SELECT SupplierID, SupplierName FROM Supplier order by SupplierID"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("SupplierID").Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("SupplierName").Value
               flxClient.RowHeight(rRow) = 240
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

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
           flxClient.AddItem ""
            flxClient.TextMatrix(rRow, 1) = "ALL"
            flxClient.TextMatrix(rRow, 2) = "ALL Clients"
            flxClient.RowHeight(rRow) = 240
            
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 240
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cmdSave_Click()
   If WhoAmI = "LesseeID" Then UpdateLesseeID
   If WhoAmI = "SupplierID" Then UpdateSupplierID
   If WhoAmI = "ClientID" Then UpdateClientID
   If WhoAmI = "PropertyID" Then UpdatePropertyID
   If WhoAmI = "UnitID" Then UpdateUnitID

'   MsgBox WhoAmI & " has been update.", vbInformation + vbOKOnly, "Edit ID"

   If WhoAmI = "LesseeID" Then PrepareLesseeList
   If WhoAmI = "SupplierID" Then PrepareSupplierList
   If WhoAmI = "ClientID" Then PrepareClientList
   If WhoAmI = "PropertyID" Then PreparePropertyList
   If WhoAmI = "UnitID" Then PrepareLesseeList

   Frame1.Enabled = True
End Sub

Private Sub Form_Activate()
   If WhoAmI = "LesseeID" Then
      PrepareLesseeList
'      fraGrid.Top = 480
'      fraGrid.Height = 2300
'      gridTenantLookup.Height = 2085
      lblIDName.Caption = "Lessee"
      Me.Caption = "Change Lessee ID"
      txtTenantID.MaxLength = 30
   End If
   If WhoAmI = "SupplierID" Then
      PrepareSupplierList
'      fraGrid.Top = 80
'      fraGrid.Height = 2655
'      gridTenantLookup.Height = 2445
      lblIDName.Caption = "Supplier"
      Me.Caption = "Change Supplier ID"
      txtTenantID.MaxLength = 10
   End If
   If WhoAmI = "ClientID" Then
      PrepareClientList
'      fraGrid.Top = 80
'      fraGrid.Height = 2655
'      gridTenantLookup.Height = 2445
      lblIDName.Caption = "Client"
      Me.Caption = "Change Client ID"
      txtTenantID.MaxLength = 10
   End If
   If WhoAmI = "PropertyID" Then
      PreparePropertyList
'      fraGrid.Top = 80
'      fraGrid.Height = 2655
'      gridTenantLookup.Height = 2445
      lblIDName.Caption = "Property"
      Me.Caption = "Change Property ID"
      Label44(1).Visible = False
      cmdProperty.Visible = False
      txtPropertyName.Visible = False
      txtTenantID.MaxLength = 10
   End If
   If WhoAmI = "UnitID" Then
      PrepareLesseeList
'      fraGrid.Top = 480
'      fraGrid.Height = 2300
'      gridTenantLookup.Height = 2085
      lblIDName.Caption = "Unit"
      Me.Caption = "Change Unit ID"
      txtTenantID.MaxLength = 12
   End If
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

       '        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
                        
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
Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 9210
   Me.Height = 8985
   Me.BackColor = MODULEBACKCOLOR
   fraGrid.Left = 0
   Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    '#
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    If WhoAmI = "SupplierID" Then
'        szSQL = "SELECT SupplierID, SupplierName FROM Supplier order by SupplierID"
'        adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'        If Not adoRst.EOF Then
'             txtClientList.Tag = adoRst.Fields("SupplierID").Value
'             txtClientList.text = adoRst.Fields("SupplierName").Value
''             txtPropertyName.Tag = "ALL"
''             txtPropertyName.text = "ALL"
''             cboPropertyList_Click
'
'        End If
        txtClientList.Tag = "ALL"
        txtClientList.text = "ALL"
        Label44(1).Visible = False
        cmdProperty.Enabled = False
        Label44(0).Caption = "Supplier"
    Else
        Label44(0).Caption = "Client"
        Label44(1).Visible = True
        cmdProperty.Enabled = True
        szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not adoRst.EOF Then
             txtClientList.Tag = adoRst.Fields("CLIENTID").Value
             txtClientList.text = adoRst.Fields("CLIENTNAME").Value
             txtPropertyName.Tag = "ALL"
             txtPropertyName.text = "ALL"
             cboPropertyList_Click
        End If
        adoRst.Close
   End If
'   adoRst.Close
   Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   WhoAmI = ""
   Unload Me
End Sub





Private Sub gridTenantLookup_Click()
   txtTenantID.text = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
'   txtTenantID.SetFocus
'   txtTenantID.SelStart = Len(txtTenantID.text)
   gridTenantLookup.TextMatrix(gridTenantLookup.row, 0) = gridTenantLookup.TextMatrix(gridTenantLookup.row, 0)
   cmdEdit.Enabled = True
   HighLightRowFlxGrid gridTenantLookup, gridTenantLookup.row
'   gridTenantLookup.SetFocus
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

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
          
         
          'If sTextBox = "1" Then
           cmdClientList.SetFocus
'           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           'End If
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
Private Sub flxClient_Click()
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    txtPropertyName.Tag = "ALL"
                    txtPropertyName.text = "ALL"
                    cboClientList_Click
                   If cmdProperty.Enabled And cmdProperty.Visible Then cmdProperty.SetFocus
            ElseIf sTextBox = "2" Then
                    txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                    cboPropertyList_Click
            ElseIf sTextBox = "3" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    cboPropertyList_Click
            End If
            picClient.Visible = False
            fraGrid.Enabled = True
        
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
