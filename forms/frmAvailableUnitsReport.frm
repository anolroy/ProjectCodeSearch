VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAvailableUnitsReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Units Report"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmAvailableUnitsReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7860
   Begin VB.CheckBox chkAll 
      Caption         =   "Select All Units"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUnits 
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rent"
      Height          =   195
      Index           =   4
      Left            =   6960
      TabIndex        =   11
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Area"
      Height          =   195
      Index           =   3
      Left            =   5520
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Number"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   6
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
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin MSForms.ComboBox cboDmdPropertyList 
      Height          =   315
      Left            =   4755
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5106;556"
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
   Begin MSForms.ComboBox cboDmdClientList 
      Height          =   315
      Left            =   915
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5106;556"
      BoundColumn     =   0
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
End
Attribute VB_Name = "frmAvailableUnitsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PrepareList(cboClient As Control, cboProperty As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   
   On Error GoTo ErrorHandler
   
   adoConn.Open getConnectionString

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

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
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing

End Sub

Private Sub cboDmdClientList_Change()
    'loadGridUnits
End Sub

Private Sub cboDmdClientList_Click()
    loadGridUnits
    chkAll.Value = 0
End Sub

Private Sub cboDmdPropertyList_Change()
    'loadGridUnits
End Sub

Private Sub cboDmdPropertyList_Click()
    loadGridUnits
End Sub

Private Sub chkAll_Click()
   Dim i As Integer

   If Not gridUnits.TextMatrix(1, 1) = "" Then
      If chkAll.Value Then
         For i = 1 To gridUnits.Rows - 1
            gridUnits.TextMatrix(i, 0) = "X"
         Next i
      Else
         For i = 1 To gridUnits.Rows - 1
            gridUnits.TextMatrix(i, 0) = ""
         Next i
      End If
   Else
     Exit Sub
   End If
End Sub

Private Sub cmdPrint_Click()
   Dim flag As Boolean
   flag = False
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, i As Integer, szaUnit() As String

   Dim rstRst As New ADODB.Recordset
   
   If gridUnits.row < 1 Then Exit Sub
   For i = 1 To gridUnits.Rows - 1
    If (gridUnits.TextMatrix(i, 0) = "X") Then
         flag = True
    End If
   Next i
   If flag = False Then
      MsgBox "Please select at least one unit.", vbCritical + vbOKOnly, "Units Report"
      Exit Sub
   End If
   adoConn.Open getConnectionString

   szSQL = "UPDATE Units SET isPrint = '';"

   adoConn.Execute szSQL

   For i = 1 To gridUnits.Rows - 1
    If (gridUnits.TextMatrix(i, 0) = "X") Then
      szSQL = "UPDATE Units SET isPrint = 'Y' " & _
              "WHERE UnitNumber = '" & gridUnits.TextMatrix(i, 1) & "';"

      adoConn.Execute szSQL
    End If
   Next i

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\VacantUnitDetailsReport.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub Form_Load()
   Me.Top = (frmMMain.Height / 2) - (Me.Height / 2) - 400
   Me.Left = (frmMMain.Width / 2) - (Me.Width / 2) - 400
   Me.BackColor = MODULEBACKCOLOR
   chkAll.BackColor = MODULEBACKCOLOR

   PrepareList cboDmdClientList, cboDmdPropertyList

   loadGridUnits
End Sub

Private Sub configGridUnits()
   Dim szHeader As String, iCol As Integer

   gridUnits.Clear
   gridUnits.Cols = 5
   gridUnits.Rows = 2
   gridUnits.RowHeight(0) = 0
   szHeader$ = "|<Unit Number|<Name|<Post Code|>Total Area|>Rent"
   gridUnits.FormatString = szHeader$
   
   gridUnits.ColWidth(0) = 200
   gridUnits.ColWidth(1) = 1200
   gridUnits.ColWidth(2) = 1600
   gridUnits.ColWidth(3) = 1600
   gridUnits.ColWidth(4) = 1200
   gridUnits.ColWidth(5) = 1200
   
   Label19(0).Left = 500
   Label19(1).Left = 2000
   Label19(2).Left = 3400
   Label19(3).Left = 4900
   Label19(4).Left = 6300
End Sub

Private Sub loadGridUnits()
   Dim conUnit As New ADODB.Connection
   Dim rstReport As New ADODB.Recordset

   Dim sSQLQuery_ As String, iRow As Integer

   configGridUnits

   If cboDmdClientList.Column(0) <> "ALL" And cboDmdPropertyList.Column(0) <> "ALL" Then
        sSQLQuery_ = "SELECT Units.UnitNumber, Units.UnitName, Units.UnitPostCode, Units.TotalArea, Units.RentalPrice " & _
                     "FROM Client, Property, Units " & _
                     "WHERE Property.PropertyID = Units.PropertyID AND Client.ClientID = Property.ClientID AND " & _
                     "Units.Occupied = 'N' AND Property.PropertyID = '" & cboDmdPropertyList.Column(0) & "'  AND " & _
                           "Client.ClientID = '" & cboDmdClientList.Column(0) & "'"
   ElseIf cboDmdClientList.Column(0) = "ALL" And cboDmdPropertyList.Column(0) <> "ALL" Then
        sSQLQuery_ = "SELECT Units.UnitNumber, Units.UnitName, Units.UnitPostCode, Units.TotalArea, Units.RentalPrice " & _
                     "FROM Client, Property, Units " & _
                     "WHERE Property.PropertyID = Units.PropertyID AND Client.ClientID = Property.ClientID AND " & _
                           "Units.Occupied = 'N' AND Property.PropertyID = '" & cboDmdPropertyList.Column(0) & "'"
   ElseIf cboDmdClientList.Column(0) <> "ALL" And cboDmdPropertyList.Column(0) = "ALL" Then
        sSQLQuery_ = "SELECT Units.UnitNumber, Units.UnitName, Units.UnitPostCode, Units.TotalArea, Units.RentalPrice " & _
                     "FROM Client, Property, Units " & _
                     "WHERE Property.PropertyID = Units.PropertyID AND Client.ClientID = Property.ClientID AND " & _
                           "Units.Occupied = 'N' AND Client.ClientID = '" & cboDmdClientList.Column(0) & "'"
   Else
        sSQLQuery_ = "SELECT Units.UnitNumber, Units.UnitName, Units.UnitPostCode, Units.TotalArea, Units.RentalPrice " & _
                     "FROM Client, Property, Units " & _
                     "WHERE Property.PropertyID = Units.PropertyID AND Client.ClientID = Property.ClientID AND " & _
                           "Units.Occupied = 'N'"
   End If

   conUnit.Open getConnectionString
   rstReport.Open sSQLQuery_, conUnit, adOpenStatic, adLockReadOnly
   
   iRow = 1
   While Not rstReport.EOF
        
        gridUnits.TextMatrix(iRow, 1) = IIf(IsNull(rstReport!UnitNumber), "", rstReport!UnitNumber)
        gridUnits.TextMatrix(iRow, 2) = IIf(IsNull(rstReport!UnitName), "", rstReport!UnitName)
        gridUnits.TextMatrix(iRow, 3) = IIf(IsNull(rstReport!UnitPostCode), "", rstReport!UnitPostCode)
        gridUnits.TextMatrix(iRow, 4) = IIf(IsNull(rstReport!TotalArea), "", rstReport!TotalArea)
        gridUnits.TextMatrix(iRow, 5) = IIf(IsNull(rstReport!RentalPrice), "", rstReport!RentalPrice)
   
        iRow = iRow + 1
        rstReport.MoveNext
        If Not rstReport.EOF Then gridUnits.AddItem ""
   Wend
   rstReport.Close
   Set rstReport = Nothing
   conUnit.Close
   Set conUnit = Nothing
End Sub

Private Sub gridUnits_Click()
    Dim i As Integer
    i = Select1RowFlxGrid(gridUnits, gridUnits.row, 0)
End Sub
