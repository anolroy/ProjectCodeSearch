VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRecharge 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kingsgate - Recharge"
   ClientHeight    =   6360
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   12510
   Icon            =   "frmRecharge.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   12510
   Begin VB.Frame Frame1 
      BackColor       =   &H00C89CA9&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5145
      TabIndex        =   10
      Top             =   5640
      Width           =   7335
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FAD5DF&
         Caption         =   "&Save"
         Height          =   400
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdGenRecharge 
         BackColor       =   &H00FAD5DF&
         Caption         =   "Select for Recharge"
         Height          =   400
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdRemoveCharge 
         BackColor       =   &H00FAD5DF&
         Caption         =   "&Remove Charge"
         Height          =   400
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FAD5DF&
         Caption         =   "C&lose"
         Height          =   400
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.TextBox txtLandLord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEFB0&
      Height          =   285
      Left            =   8740
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox txtReDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtTotalNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtTotalRecharge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtRecharge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11040
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboUnits 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "cboUnits"
      Top             =   120
      Width           =   3375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRecharge 
      Height          =   4455
      Left            =   100
      TabIndex        =   14
      Top             =   720
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   13
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin MSForms.CheckBox chkSelectAll 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
      VariousPropertyBits=   746588179
      BackColor       =   16768960
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1931;450"
      Value           =   "0"
      Caption         =   "Select All"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   9885
      TabIndex        =   8
      Top             =   5280
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recharge Expenses/Costs"
      Height          =   195
      Left            =   100
      TabIndex        =   3
      Top             =   480
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Landlord"
      Height          =   195
      Left            =   7725
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
      Height          =   195
      Left            =   100
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmRecharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szUnitID As String
Dim bData As Boolean
Dim cNet As Currency
Dim iSelected As Integer

Private Sub cboUnits_Click()
   Dim i As Integer, j As Integer, match As Integer, SQLStr1 As String
   Dim adoConn As New ADODB.Connection
   Dim Rst1 As New ADODB.Recordset
   
   txtLandLord.text = ""
   
   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   Dim szUnit() As String
   szUnit = Split(cboUnits.text, " \ ")
   szUnitID = Trim(szUnit(0))
   'Get the record for selected unit
   SQLStr1 = "SELECT UnitNumber, PropertyID, UnitName, UnitAddressLine1, " & _
               "UnitAddressLine2, UnitAddressLine3, UnitAddressLine4, " & _
               "UnitPostCode, Occupied, TenantCompanyName, SageAccountNumber, " & _
               "Frontage, RateableValue, RatesPayable, GroundFloorArea, " & _
               "MezzanineArea, TotalArea, UnitType, LandLord, Management, Memo " & _
             "FROM Units WHERE UnitNumber = '" & szUnitID & "'"
   Rst1.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   szUnit(0) = GetClientByUnit(szUnitID, adoConn)
   If szUnit(0) <> "ERROR" Then txtLandLord.text = szUnit(0)

   Rst1.Close
   adoConn.Close
   
   bData = False
   
   Call CollectData
End Sub

Private Sub CollectData()
   Dim iRow As Integer
   Dim adoConn As New ADODB.Connection
   
   If flxRecharge.Rows > 2 Then
      For iRow = 2 To flxRecharge.Rows - 1
         flxRecharge.RemoveItem 2
      Next iRow
   End If
   For iRow = 0 To 10
      flxRecharge.TextMatrix(1, iRow) = ""
   Next iRow
   flxRecharge.row = 1
   For iRow = 1 To flxRecharge.Cols - 1
      flxRecharge.col = iRow
      flxRecharge.CellBackColor = RGB(255, 255, 255)
   Next iRow

   Dim szSQL As String, szSQLCN As String, szSqlBank As String, szBackUp As String
   Dim adoPI      As New ADODB.Recordset
   Dim rstCN      As New ADODB.Recordset
   Dim rstBank    As New ADODB.Recordset
   Dim rstBackUp  As New ADODB.Recordset

   adoConn.Open getConnectionString

'Collect all "purchase invoice" transactions which have not been charged yet
'     ====== tblPurInv =========
   szSQL = "SELECT TRAN_ID,SUPP_AC,TRAN_DATE,TRAN_TYPE,INV_NO,DESCRIPTION,OUT_PAYMENT," & _
             "TRANS,NOMINAL_CODE,DEPT_ID,PROJ_REF,COST_CODE,NET_AMOUNT,VAT " & _
           "FROM tblPurInv " & _
           "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
Debug.Print szSQL
   adoPI.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

'Collect all "credit note" transactions which have not been charged yet
'     ===== tlbCreditNote ========
   szSQLCN = "SELECT TRAN_ID,SUPP_AC,TRAN_DATE,TRAN_TYPE,INV_NO,DESCRIPTION,OUT_PAYMENT," & _
             "TRANS,NOMINAL_CODE,DEPT_ID,PROJ_REF,COST_CODE,NET_AMOUNT,VAT " & _
             "FROM tlbCreditNote " & _
             "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
   rstCN.Open szSQLCN, adoConn, adOpenDynamic, adLockOptimistic
   
'Collect all "bank payment" transactions which have not been charged yet and rechargable
'     ===== tlbBankPayment ========
   szSqlBank = "SELECT TRAN_ID,BANK_AC,TRAN_DATE,TRAN_TYPE,NOMINAL_CODE,DEPT_ID," & _
                  "PROJ_REF,COST_CODE,DESCRIPTION,NET_AMOUNT,TAX_CODE,VAT,TRANS " & _
               "FROM tlbBankPayment " & _
               "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No AND " & _
                  "RECHARABLE=Yes " & _
               "ORDER BY TRAN_TYPE"
   rstBank.Open szSqlBank, adoConn, adOpenDynamic, adLockOptimistic

   szBackUp = "SELECT MY_ID, UNIT_ID, LANDLORD_ID, TRAN_ID, TRAN_DATE, " & _
                  "TRAN_TYPE, TRANS, INV_NO, NOMINAL_CODE, DEPT_ID, PROJ_REF, " & _
                  "COST_CODE, DESCRIPTION, NET_AMOUNT, VAT, RECHARGE, CHECKED " & _
              "FROM tlbRechargePre " & _
              "WHERE UNIT_ID='" & szUnitID & "';"
   rstBackUp.Open szBackUp, adoConn, adOpenDynamic, adLockOptimistic

   If adoPI.EOF And rstCN.EOF And rstBank.EOF And rstBackUp.EOF Then GoTo ErrorHandler

   flxRecharge.RowHeightMin = 285

   Dim szTempTable As String
   Dim Rst1 As New ADODB.Recordset
   Dim szLandLord() As String

   szTempTable = "SELECT MY_ID, UNIT_ID, LANDLORD_ID, TRAN_ID, TRAN_DATE, " & _
                  "TRAN_TYPE, TRANS, INV_NO, NOMINAL_CODE, DEPT_ID, PROJ_REF, " & _
                  "COST_CODE, DESCRIPTION, NET_AMOUNT, VAT, RECHARGE, CHECKED " & _
                 "FROM tlbRechargePre"
   Rst1.Open szTempTable, adoConn, adOpenDynamic, adLockOptimistic
   szLandLord = Split(cboUnits.text, " \ ")

   Dim iCol As Integer

   iRow = 1
' View all data from tlbRechargePre which saved backup data
   While Not rstBackUp.EOF
      If rstBackUp!Checked = True Then
         flxRecharge.TextMatrix(iRow, 0) = "X"
         iSelected = iSelected + 1                       'counting selected rows
         flxRecharge.row = iRow
         For iCol = 1 To flxRecharge.Cols - 1
            flxRecharge.col = iCol
            flxRecharge.CellBackColor = RGB(180, 255, 180)
         Next iCol
      End If
      flxRecharge.TextMatrix(iRow, 1) = rstBackUp!TRAN_ID
      flxRecharge.TextMatrix(iRow, 2) = rstBackUp!TRAN_DATE
      flxRecharge.TextMatrix(iRow, 3) = rstBackUp!TRAN_TYPE
      flxRecharge.TextMatrix(iRow, 4) = rstBackUp!TRANS
      flxRecharge.TextMatrix(iRow, 5) = rstBackUp!INV_NO
      flxRecharge.TextMatrix(iRow, 6) = rstBackUp!Nominal_code
      flxRecharge.TextMatrix(iRow, 7) = rstBackUp!DEPT_ID
      flxRecharge.TextMatrix(iRow, 8) = rstBackUp!PROJ_REF
      flxRecharge.TextMatrix(iRow, 9) = rstBackUp!COST_CODE
      flxRecharge.TextMatrix(iRow, 10) = rstBackUp!description
      flxRecharge.TextMatrix(iRow, 11) = Format(rstBackUp!NET_AMOUNT, "0.00")
      flxRecharge.TextMatrix(iRow, 12) = Format(rstBackUp!RECHARGE, "0.00")

      rstBackUp.MoveNext
      If Not rstBackUp.EOF Then flxRecharge.AddItem ""
      iRow = iRow + 1
      bData = True
   Wend
   rstBackUp.Close
   Set rstBackUp = Nothing
   
   
   While Not adoPI.EOF
      If Not adoPI.EOF And bData Then flxRecharge.AddItem ""
      flxRecharge.TextMatrix(iRow, 1) = adoPI!TRAN_ID
      flxRecharge.TextMatrix(iRow, 2) = adoPI!TRAN_DATE
      flxRecharge.TextMatrix(iRow, 3) = adoPI!TRAN_TYPE
      flxRecharge.TextMatrix(iRow, 4) = adoPI!TRANS
      flxRecharge.TextMatrix(iRow, 5) = adoPI!INV_NO
      flxRecharge.TextMatrix(iRow, 6) = adoPI!Nominal_code
      flxRecharge.TextMatrix(iRow, 7) = adoPI!DEPT_ID
      flxRecharge.TextMatrix(iRow, 8) = adoPI!PROJ_REF
      flxRecharge.TextMatrix(iRow, 9) = adoPI!COST_CODE
      flxRecharge.TextMatrix(iRow, 10) = adoPI!description
      flxRecharge.TextMatrix(iRow, 11) = Format(adoPI!NET_AMOUNT, "0.00")
      flxRecharge.TextMatrix(iRow, 12) = "0.00"
      
      bData = True

      Rst1.AddNew
      Rst1!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
      Rst1!UNIT_ID = szUnitID
      Rst1!LANDLORD_ID = szLandLord(0)
      If flxRecharge.TextMatrix(iRow, 0) <> "" Then Rst1!Checked = 1
      Rst1!TRAN_DATE = Format(flxRecharge.TextMatrix(iRow, 2), "DD MMMM YYYY")
      Rst1!TRANS = flxRecharge.TextMatrix(iRow, 4)
      Rst1!TRAN_TYPE = flxRecharge.TextMatrix(iRow, 3)
      Rst1!TRAN_ID = flxRecharge.TextMatrix(iRow, 1)
      Rst1!Nominal_code = flxRecharge.TextMatrix(iRow, 6)
      Rst1!DEPT_ID = flxRecharge.TextMatrix(iRow, 7)
      Rst1!PROJ_REF = flxRecharge.TextMatrix(iRow, 8)
      Rst1!COST_CODE = flxRecharge.TextMatrix(iRow, 9)
      Rst1!description = flxRecharge.TextMatrix(iRow, 10)
      Rst1!NET_AMOUNT = CCur(flxRecharge.TextMatrix(iRow, 11))
      Rst1!INV_NO = flxRecharge.TextMatrix(iRow, 5)
      Rst1!RECHARGE = CCur(flxRecharge.TextMatrix(iRow, 12))
      Rst1.Update

      adoPI.MoveNext
      flxRecharge.RowHeight(flxRecharge.Rows - 1) = 285
      iRow = iRow + 1
   Wend
   adoPI.Close
   Set adoPI = Nothing
   
   While Not rstCN.EOF
      If Not rstCN.EOF And bData Then flxRecharge.AddItem ""
      flxRecharge.RowHeight(flxRecharge.Rows - 1) = 285
      flxRecharge.TextMatrix(iRow, 1) = rstCN!TRAN_ID
      flxRecharge.TextMatrix(iRow, 2) = rstCN!TRAN_DATE
      flxRecharge.TextMatrix(iRow, 3) = rstCN!TRAN_TYPE
      flxRecharge.TextMatrix(iRow, 4) = rstCN!TRANS
      flxRecharge.TextMatrix(iRow, 5) = rstCN!INV_NO
      flxRecharge.TextMatrix(iRow, 6) = rstCN!Nominal_code
      flxRecharge.TextMatrix(iRow, 7) = rstCN!DEPT_ID
      flxRecharge.TextMatrix(iRow, 8) = rstCN!PROJ_REF
      flxRecharge.TextMatrix(iRow, 9) = rstCN!COST_CODE
      flxRecharge.TextMatrix(iRow, 10) = rstCN!description
      flxRecharge.TextMatrix(iRow, 11) = Format(rstCN!NET_AMOUNT, "0.00")
      flxRecharge.TextMatrix(iRow, 12) = "0.00"

      Rst1.AddNew
      Rst1!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
      Rst1!UNIT_ID = szUnitID
      Rst1!LANDLORD_ID = szLandLord(0)
      If flxRecharge.TextMatrix(iRow, 0) <> "" Then Rst1!Checked = 1
      Rst1!TRAN_DATE = Format(flxRecharge.TextMatrix(iRow, 2), "DD MMMM YYYY")
      Rst1!TRANS = flxRecharge.TextMatrix(iRow, 4)
      Rst1!TRAN_TYPE = flxRecharge.TextMatrix(iRow, 3)
      Rst1!TRAN_ID = flxRecharge.TextMatrix(iRow, 1)
      Rst1!Nominal_code = flxRecharge.TextMatrix(iRow, 6)
      Rst1!DEPT_ID = flxRecharge.TextMatrix(iRow, 7)
      Rst1!PROJ_REF = flxRecharge.TextMatrix(iRow, 8)
      Rst1!COST_CODE = flxRecharge.TextMatrix(iRow, 9)
      Rst1!description = flxRecharge.TextMatrix(iRow, 10)
      Rst1!NET_AMOUNT = CCur(flxRecharge.TextMatrix(iRow, 11))
      Rst1!INV_NO = flxRecharge.TextMatrix(iRow, 5)
      Rst1!RECHARGE = CCur(flxRecharge.TextMatrix(iRow, 12))
      Rst1.Update

      rstCN.MoveNext
      iRow = iRow + 1
      bData = True
   Wend
   
   rstCN.Close
   Set rstCN = Nothing
   
   While Not rstBank.EOF
      If Not rstBank.EOF And bData Then flxRecharge.AddItem ""
      flxRecharge.RowHeight(flxRecharge.Rows - 1) = 285
      flxRecharge.TextMatrix(iRow, 1) = rstBank!TRAN_ID
      flxRecharge.TextMatrix(iRow, 2) = rstBank!TRAN_DATE
      flxRecharge.TextMatrix(iRow, 3) = rstBank!TRAN_TYPE
      flxRecharge.TextMatrix(iRow, 4) = rstBank!TRANS
      flxRecharge.TextMatrix(iRow, 5) = ""
      flxRecharge.TextMatrix(iRow, 6) = rstBank!Nominal_code
      flxRecharge.TextMatrix(iRow, 7) = rstBank!DEPT_ID
      flxRecharge.TextMatrix(iRow, 8) = rstBank!PROJ_REF
      flxRecharge.TextMatrix(iRow, 9) = rstBank!COST_CODE
      flxRecharge.TextMatrix(iRow, 10) = rstBank!description
      flxRecharge.TextMatrix(iRow, 11) = Format(rstBank!NET_AMOUNT, "0.00")
      flxRecharge.TextMatrix(iRow, 12) = "0.00"

      Rst1.AddNew
      Rst1!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
      Rst1!UNIT_ID = szUnitID
      Rst1!LANDLORD_ID = szLandLord(0)
      If flxRecharge.TextMatrix(iRow, 0) <> "" Then Rst1!Checked = 1
      Rst1!TRAN_DATE = Format(flxRecharge.TextMatrix(iRow, 2), "DD MMMM YYYY")
      Rst1!TRANS = flxRecharge.TextMatrix(iRow, 4)
      Rst1!TRAN_TYPE = flxRecharge.TextMatrix(iRow, 3)
      Rst1!TRAN_ID = flxRecharge.TextMatrix(iRow, 1)
      Rst1!Nominal_code = flxRecharge.TextMatrix(iRow, 6)
      Rst1!DEPT_ID = flxRecharge.TextMatrix(iRow, 7)
      Rst1!PROJ_REF = flxRecharge.TextMatrix(iRow, 8)
      Rst1!COST_CODE = flxRecharge.TextMatrix(iRow, 9)
      Rst1!description = flxRecharge.TextMatrix(iRow, 10)
      Rst1!NET_AMOUNT = CCur(flxRecharge.TextMatrix(iRow, 11))
      Rst1!INV_NO = flxRecharge.TextMatrix(iRow, 5)
      Rst1!RECHARGE = CCur(flxRecharge.TextMatrix(iRow, 12))
      Rst1.Update
      
      rstBank.MoveNext
      iRow = iRow + 1
      bData = True
   Wend
   rstBank.Close
   Set rstBank = Nothing

   Rst1.Close
   Set Rst1 = Nothing

   flxRecharge.col = 1
   flxRecharge.Sort = flexSortNumericAscending

   txtTotalNet.text = Format(CStr(CalculateTotal(11)), "0.00")
   txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")

'**********************************************************************************
' UPDATE SOURCE TABLES. RECHARGED TURN TO 'YES'

   szSQL = "SELECT RECHARGED " & _
             "FROM tblPurInv " & _
             "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
   adoPI.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   szSQLCN = "SELECT RECHARGED " & _
             "FROM tlbCreditNote " & _
             "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
   rstCN.Open szSQLCN, adoConn, adOpenDynamic, adLockOptimistic

   szSqlBank = "SELECT RECHARGED " & _
               "FROM tlbBankPayment " & _
               "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No " & _
               "ORDER BY TRAN_TYPE"
   rstBank.Open szSqlBank, adoConn, adOpenDynamic, adLockOptimistic
   
   If adoPI.EOF And rstCN.EOF And rstBank.EOF Then GoTo ErrorHandler
      
   flxRecharge.RowHeight(1) = 285

   iRow = 1
   adoPI.MoveFirst
   While Not adoPI.EOF
      adoPI!RECHARGED = 1
      adoPI.Update
      adoPI.MoveNext
   Wend
   adoPI.Close
   Set adoPI = Nothing
   
   rstCN.MoveFirst
   While Not rstCN.EOF
      rstCN!RECHARGED = 1
      rstCN.Update
      rstCN.MoveNext
   Wend
   rstCN.Close
   Set rstCN = Nothing
   
   While Not rstBank.EOF
      rstBank!RECHARGED = 1
      rstBank.Update
      rstBank.MoveNext
   Wend
   rstBank.Close
   Set rstBank = Nothing
   
   adoConn.Close
   Set adoConn = Nothing
'****************************************************************************
   Exit Sub

ErrorHandler:
   adoPI.Close
   rstCN.Close
   rstBank.Close

   Set adoPI = Nothing
   Set rstCN = Nothing
   Set rstBank = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function CalculateTotal(iCol As Integer) As Currency
   Dim iRow As Integer
   Dim cTotal As Currency
   
   cTotal = 0
   For iRow = 1 To flxRecharge.Rows - 1
      If Trim(flxRecharge.TextMatrix(iRow, 3)) = "BP" Or Trim(flxRecharge.TextMatrix(iRow, 3)) = "PI" Then
         cTotal = cTotal + CCur(flxRecharge.TextMatrix(iRow, iCol))
      End If
      If Trim(flxRecharge.TextMatrix(iRow, 3)) = "BR" Or Trim(flxRecharge.TextMatrix(iRow, 3)) = "PC" Then
         cTotal = cTotal - CCur(flxRecharge.TextMatrix(iRow, iCol))
      End If
   Next iRow
   CalculateTotal = cTotal
End Function

Private Sub cboUnits_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub chkSelectAll_Click()
   Dim iRow As Integer, iCol As Integer

   If chkSelectAll.Value Then
      If Not bData Then Exit Sub
      For iRow = 1 To flxRecharge.Rows - 1
         flxRecharge.row = iRow
         flxRecharge.TextMatrix(iRow, 0) = "X"
         For iCol = 1 To flxRecharge.Cols - 1
            flxRecharge.col = iCol
            flxRecharge.CellBackColor = RGB(180, 255, 180)
            flxRecharge.TextMatrix(flxRecharge.row, 12) = Format(flxRecharge.TextMatrix(flxRecharge.row, 11), "0.00")
         Next iCol
      Next iRow
      
      txtTotalNet.text = Format(CStr(CalculateTotal(11)), "0.00")
      txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
   Else
      For iRow = 1 To flxRecharge.Rows - 1
         flxRecharge.row = iRow
         flxRecharge.TextMatrix(iRow, 0) = ""
         For iCol = 1 To flxRecharge.Cols - 1
            flxRecharge.col = iCol
            flxRecharge.CellBackColor = RGB(255, 255, 255)
            flxRecharge.TextMatrix(flxRecharge.row, 12) = "0.00"
         Next iCol
      Next iRow
      
      txtTotalNet.text = Format(CStr(CalculateTotal(11)), "0.00")
      txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdGenRecharge_Click()
   If iSelected = 0 Then
      ShowMsgInTaskBar "No Selection", , "N"
      Exit Sub
   End If
   If MsgBox("Are you sure to recharge?", vbYesNo, "Recharged") = vbNo Then Exit Sub

   Dim szLandLord() As String, szSQL As String
   Dim iRow As Integer
   Dim Conn1 As New ADODB.Connection
   Dim rstRecharged As New ADODB.Recordset, rstBackUp As New ADODB.Recordset

   If bData = False Then
      ShowMsgInTaskBar "No data has been recharged!", , "N"
      Exit Sub
   End If

   On Error GoTo ErrorHandler

   Conn1.Open getConnectionString

   szSQL = "SELECT MY_ID, RECHARGE_DATE, CLIENT_ID, TYPE, EX_REF, " & _
            "TRANS, UNIT_ID, NOMINAL_CODE, DEPT_ID, PROJ_REF, DESCRIPTION, " & _
            "RECHARGE, VAT, UPDATE_SAGE, TYPE_ORG " & _
           "FROM tlbRecharged"
   rstRecharged.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   szLandLord = Split(txtLandLord.text, " \ ")
   For iRow = 1 To flxRecharge.Rows - 1
      If flxRecharge.TextMatrix(iRow, 0) = "X" Then
         rstRecharged.AddNew
         rstRecharged!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
         rstRecharged!RECHARGE_DATE = Format(txtReDate.text, "DD MMMM YYYY")
         rstRecharged!CLIENT_ID = szLandLord(0)
         If flxRecharge.TextMatrix(iRow, 3) = "PI" Or flxRecharge.TextMatrix(iRow, 3) = "BP" Then
            rstRecharged!Type = "SI"
         Else
            rstRecharged!Type = "SC"
         End If
         rstRecharged!EX_REF = flxRecharge.TextMatrix(iRow, 5)
         rstRecharged!TRANS = flxRecharge.TextMatrix(iRow, 4)
         rstRecharged!UNIT_ID = szUnitID
' i have blocked the following code because of disconnecting of sage. we might need it later when we will
' introduce nominal code for supplier.
'         rstRecharged!NOMINAL_CODE = LandLordNominalCode(szLandLord(0))
         rstRecharged!DEPT_ID = flxRecharge.TextMatrix(iRow, 7)
         rstRecharged!PROJ_REF = ""
         rstRecharged!description = flxRecharge.TextMatrix(iRow, 10)
         rstRecharged!RECHARGE = CCur(flxRecharge.TextMatrix(iRow, 12))
         rstRecharged!vat = LLVat(CCur(flxRecharge.TextMatrix(iRow, 12)))
         rstRecharged!TYPE_ORG = flxRecharge.TextMatrix(iRow, 3)
         rstRecharged.Update
      End If
   Next iRow

   rstRecharged.Close

   szSQL = "DELETE * FROM tlbRechargePre WHERE "
   For iRow = 1 To flxRecharge.Rows - 1
      If flxRecharge.TextMatrix(iRow, 0) = "X" Then
         szSQL = szSQL + "(TRAN_ID= '" + flxRecharge.TextMatrix(iRow, 1) + _
                  "' AND TRAN_TYPE='" + flxRecharge.TextMatrix(iRow, 3) + "') OR "
      End If
   Next iRow
   szSQL = Left(szSQL, Len(szSQL) - 4)

   rstRecharged.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   Dim iRows As Integer
   iRows = flxRecharge.Rows - 1
   iRow = 1
   While iRow <= iRows
      If flxRecharge.TextMatrix(iRow, 0) = "X" Then
         flxRecharge.RemoveItem iRow
         iRows = iRows - 1
         iRow = iRow - 1
      End If
      iRow = iRow + 1
   Wend

   ShowMsgInTaskBar "Recharge has been done successfully", , "N"
   rstRecharged.Close
   Set rstRecharged = Nothing
   Conn1.Close
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar "The SDO generated the following error: " & Error, , "N"
   rstRecharged.Close
   Set rstRecharged = Nothing
   Conn1.Close
End Sub

Private Sub cmdNoSelect_Click()
   Dim iRow As Integer, iCol As Integer

   For iRow = 1 To flxRecharge.Rows - 1
      flxRecharge.row = iRow
      flxRecharge.TextMatrix(iRow, 0) = ""
      For iCol = 1 To flxRecharge.Cols - 1
         flxRecharge.col = iCol
         flxRecharge.CellBackColor = RGB(255, 255, 255)
         flxRecharge.TextMatrix(flxRecharge.row, 12) = "0.00"
      Next iCol
   Next iRow

   txtTotalNet.text = Format(CStr(CalculateTotal(11)), "0.00")
   txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
End Sub

Private Sub cmdRemoveCharge_Click()
   If iSelected = 0 Then
      ShowMsgInTaskBar "No Selection", , "N"
      Exit Sub
   End If
   If MsgBox("Are you sure to remove?", vbYesNo, "Remove") = vbNo Then Exit Sub

   On Error GoTo ErrorHandler

   Dim iRow As Integer
   Dim szSQL As String
   Dim Conn1 As New ADODB.Connection
   Dim Rst1 As New ADODB.Recordset

   szSQL = "DELETE * FROM tlbRechargePre WHERE "
   For iRow = 1 To flxRecharge.Rows - 1
      If flxRecharge.TextMatrix(iRow, 0) = "X" Then
         szSQL = szSQL + "(TRAN_ID= '" + flxRecharge.TextMatrix(iRow, 1) + _
                  "' AND TRAN_TYPE='" + flxRecharge.TextMatrix(iRow, 3) + "') OR "
      End If
   Next iRow
   szSQL = Left(szSQL, Len(szSQL) - 4)

   Conn1.Open getConnectionString

   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   Dim iRows As Integer
   iRows = flxRecharge.Rows - 1
   iRow = 1
   While iRow <= iRows
      If flxRecharge.TextMatrix(iRow, 0) = "X" Then
         flxRecharge.RemoveItem iRow
         iRows = iRows - 1
         iRow = iRow - 1
      End If
      iRow = iRow + 1
   Wend
   
   ShowMsgInTaskBar "Data has been removed successfully"
   Rst1.Close
   Set Rst1 = Nothing
   Conn1.Close
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar "The SDO generated the following error: " & Error, , "N"
   Rst1.Close
   Set Rst1 = Nothing
   Conn1.Close
End Sub

Private Sub cmdSave_Click()
   Dim szLandLord() As String, szSQL As String, szBackUp As String
   Dim iRow As Integer
   Dim Conn1 As New ADODB.Connection
   Dim Rst1 As New ADODB.Recordset, rstBackUp As New ADODB.Recordset

   If bData = False Then
      ShowMsgInTaskBar "No data has been saved!", , "N"
      Exit Sub
   End If

   Conn1.Open getConnectionString

   szBackUp = "DELETE * " & _
              "FROM tlbRechargePre " & _
              "WHERE UNIT_ID='" & szUnitID & "';"

   rstBackUp.Open szBackUp, Conn1, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT MY_ID, UNIT_ID, LANDLORD_ID, TRAN_ID, TRAN_DATE, " & _
               "TRAN_TYPE, TRANS, INV_NO, NOMINAL_CODE, DEPT_ID, PROJ_REF, " & _
               "COST_CODE, DESCRIPTION, NET_AMOUNT, VAT, RECHARGE, CHECKED " & _
           "FROM tlbRechargePre"
   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   szLandLord = Split(cboUnits.text, " \ ")
   For iRow = 1 To flxRecharge.Rows - 1
      Rst1.AddNew
      Rst1!My_ID = Format(Now, "yyyymmddhhmmss") & CStr(iRow)
      Rst1!UNIT_ID = szUnitID
      Rst1!LANDLORD_ID = szLandLord(0)
      If flxRecharge.TextMatrix(iRow, 0) <> "" Then Rst1!Checked = 1
      Rst1!TRAN_DATE = Format(flxRecharge.TextMatrix(iRow, 2), "DD MMMM YYYY")
      Rst1!TRANS = flxRecharge.TextMatrix(iRow, 4)
      Rst1!TRAN_TYPE = flxRecharge.TextMatrix(iRow, 3)
      Rst1!TRAN_ID = flxRecharge.TextMatrix(iRow, 1)
      Rst1!Nominal_code = flxRecharge.TextMatrix(iRow, 6)
      Rst1!DEPT_ID = flxRecharge.TextMatrix(iRow, 7)
      Rst1!PROJ_REF = flxRecharge.TextMatrix(iRow, 8)
      Rst1!COST_CODE = flxRecharge.TextMatrix(iRow, 9)
      Rst1!description = flxRecharge.TextMatrix(iRow, 10)
      Rst1!NET_AMOUNT = CCur(flxRecharge.TextMatrix(iRow, 11))
      Rst1!INV_NO = flxRecharge.TextMatrix(iRow, 5)
      Rst1!RECHARGE = CCur(flxRecharge.TextMatrix(iRow, 12))
      Rst1.Update
   Next iRow

   ShowMsgInTaskBar "Data has been saved successfully."

   Rst1.Close
   Conn1.Close

   Set Rst1 = Nothing

' UPDATE SOURCE TABLES. RECHARGED TURN TO 'YES'
   Dim szSqlPI As String, szSQLCN As String, szSqlBank As String
   Dim rstPI As New ADODB.Recordset, rstCN As New ADODB.Recordset, rstBank As New ADODB.Recordset

   Conn1.Open getConnectionString

   szSqlPI = "SELECT RECHARGED " & _
             "FROM tblPurInv " & _
             "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
   rstPI.Open szSqlPI, Conn1, adOpenDynamic, adLockOptimistic

   szSQLCN = "SELECT RECHARGED " & _
             "FROM tlbCreditNote " & _
             "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No"
   rstCN.Open szSQLCN, Conn1, adOpenDynamic, adLockOptimistic

   szSqlBank = "SELECT RECHARGED " & _
               "FROM tlbBankPayment " & _
               "WHERE UNIT_ID='" & szUnitID & "' and RECHARGED=No " & _
               "ORDER BY TRAN_TYPE"
   rstBank.Open szSqlBank, Conn1, adOpenDynamic, adLockOptimistic
   
   If rstPI.EOF And rstCN.EOF And rstBank.EOF Then GoTo ErrorHandler
      
   flxRecharge.RowHeight(1) = 285

   iRow = 1
   rstPI.MoveFirst
   While Not rstPI.EOF
      rstPI!RECHARGED = 1
      rstPI.Update
      rstPI.MoveNext
   Wend
   rstPI.Close
   Set rstPI = Nothing
   
   rstCN.MoveFirst
   While Not rstCN.EOF
      rstCN!RECHARGED = 1
      rstCN.Update
      rstCN.MoveNext
   Wend
   rstCN.Close
   Set rstCN = Nothing
   
   While Not rstBank.EOF
      rstBank!RECHARGED = 1
      rstBank.Update
      rstBank.MoveNext
   Wend
   rstBank.Close
   Set rstBank = Nothing
   Conn1.Close
'****************************************
   
   Set Conn1 = Nothing
   
   If flxRecharge.Rows > 2 Then
      For iRow = 2 To flxRecharge.Rows - 1
         flxRecharge.RemoveItem 2
      Next iRow
   End If
   For iRow = 0 To 12
      flxRecharge.TextMatrix(1, iRow) = ""
   Next iRow
   cboUnits.text = ""
   txtLandLord.text = ""
   txtTotalNet.text = "0.00"
   txtTotalRecharge.text = "0.00"
   
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar "No Data", , "N"
   bData = False
   
   If flxRecharge.Rows > 2 Then
      For iRow = 2 To flxRecharge.Rows - 1
         flxRecharge.RemoveItem 2
      Next iRow
   End If
   For iRow = 0 To 12
      flxRecharge.TextMatrix(1, iRow) = ""
   Next iRow
   cboUnits.text = ""
   txtLandLord.text = ""
   txtTotalNet.text = "0.00"
   txtTotalRecharge.text = "0.00"
End Sub

Private Sub cmdSelectAll_Click()
   Dim iRow As Integer, iCol As Integer
   
   If Not bData Then Exit Sub
   For iRow = 1 To flxRecharge.Rows - 1
      flxRecharge.row = iRow
      flxRecharge.TextMatrix(iRow, 0) = "X"
      For iCol = 1 To flxRecharge.Cols - 1
         flxRecharge.col = iCol
         flxRecharge.CellBackColor = RGB(180, 255, 180)
         flxRecharge.TextMatrix(flxRecharge.row, 12) = Format(flxRecharge.TextMatrix(flxRecharge.row, 11), "0.00")
      Next iCol
   Next iRow
   
   txtTotalNet.text = Format(CStr(CalculateTotal(11)), "0.00")
   txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
End Sub

'Private Sub dtSPDate_DateClick(ByVal DateClicked As Date)
'   txtReDate.text = Format(dtSPDate.Value, "dd/mm/yyyy")
'   dtSPDate.Visible = False
'End Sub
'
Private Sub flxRecharge_Click()
   Dim iRow As Integer
   Dim i As Integer
   Dim iFlxLeft As Integer

   iFlxLeft = flxRecharge.Left
   For i = 0 To flxRecharge.col - 1
      iFlxLeft = iFlxLeft + flxRecharge.ColWidth(i)
   Next i

'Showing the text box in the net col for editing
   If flxRecharge.col = 12 And flxRecharge.TextMatrix(flxRecharge.row, 0) = "X" Then
      txtRecharge.Top = flxRecharge.Top + (flxRecharge.RowHeight(flxRecharge.row) * flxRecharge.row) - 15
      txtRecharge.Left = iFlxLeft + 10
      txtRecharge.Width = flxRecharge.ColWidth(12)
      txtRecharge.text = "0.00"
      txtRecharge.ZOrder 0
      If flxRecharge.TextMatrix(flxRecharge.row, 12) <> "" Then txtRecharge.text = Format(flxRecharge.TextMatrix(flxRecharge.row, 12), "0.00")
      txtRecharge.SetFocus
      Exit Sub
   End If
'Select or unselect a row
   If flxRecharge.TextMatrix(flxRecharge.row, 0) = "X" Then
      flxRecharge.TextMatrix(flxRecharge.row, 0) = ""
      flxRecharge.TextMatrix(flxRecharge.row, 12) = "0.00"
      For iRow = 1 To flxRecharge.Cols - 1
         flxRecharge.col = iRow
         flxRecharge.CellBackColor = RGB(255, 255, 255)
      Next iRow
   Else
      If Not bData Then Exit Sub
      flxRecharge.TextMatrix(flxRecharge.row, 0) = "X"
      iSelected = iSelected + 1
      flxRecharge.TextMatrix(flxRecharge.row, 12) = Format(flxRecharge.TextMatrix(flxRecharge.row, 11), "0.00")

      For iRow = 1 To flxRecharge.Cols - 1
         flxRecharge.col = iRow
         flxRecharge.CellBackColor = RGB(180, 255, 180)
      Next iRow
   End If
   txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   txtRecharge.ZOrder 1
   bData = False
   iSelected = 0

   On Error GoTo ErrH1
   Me.Caption = "Recharge"

   cboUnits.text = ""

   Call LoadCboUnits

   Call ConfigFlxRecharge

   txtReDate.text = Format(Now, "dd/mm/yyyy")

   Call WheelHook(Me.hWnd)
   Exit Sub
ErrH1:
   If Err.Number = 40002 Then
      If MsgBox("Please check DSN - " & Adsn & " is set up correctly.", vbRetryCancel, "DSN Error") = vbRetry Then
          Resume
      Else
          Resume Next
      End If
   ElseIf Err.Number <> 0 Then
      ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
      Resume Next
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Public Sub LoadCboUnits()
   Dim Conn1 As New ADODB.Connection
   Dim Rst2 As New ADODB.Recordset
   Dim SQLStr2  As String

   cboUnits.Enabled = True

   'Reset screen to show all the units in cboUnits.
   Conn1.Open getConnectionString

   SQLStr2 = "SELECT U.*, L.CompanyName AS TCN " & _
             "FROM Units AS U, LeaseDetails AS L " & _
             "WHERE U.UnitNumber = L.UnitNumber " & _
             "ORDER BY U.UnitNumber"
'Debug.Print SQLStr2
   Rst2.Open SQLStr2, Conn1, adOpenStatic, adLockReadOnly

   If Rst2.EOF = False Then
      While Rst2.EOF = False
         cboUnits.AddItem Rst2!UnitNumber & " \ " & Rst2!TCN
         Rst2.MoveNext
      Wend
   End If

   Rst2.Close
   Conn1.Close

   Set Rst2 = Nothing
   Set Conn1 = Nothing
End Sub

Private Sub ConfigFlxRecharge()
   flxRecharge.ColWidth(0) = 250
   flxRecharge.ColWidth(1) = 600
   flxRecharge.TextMatrix(0, 1) = "No."
   
   flxRecharge.ColWidth(2) = 1100
   flxRecharge.TextMatrix(0, 2) = "Date"
   
   flxRecharge.ColWidth(3) = 1000
   flxRecharge.TextMatrix(0, 3) = "Type"
   
   flxRecharge.ColWidth(4) = 1100
   flxRecharge.TextMatrix(0, 4) = "Trans"
   
   flxRecharge.ColWidth(5) = 900
   flxRecharge.TextMatrix(0, 5) = "Inv No"
   
   flxRecharge.ColWidth(6) = 900
   flxRecharge.TextMatrix(0, 6) = "N/C"
   
   flxRecharge.ColWidth(7) = 700
   flxRecharge.TextMatrix(0, 7) = "Dept"
   
   flxRecharge.ColWidth(8) = 700
   flxRecharge.TextMatrix(0, 8) = "Proj"
   
   flxRecharge.ColWidth(9) = 900
   flxRecharge.TextMatrix(0, 9) = "C Code"

   flxRecharge.ColWidth(10) = 1900
   flxRecharge.TextMatrix(0, 10) = "Details"

   flxRecharge.ColWidth(11) = 1100
   flxRecharge.TextMatrix(0, 11) = "Net"
   txtTotalNet.Width = 1100

   flxRecharge.ColWidth(12) = 1100
   flxRecharge.TextMatrix(0, 12) = "Recharge"
   txtTotalRecharge.Width = 1100

   flxRecharge.RowHeight(1) = 285
End Sub

Private Sub txtRecharge_GotFocus()
   SelTxtInCtrl txtRecharge
End Sub

Private Sub txtRecharge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      cmdSave.SetFocus
      Exit Sub
   End If

   If KeyAscii = 13 Or KeyAscii = 10 Then cmdSave.SetFocus

   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtRecharge_LostFocus()
   'Write the value in the FlexGrid
   If txtRecharge.text = "0.00" Then
      flxRecharge.TextMatrix(flxRecharge.row, 12) = "0.00"
   Else
      If CCur(flxRecharge.TextMatrix(flxRecharge.row, 11)) < CCur(txtRecharge.text) Then
         If MsgBox("Recharge amount is bigger than Net amount. Are you sure?", vbYesNo, "RECHARGE") = vbYes Then
            flxRecharge.TextMatrix(flxRecharge.row, 12) = Format(txtRecharge.text, "0.00")
         End If
      End If
   End If
   txtRecharge.ZOrder 1
   txtTotalRecharge.text = Format(CStr(CalculateTotal(12)), "0.00")
End Sub

Private Sub txtReDate_Change()
   TextBoxChangeDate txtReDate
End Sub

Private Sub txtReDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtReDate, KeyAscii
End Sub

Private Sub txtReDate_LostFocus()
   TextBoxFormatDate txtReDate
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
