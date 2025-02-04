VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmServiceCharge11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Charge Budget"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   Icon            =   "frmServiceCharge11.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11925
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   345
      Left            =   10035
      TabIndex        =   30
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
      Height          =   345
      Left            =   6480
      TabIndex        =   23
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdNew 
      Caption         =   "&New"
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdEdit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   1450
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   2825
      TabIndex        =   6
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSCBdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   8760
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdClose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   10140
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Index           =   1
      Left            =   75
      TabIndex        =   10
      Top             =   915
      Width           =   11160
      Begin VB.TextBox txtBudgetId 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "Budget ID"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMatrixRow 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Text            =   "Matrix Row"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDetails 
         Caption         =   ">>"
         Height          =   315
         Left            =   6400
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtBudget 
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
         Height          =   295
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   375
         Width           =   2200
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
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
         Left            =   6960
         TabIndex        =   3
         Top             =   360
         Width           =   1905
      End
      Begin VB.TextBox txtSCTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5055
         Width           =   1935
      End
      Begin VB.TextBox txtSCBudgetTotal 
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
         Left            =   4320
         TabIndex        =   11
         Top             =   5055
         Width           =   2715
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetails 
         Height          =   4305
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   7594
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
         Caption         =   "Total Area"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price/SqFoot"
         Height          =   195
         Index           =   4
         Left            =   8880
         TabIndex        =   18
         Top             =   120
         Width           =   945
      End
      Begin MSForms.ComboBox cboSCFund 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
         VariousPropertyBits=   1755334683
         DisplayStyle    =   7
         Size            =   "7223;556"
         TextColumn      =   2
         ColumnCount     =   6
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;70555"
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   17
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   15
         Top             =   5055
         Width           =   390
      End
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   3
      Left            =   6465
      TabIndex        =   28
      Top             =   135
      Width           =   630
   End
   Begin MSForms.ComboBox cboProperty 
      Height          =   315
      Left            =   7275
      TabIndex        =   27
      Top             =   120
      Width           =   3975
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7011;556"
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
   Begin MSForms.ComboBox cboClient 
      Height          =   315
      Left            =   1275
      TabIndex        =   26
      Top             =   120
      Width           =   3975
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7011;556"
      TextColumn      =   2
      ColumnCount     =   5
      ListRows        =   20
      cColumnInfo     =   5
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1411;3527;0;0;0"
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Year:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   930
   End
   Begin MSForms.ComboBox cboBudgetYears 
      Height          =   315
      Left            =   1275
      TabIndex        =   24
      Top             =   480
      Width           =   3975
      VariousPropertyBits=   1755334683
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7011;556"
      TextColumn      =   2
      ColumnCount     =   3
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
End
Attribute VB_Name = "frmServiceCharge11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private bFormLoaded        As Boolean
Dim adoConn                As New ADODB.Connection
Dim szSQL                  As String
Dim flgEdit                As Integer
Dim flgNew                 As Integer
Dim nextRow                As Integer

Dim flgChange              As Integer
Dim detailsMatrix(59, 59)  As clsSCDtl
Dim matrixLength           As Integer
Dim bufferMatrix()         As clsSCDtl


Public Function fillBufferMatrix(ByVal bDetails As clsSCDtl, ByVal num As Integer)
   ReDim Preserve bufferMatrix(num) As clsSCDtl

   Set bufferMatrix(num - 1) = New clsSCDtl

   bufferMatrix(num - 1).setBudgetAmount bDetails.getBudgetAmount
   bufferMatrix(num - 1).setBudgetDetailId bDetails.getBudgetDetailID
   bufferMatrix(num - 1).setBudgetId bDetails.getBudgetId
   bufferMatrix(num - 1).setFlgDel bDetails.getFlgDel
   bufferMatrix(num - 1).setNCode bDetails.getNCode
   bufferMatrix(num - 1).setNName bDetails.getNName
End Function

Private Sub cboBudgetYears_Change()
   Dim iRow As Integer

   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If cboBudgetYears.Value = flxSCBudgetDetails.TextMatrix(iRow, 9) And _
            cboProperty.Value = flxSCBudgetDetails.TextMatrix(iRow, 1) Then
         flxSCBudgetDetails.RowHeight(iRow) = 240
      Else
         flxSCBudgetDetails.RowHeight(iRow) = 0
      End If
   Next iRow

   SCSumTotal
End Sub

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
   TotalCol = adoRst.Fields.count - 1

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
           "WHERE ClientID = '" & cboClient.Column(0) & "' " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

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

Private Sub cboClient_Change()
   If Not bFormLoaded Then Exit Sub
   If cboClient.text = "" Then Exit Sub

'   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   On Error GoTo ErrorHandler

   adoConn.Open getConnectionString

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & cboClient.Column(0) & "' " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

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
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

   LoadFY

NoRes:
   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
'   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoConn.Close
   Set adoRst = Nothing
'   Set adoConn = Nothing
End Sub

Private Sub cboProperty_Change()
'   adoConn.Open getConnectionString

'   LoadFlxSCBudgetDetails

'   adoConn.Close
End Sub

Private Sub cboSCFund_Click()
   If Not txtBudget.Locked Then
      txtBudget.SetFocus
   Else
      If Not txtTotalArea.Locked Then
         frmServiceCharge.Show
         txtTotalArea.SetFocus
      End If
   End If
End Sub

Private Sub cmdDetails_Click()
   Load frmServiceChargeDetails1
   frmServiceChargeDetails1.bNewBudget = IIf(flgNew = 1, True, False)
   frmServiceChargeDetails1.Show
   Me.Enabled = False
End Sub

Public Function initialiseMatrix() As Integer
   Dim i As Integer, j As Integer

   For i = 0 To 59
      For j = 0 To 59
         Set detailsMatrix(i, j) = New clsSCDtl
      Next j
   Next i
End Function

Public Function getDetailsFromMatrix(ByVal row As Integer, ByVal col As Integer) As clsSCDtl
   Set getDetailsFromMatrix = detailsMatrix(row, col)
End Function

Private Sub cmdImport_Click()
   If cboClient.ListIndex = -1 Then
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      cboClient.SetFocus
      Exit Sub
   End If
   If cboBudgetYears.ListIndex = -1 Then
      ShowMsgInTaskBar "Please select a budget year", "Y", "N"
      cboBudgetYears.SetFocus
      Exit Sub
   End If
   If cboProperty.ListIndex = -1 Then
      ShowMsgInTaskBar "Please select a property", "Y", "N"
      cboProperty.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you wish to import a Service Charge budget?", vbQuestion + vbYesNo, "Import Service Charge Budget") = vbNo Then Exit Sub

   Dim ofn                    As OPENFILENAME
   Dim lHwnd                  As Long
   Const HKEY_LOCAL_MACHINE   As Long = &H80000002
   Dim szOldFile_PathName     As String
   Dim szNewFile_Path         As String
   Dim szNewFile_Name         As String
   Dim szNewFile_PathName     As String
   Dim fso                    As Object
   Dim szImportFile           As String
   Dim szNC                   As String
   Dim szFund                 As String

   On Error GoTo FileError

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "MS Office Excel Workbook 2007-2010 (*.xlsx)" + Chr$(0) + "*.xlsx" + Chr$(0) + _
                     "MS Office Excel Workbook 97-2003 (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + _
                     "CSV Files (*.csv)" + Chr$(0) + "*.csv" + Chr$(0)

   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir
   ofn.lpstrTitle = "Select File to Save"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   szImportFile = ofn.lpstrFile

   If szImportFile = "" Then
      ShowMsgInTaskBar "Please select an input file to import", "Y", "N"
      cmdImport.SetFocus
      Exit Sub
   End If

'  Get the list of nominal code
   Dim adoRst As New Recordset

   adoConn.Open getConnectionString

   adoRst.Open "SELECT Code & ' # '  & Name FROM NominalLedger WHERE ClientID = '" & cboClient.Value & "';", adoConn, adOpenStatic, adLockReadOnly
   szNC = SQL2String(adoRst, 0)
   adoRst.Close

   adoRst.Open "SELECT FundName & ' # '  & FundID FROM Fund;", adoConn, adOpenStatic, adLockReadOnly
   szFund = SQL2String(adoRst, 0)
   adoRst.Close

   Dim oXL As New Excel.Application
   Dim oWB As Workbook
   Dim oWS As Worksheet

   Set oWB = oXL.Workbooks.Open(szImportFile)
   Set oWS = oWB.Worksheets(1) 'Specify your worksheet name

   Dim iRow             As Integer
   Dim szNC_NotExist    As String
   Dim szFund_NotExist  As String
   Dim szBudgetID       As String

   If UCase(oWS.Range("A1").Value) <> "FUND" Then
      ShowMsgInTaskBar "The file format is wrong", "Y", "N"

      oWB.Close
      oXL.Quit

      Set oWS = Nothing
      Set oWB = Nothing
      Set oXL = Nothing
      Exit Sub
   End If

   iRow = 2
   While oWS.Range("A" & CStr(iRow)).Value <> vbEmpty
      If InStr(UCase(szFund), UCase(oWS.Range("A" & iRow).Value)) <= 0 Then
         szFund_NotExist = szFund_NotExist & oWS.Range("A" & iRow).Value & ", "
      Else
         If InStr(szNC, oWS.Range("B" & iRow).Value) <= 0 Then
            szNC_NotExist = szNC_NotExist & oWS.Range("B" & iRow).Value & ", "
         Else
            szBudgetID = getBudgetId(cboProperty.Value, oWS.Range("A" & iRow).Value, cboBudgetYears.Value)
            If szBudgetID <> "NULL" Then
               szSQL = "SELECT C.* " & _
                       "FROM GlobalSCDtls AS C INNER JOIN GlobalSC AS P ON P.BudgetID = C.BudgetID " & _
                       "WHERE C.BudgetID = '" & szBudgetID & "' AND " & _
                             "P.FinancialYear = '" & cboBudgetYears.Value & "' AND " & _
                             "C.NC = '" & oWS.Range("B" & iRow).Value & "';"
               adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
               
               If Not adoRst.EOF Then
                  adoRst.Fields.Item("BudgetAmt").Value = oWS.Range("C" & iRow).Value
               Else
                  adoRst.AddNew
                  adoRst.Fields.Item("BudgetDtlID").Value = UniqueID()
                  adoRst.Fields.Item("BudgetID").Value = szBudgetID
                  adoRst.Fields.Item("NC").Value = oWS.Range("B" & iRow).Value
                  adoRst.Fields.Item("NN").Value = GetNN(szNC, oWS.Range("B" & iRow).Value)
                  adoRst.Fields.Item("BudgetAmt").Value = oWS.Range("C" & iRow).Value
               End If
               adoRst.Update
               adoRst.Close
            Else
               szBudgetID = UniqueID()
               adoRst.Open "SELECT * FROM GlobalSC;", adoConn, adOpenDynamic, adLockOptimistic
               adoRst.AddNew
               adoRst.Fields.Item("BudgetID").Value = szBudgetID
               adoRst.Fields.Item("PropertyID").Value = cboProperty.Value
               adoRst.Fields.Item("Fund").Value = GetFundID(szFund, oWS.Range("A" & iRow).Value)
               adoRst.Fields.Item("SCArea").Value = 1
               adoRst.Fields.Item("FinancialYear").Value = cboBudgetYears.Value
               adoRst.Update
               adoRst.Close

               adoRst.Open "SELECT * FROM GlobalSCDtls;", adoConn, adOpenDynamic, adLockOptimistic
               adoRst.AddNew
               adoRst.Fields.Item("BudgetDtlID").Value = UniqueID()
               adoRst.Fields.Item("BudgetID").Value = szBudgetID
               adoRst.Fields.Item("NC").Value = oWS.Range("B" & iRow).Value
               adoRst.Fields.Item("NN").Value = GetNN(szNC, oWS.Range("B" & iRow).Value)
               adoRst.Fields.Item("BudgetAmt").Value = oWS.Range("C" & iRow).Value
               adoRst.Update
               adoRst.Close
            End If
         End If
      End If
      iRow = iRow + 1
   Wend

   oWB.Close
   oXL.Quit

   Set oWS = Nothing
   Set oWB = Nothing
   Set oXL = Nothing

   Dim adoChild      As New ADODB.Recordset

   adoRst.Open "SELECT * FROM GlobalSC;", adoConn, adOpenDynamic, adLockOptimistic

   While Not adoRst.EOF
      szSQL = "SELECT SUM(C.BudgetAmt) " & _
              "FROM GlobalSCDtls AS C " & _
              "WHERE C.BudgetID = '" & adoRst.Fields.Item("BudgetID").Value & "';"

      adoChild.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      
      If adoRst.Fields.Item("TotalBudget").Value <> Val(adoChild.Fields.Item(0).Value) Then
         adoRst.Fields.Item("TotalBudget").Value = Val(adoChild.Fields.Item(0).Value)
         adoRst.Fields.Item("PPSF").Value = adoRst.Fields.Item("TotalBudget").Value
         adoRst.Update
      End If
      
      adoChild.Close
      adoRst.MoveNext
   Wend

   Set adoRst = Nothing
   Set adoChild = Nothing

   ConfigFlxSCBudgetDetails
   LoadFlxSCBudgetDetails
   LoadMatrix

   adoConn.Close

   SCSumTotal
   cboBudgetYears.ListIndex = -1
   cboBudgetYears.ListIndex = 0
   
   If szNC_NotExist = "" And szFund_NotExist = "" Then
      ShowMsgInTaskBar "Import has been done successfully", "Y", "P"
   Else
      If szNC_NotExist <> "" Then
         ShowMsgInTaskBar "Nominal Codes do not found", "Y", "N"
         szNC_NotExist = Mid(Trim(szNC_NotExist), 1, Len(Trim(szNC_NotExist)) - 1)
      Else
         szNC_NotExist = "NULL"
      End If
      If szFund_NotExist <> "" Then
         ShowMsgInTaskBar "Funds do not found", "Y", "N"
         szFund_NotExist = Mid(Trim(szFund_NotExist), 1, Len(Trim(szFund_NotExist)) - 1)
      Else
         szFund_NotExist = "NULL"
      End If

      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SC_Imp_Report.rpt")

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue szNC_NotExist
      Report.ParameterFields(2).AddCurrentValue szFund_NotExist

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   Exit Sub
FileError:
   ShowMsgInTaskBar "File not does not exists", "Y", "N"
End Sub

Private Function GetNN(szNCList As String, szNC As String) As String
   Dim szaTemp() As String
   Dim szaTem1() As String
   Dim i         As Integer

   szaTemp = Split(szNCList, ", ")

   For i = 0 To UBound(szaTemp)
      If InStr(szaTemp(i), szNC) > 0 Then
         szaTem1 = Split(szaTemp(i), " # ")
         GetNN = szaTem1(1)
         Exit Function
      End If
   Next i
   GetNN = "NULL"
End Function

Private Function GetFundID(szFundList As String, szFundName As String) As Single
   Dim szaTemp() As String
   Dim szaTem1() As String
   Dim i         As Integer

   szaTemp = Split(szFundList, ", ")

   For i = 0 To UBound(szaTemp)
      If InStr(szaTemp(i), szFundName) > 0 Then
         szaTem1 = Split(szaTemp(i), " # ")
         GetFundID = Val(szaTem1(1))
         Exit Function
      End If
   Next i
   GetFundID = 0
End Function

Private Function getBudgetId(szPropID As String, szFundName As String, szFY As String) As String
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

   szSQL = "SELECT BudgetID " & _
           "FROM   GlobalSC AS G, Fund AS F " & _
           "WHERE  G.Fund = F.FundID AND " & _
                  "G.PropertyID = '" & szPropID & "' AND " & _
                  "F.FundName = '" & szFundName & "' AND " & _
                  "G.FinancialYear = '" & szFY & "';"
                  
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      getBudgetId = adoRst.Fields.Item(0).Value
   Else
      getBudgetId = "NULL"
   End If

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub cmdSCBdCancel_Click()
   If cmdSCBdEdit.Caption = "&Update" Then
      If MsgBox("Do you want to cancel the operation?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
         rollBackServiceChargeDetails
         ReDim bufferMatrix(0) As clsSCDtl
         ControlsModeRentBudgetDetails DefaultMode
         Exit Sub
      End If
   End If

   If MsgBox("Are you sure to cancel the changes made since the last Save?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
      flxSCBudgetDetails.Clear
      initialiseGrid
'      If frmMMain.IsRibbonVersion Then
      cboBudgetYears.ListIndex = 0
'      End If
      ControlsModeRentBudgetDetails DefaultMode
      SCSumTotal
   End If
End Sub

Private Sub rollBackServiceChargeDetails()
   Dim i As Integer

   For i = 0 To 59
      If detailsMatrix(txtMatrixRow.text, i).getBudgetId <> "" Then
         Set detailsMatrix(txtMatrixRow.text, i) = New clsSCDtl
      Else
         Exit For
      End If
   Next i

   If Not IsEmpty(bufferMatrix) Then
      For i = 0 To UBound(bufferMatrix) - 1
         Set detailsMatrix(txtMatrixRow.text, i) = bufferMatrix(i)
      Next i
   End If
End Sub

Private Sub cmdSCBdClose_Click()
   If flgChange = 1 Then
      If MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNo, "New Budget") = vbYes Then
         cmdSCBDSave_Click
      End If
   End If

   Me.Hide
   Unload Me
End Sub

Private Sub cmdSCBdDelete_Click()
   If MsgBox("Do you want to delete the selected Rent Charge Budget Detail?", vbQuestion + vbYesNo, "Saving") = vbNo Then
         Exit Sub
   End If
   If Trim(flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)) <> "" Then
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X"
      flxSCBudgetDetails.RowHeight(flxSCBudgetDetails.row) = 0
   Else
      flxSCBudgetDetails.RemoveItem (flxSCBudgetDetails.row)
   End If
   flgChange = 1
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub cmdSCBDEdit_Click()
   Dim col As Integer

   If cmdSCBdEdit.Caption = "&Edit" Then
      flgEdit = 1
      flgNew = 0
      ControlsModeRentBudgetDetails EditMode

      If hasNoLines Then
         cmdDetails.Enabled = False
      Else
         cmdDetails.Enabled = True
         txtBudget.Locked = True
      End If

      cboSCFund.SetFocus

      For col = 0 To 59
         If getDetailsFromMatrix(txtMatrixRow.text, col).getBudgetDetailID = "" Then
            Exit For
         Else
            fillBufferMatrix getDetailsFromMatrix(txtMatrixRow.text, col), col + 1
         End If
      Next col
   Else
      If cboSCFund.Value = "" Then
         ShowMsgInTaskBar "Please select the fund.", , "N"
         cboSCFund.SetFocus
         Exit Sub
      End If
      If txtBudget.text = "" Then
         ShowMsgInTaskBar "Please enter the total budget.", , "N"
         txtBudget.SetFocus
         Exit Sub
      End If
      If txtTotalArea.text = "" Then
         ShowMsgInTaskBar "Please enter the total area.", , "N"
         txtTotalArea.SetFocus
         Exit Sub
      End If

      updateGrid

      ReDim bufferMatrix(0) As clsSCDtl
      flgChange = 1
      If flgNew = 1 Then
         nextRow = nextRow + 1
         flgNew = 0
      End If

      SCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
   End If
End Sub

Private Sub cmdSCBdNew_Click()
   If cboBudgetYears.text = "" Then
      ShowMsgInTaskBar "Please select a financial year", "Y", "N"
      cboBudgetYears.SetFocus
      Exit Sub
   End If

   flgChange = 1
   txtMatrixRow.text = nextRow
   txtBudgetId.text = UniqueID()
   flgNew = 1
   flgEdit = 0
   ControlsModeRentBudgetDetails NewEntryMode

'   If MsgBox("Do you wish to analyse the budget?", vbQuestion + vbYesNo, "Analyse Budget") = vbYes Then
   txtBudget.Locked = True
   cmdDetails_Click
'   Else
'      cmdDetails.Enabled = False
'      txtBudget.Locked = False
'      cboSCFund.SetFocus
'   End If
End Sub

Private Sub cmdSCBDSave_Click()
   Dim i          As Integer
   Dim iRow       As Integer
   Dim iRowChild  As Integer
   Dim Rst        As New ADODB.Recordset

   adoConn.Open getConnectionString

   i = 0
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(iRow, 7) <> "X" Then
         SaveUpdateSC iRow
         txtBudgetId.text = flxSCBudgetDetails.TextMatrix(iRow, 0)
         saveSCDetails i
      Else
         deleteSC flxSCBudgetDetails.TextMatrix(iRow, 0)
      End If
   i = i + 1
   Next iRow

   Update_SC_Lease

   LoadMatrix

'   If frmMMain.IsRibbonVersion Then
'   xfrmGlobalx.cmdYearlyService.Caption = Format(TotalSCProperty(frmGlobalx.cboProperty.value), "£0.00")
'   Else
'      frmGlobal1.lblSCTotal.Caption = Format(TotalSCProperty(frmGlobal1.cboProperty.value), "£0.00")
'   End If

   Set Rst = Nothing
   adoConn.Close
'   Set adoConn = Nothing
   flgEdit = 0
   flgNew = 0
   flgChange = 0

   ControlsModeRentBudgetDetails SavedMode
End Sub

Private Sub Update_SC_Lease()
   Dim Rst     As New ADODB.Recordset
   Dim adoRst  As New ADODB.Recordset

   szSQL = "SELECT * FROM LServiceCharges WHERE ChargingMethod = 4 OR ChargingMethod = 2;"

   With Rst
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      While Not .EOF
         If .Fields.Item("ChargingMethod").Value = 2 Then
            szSQL = "SELECT GSC.TotalBudget, F.PartOfYear " & _
                    "FROM   GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
                        "Frequencies AS F, LServiceCharges AS LSC " & _
                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
                        "L.UnitNumber = U.UnitNumber AND " & _
                        "L.LeaseID = LSC.LeaseID AND " & _
                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
                        "LSC.SCFrequency = F.ID AND " & _
                        "U.PropertyID = GSC.PropertyID;"
'Debug.Print szSQL
            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            
            If Not adoRst.EOF Then
               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("TotalBudget").Value * _
                                                .Fields.Item("CMFigure").Value / 100
               .Fields.Item("SCAmount").Value = (adoRst.Fields.Item("TotalBudget").Value * _
                                                (.Fields.Item("CMFigure").Value / 100)) / _
                                                adoRst.Fields.Item("PartOfYear").Value
            Else
               .Fields.Item("SCTotal").Value = 0
               .Fields.Item("SCAmount").Value = 0
            End If
            adoRst.Close
         Else                             '.ChargingMethod = 4
            szSQL = "SELECT GSC.PPSF, U.TotalArea, F.PartOfYear " & _
                    "FROM GlobalSC AS GSC, LeaseDetails AS L, Units AS U " & _
                        "Frequencies AS F, LServiceCharges AS LSC " & _
                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
                        "L.UnitNumber = U.UnitNumber AND " & _
                        "L.LeaseID = LSC.LeaseID AND " & _
                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
                        "LSC.SCFrequency = F.ID AND " & _
                        "U.PropertyID = GSC.PropertyID;"
'Debug.Print szSQL
            If Not adoRst.EOF Then
               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("PPSF").Value * _
                                               adoRst.Fields.Item("TotalArea").Value
               .Fields.Item("SCAmount").Value = adoRst.Fields.Item("PPSF").Value * _
                                                adoRst.Fields.Item("TotalArea").Value / _
                                                adoRst.Fields.Item("PartOfYear").Value
               .Fields.Item("CMFigure").Value = adoRst.Fields.Item("PPSF").Value * _
                                                adoRst.Fields.Item("TotalArea").Value

            Else
               .Fields.Item("SCTotal").Value = 0
               .Fields.Item("SCAmount").Value = 0
               .Fields.Item("CMFigure").Value = 0
            End If
            adoRst.Close
         End If
         .Update
         .MoveNext
      Wend
      .Close
   End With

   Set adoRst = Nothing
   Set Rst = Nothing
End Sub

Private Function saveSCDetails(ByVal row As Integer)
   Dim col As Integer

   For col = 0 To 59
      If frmServiceCharge.getDetailsFromMatrix(row, col).getBudgetDetailID = "" Then
          Exit For
      Else
         If frmServiceCharge.getDetailsFromMatrix(row, col).getFlgDel = "D" Then
            deleteSCD frmServiceCharge.getDetailsFromMatrix(row, col)
         Else
            SaveUpdateSCD frmServiceCharge.getDetailsFromMatrix(row, col)
         End If
      End If
   Next col
End Function

Private Function SaveUpdateSCD(ByVal budgetDetails As clsSCDtl)
   Dim Rst     As New ADODB.Recordset

   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   Rst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT * " & _
            "FROM GlobalSCDtls;"
   Rst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   Rst.AddNew
   Rst!BudgetDtlID = budgetDetails.getBudgetDetailID
   Rst!budgetId = budgetDetails.getBudgetId
   Rst!NN = budgetDetails.getNName
   Rst!NC = budgetDetails.getNCode
   Rst!BudgetAmt = budgetDetails.getBudgetAmount

   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSCD(ByVal budgetDetails As clsSCDtl)
   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   adoConn.Execute szSQL
End Function

Private Function SaveUpdateSC(ByVal Index As Integer)
   Dim Rst     As New ADODB.Recordset

   szSQL = "SELECT * " & _
            "FROM GlobalSC where budgetId='" & flxSCBudgetDetails.TextMatrix(Index, 0) & "';"
   Rst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   If Rst.RecordCount < 1 Then
      Rst.AddNew
      Rst!budgetId = flxSCBudgetDetails.TextMatrix(Index, 0)
      If frmMMain.IsRibbonVersion Then Rst!FinancialYear = cboBudgetYears.Value
      Rst!propertyID = cboProperty.Value
   Else
      Rst!propertyID = flxSCBudgetDetails.TextMatrix(Index, 1)
   End If

'   If frmMMain.IsRibbonVersion Then
'   Rst!PropertyID = cboProperty.Value
   
'   Else
'      Rst!PropertyID = frmGlobal1.cboProperty.value
'   End If
   Rst!Fund = flxSCBudgetDetails.TextMatrix(Index, 2)
   Rst!TotalBudget = flxSCBudgetDetails.TextMatrix(Index, 4)
   Rst!SCArea = flxSCBudgetDetails.TextMatrix(Index, 5)
   Rst!ppsf = flxSCBudgetDetails.TextMatrix(Index, 6)

   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSC(ByVal bId As String)
   szSQL = "DELETE * " & _
            "FROM GlobalSC " & _
            "WHERE BudgetID = '" & bId & "';"
   adoConn.Execute szSQL

'   szSQL = "DELETE GlobalSCDtls.* " & _
'           "FROM   GlobalSCDtls " & _
'           "WHERE  GlobalSCDtls.BudgetID NOT IN (" & _
'               "SELECT GlobalSC.BudgetID " & _
'               "FROM GlobalSC);"
    szSQL = "DELETE GlobalSCDtls.* From GlobalSCDtls LEFT JOIN GlobalSC ON GlobalSCDtls.BudgetID =GlobalSC.BudgetID WHERE  GlobalSC.BudgetID  IS NULL;"
   adoConn.Execute szSQL
End Function

Private Sub cmdSCClose_Click()
   initialiseGrid
   Me.Hide
End Sub

Private Sub ConfigFlxSCBudgetDetails()
   Dim szFlxHeader As String

   flxSCBudgetDetails.Rows = 1
   flxSCBudgetDetails.RowHeight(0) = 0
   flxSCBudgetDetails.Clear
   flxSCBudgetDetails.Cols = 10
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF|FY"
   flxSCBudgetDetails.FormatString = szFlxHeader$

   flxSCBudgetDetails.ColWidth(0) = 0
   flxSCBudgetDetails.ColWidth(1) = 0
   flxSCBudgetDetails.ColWidth(2) = 0
   flxSCBudgetDetails.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(0).Left
   flxSCBudgetDetails.ColWidth(4) = lblRentCharges(3).Left - lblRentCharges(2).Left
   flxSCBudgetDetails.ColWidth(5) = lblRentCharges(4).Left - lblRentCharges(3).Left
   flxSCBudgetDetails.ColWidth(6) = flxSCBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxSCBudgetDetails.ColWidth(7) = 0
   flxSCBudgetDetails.ColWidth(8) = 0
   flxSCBudgetDetails.ColWidth(9) = 0
End Sub

Private Sub LoadFY()
   Dim rRow    As Integer
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

'   If frmMMain.IsRibbonVersion Then
   szSQL = "SELECT FYrID, FinancialYear, FY_Description " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.PropertyID = '" & cboProperty.Value & "';"
'   Else
'      szSQL = "SELECT FYrID, FinancialYear, FY_Description " & _
'              "FROM   FinancialYear AS F, Property AS P " & _
'              "WHERE  F.ClientID = P.ClientID AND " & _
'                     "P.PropertyID = '" & frmGlobal1.cboProperty.value & "';"
'   End If

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Financial year has not created.", "Y", "N"
   Else
      ReDim Data(2, adoRst.RecordCount) As String

      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("FYrID").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("FinancialYear").Value)
         Data(2, rRow) = Trim(adoRst.Fields.Item("FY_Description").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend

      cboBudgetYears.Clear
      cboBudgetYears.Column() = Data()
      cboBudgetYears.ListIndex = -1
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading financial year.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub LoadFlxSCBudgetDetails()
   Dim i       As Integer
   Dim Rst     As New ADODB.Recordset

   szSQL = "SELECT g.*,f.FundName " & _
            "FROM GlobalSC g, Fund f " & _
            "WHERE CInt(g.Fund)=f.FundId;"
   Rst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   If Not Rst.EOF Then
      While Not Rst.EOF
         flxSCBudgetDetails.AddItem ""
         flxSCBudgetDetails.TextMatrix(i, 0) = Rst!budgetId
         flxSCBudgetDetails.TextMatrix(i, 1) = Rst!propertyID
         flxSCBudgetDetails.TextMatrix(i, 2) = Rst!Fund
         flxSCBudgetDetails.TextMatrix(i, 3) = Rst!FundName
         flxSCBudgetDetails.TextMatrix(i, 4) = Format(Rst!TotalBudget, "0.00")
         flxSCBudgetDetails.TextMatrix(i, 5) = Rst!SCArea
         flxSCBudgetDetails.TextMatrix(i, 6) = Format(Rst!ppsf, "0.00")
         flxSCBudgetDetails.TextMatrix(i, 8) = i - 1
         flxSCBudgetDetails.TextMatrix(i, 9) = IIf(IsNull(Rst!FinancialYear), "", Rst!FinancialYear)
         flxSCBudgetDetails.RowHeight(i) = 0
         i = i + 1
         Rst.MoveNext
      Wend
   End If
   nextRow = i - 1
   initialiseMatrix
   flxSCBudgetDetails.row = 0
   flxSCBudgetDetails.col = 0

   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub LoadMatrix()
   Dim i As Integer

   For i = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(i, 0) <> "" Then
         PopulateMatrix flxSCBudgetDetails.TextMatrix(i, 0), flxSCBudgetDetails.TextMatrix(i, 8)
      End If
   Next i
End Sub

Private Sub PopulateMatrix(bId As String, row As Integer)
   Dim i    As Integer
   Dim Rst  As New ADODB.Recordset

   szSQL = "SELECT g.* " & _
            "FROM  GlobalSCDtls AS g " & _
            "WHERE g.BudgetID = '" & bId & "';"
   Rst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 0
   If Not Rst.EOF Then
      While Not Rst.EOF
         getDetailsFromMatrix(row, i).setBudgetDetailId Rst!BudgetDtlID
         getDetailsFromMatrix(row, i).setBudgetId Rst!budgetId
         getDetailsFromMatrix(row, i).setNCode Rst!NC
         getDetailsFromMatrix(row, i).setNName IIf(IsNull(Rst!NN), "", Rst!NN)
         getDetailsFromMatrix(row, i).setBudgetAmount Format(Rst!BudgetAmt, "0.00")
         i = i + 1
         Rst.MoveNext
      Wend
   End If

   Rst.Close
   Set Rst = Nothing
End Sub

Private Sub SCSumTotal()
   Dim iRow As Integer

   txtSCBudgetTotal.text = "0.00"
   txtSCTotalArea.text = "0"
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.RowHeight(iRow) > 0 Then
         txtSCBudgetTotal.text = Format(Val(txtSCBudgetTotal.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 4)), "0.00")
         txtSCTotalArea.text = Val(txtSCTotalArea.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 5))
      End If
   Next iRow
End Sub

Private Sub Command1_Click()
   cboBudgetYears_Change
End Sub

Private Sub flxSCBudgetDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub flxSCBudgetDetails_RowColChange()
   If flxSCBudgetDetails.TextMatrix(1, 0) = "" Then Exit Sub
   If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X" Then Exit Sub

   On Error Resume Next

   cboSCFund.Value = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2)
   txtBudget.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4)
   txtTotalArea.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5)
   txtPpsf.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6)
   txtMatrixRow.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 8)
   txtBudgetId.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)

   ControlsModeRentBudgetDetails GridRowOnSelection

   cmdSCBdEdit.SetFocus
End Sub

Private Function hasNoLines() As Boolean
   Dim i As Integer

   hasNoLines = True
   i = 0
   While i < 60
      If detailsMatrix(flxSCBudgetDetails.row - 1, i).getBudgetId <> "" Then
         If detailsMatrix(flxSCBudgetDetails.row - 1, i).getFlgDel <> "X" Then
            hasNoLines = False
            Exit Function
         End If
      Else
         Exit Function
      End If
      i = i + 1
   Wend
End Function

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   On Error Resume Next

   Select Case mode
      Case ComponentMode.DefaultMode
         cboSCFund.text = ""
         cboSCFund.Enabled = False
         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = False
         cboBudgetYears.Enabled = True
         cmdSCBdEdit.Caption = "&Edit"
         cmdSCBdEdit.ToolTipText = "Edit selected budget"
         If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
            cmdSCBdCancel.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
            cmdSCBdCancel.Enabled = False
         End If
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdNew.SetFocus

      Case ComponentMode.SavedMode
         cboSCFund.text = ""
         cboSCFund.Enabled = False
         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = False
         cboBudgetYears.Enabled = True
         cmdSCBdEdit.Caption = "&Edit"
         cmdSCBdEdit.ToolTipText = "Edit selected budget"
         If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
            cmdSCBdCancel.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
            cmdSCBdCancel.Enabled = False
         End If

         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdClose.SetFocus

      Case ComponentMode.EditMode

         cboSCFund.Enabled = True
         txtTotalArea.Locked = False
         txtBudget.Locked = False
         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
         cmdSCBdEdit.Caption = "&Update"
         cmdSCBdEdit.ToolTipText = "Update selected budget"
         cmdSCBdEdit.Enabled = True
         cmdSCBdSave.Enabled = False
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.NewEntryMode
         cboSCFund.text = ""
         cboSCFund.Enabled = True
         txtBudget.text = ""
         txtTotalArea.text = ""
         txtTotalArea.Locked = False
         txtPpsf.text = ""

         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
         cmdSCBdEdit.Caption = "&Update"
         cmdSCBdEdit.Enabled = True
         cboBudgetYears.Enabled = False
         cmdSCBdSave.Enabled = False
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False
         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.GridRowOnSelection
         txtBudget.Locked = True
         cmdSCBdNew.Enabled = True
         cmdSCBdEdit.Enabled = True
         cboBudgetYears.Enabled = False

        If flgChange = 1 Then
            cmdSCBdSave.Enabled = True
         Else
            cmdSCBdSave.Enabled = False
         End If
         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = True
   End Select
End Sub

Private Sub Form_Activate()
   bFormLoaded = True
End Sub

Private Sub Form_Load()
   bFormLoaded = False
   Me.Width = 11430
   Me.Height = 7545
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Refresh
   Me.BackColor = MODULEBACKCOLOR
   Frame1(1).BackColor = Me.BackColor

   initialiseGrid
   SCSumTotal

   cboBudgetYears.ListIndex = 0
'   Call WheelHook(Me.hwnd)
End Sub

Private Sub initialiseGrid()
   Call ConfigFlxSCBudgetDetails

   flgEdit = 0
   flgChange = 0
   flgNew = 0

   adoConn.Open getConnectionString

'  Loading Clients and Properties
   PrepareList adoConn, cboClient, cboProperty

   LoadFlxSCBudgetDetails
   LoadMatrix
   LoadFund
   LoadFY

   adoConn.Close

   ControlsModeRentBudgetDetails DefaultMode
End Sub
'
'Private Function TotalSCProperty(szPropertyID As String) As Double
'   Dim Rst2 As New ADODB.Recordset
'
'   szSQL = "SELECT GlobalSC.PropertyID, SUM(GlobalSC.TotalBudget) AS TOTALRENT " & _
'           "From   GlobalSC " & _
'           "WHERE  GlobalSC.PropertyID = '" & szPropertyID & "' " & _
'           "GROUP BY GlobalSC.PropertyID;"
''Debug.Print szSQL
'   Rst2.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
'      TotalSCProperty = CDbl(Rst2!TOTALRENT)
'   Else
'      TotalSCProperty = 0
'   End If
'
'   Rst2.Close
'   Set Rst2 = Nothing
'End Function

Private Sub LoadFund()
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund " & _
           "WHERE CategoryCode = 2;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      ReDim Data(2, adoRst.RecordCount) As String
    
      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
    
      cboSCFund.Clear
      cboSCFund.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If frmMMain.IsRibbonVersion Then
   Enabled = True
'   Else
'   frmGlobal1.Enabled = True
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set adoConn = Nothing

'   Call WheelUnHook(Me.hwnd)
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub txtBudget_Change()
   If Trim(txtBudget.text) <> "" And Val(Trim(txtBudget.text)) > 0 Then
    txtBudget.Locked = False
    computePpsf
   End If
End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
   If Not txtBudget.Locked Then DigitTextKeyPress txtBudget, KeyAscii
End Sub

Private Sub txtBudget_LostFocus()
   computePpsf
'   If flgEdit = 1 Or flgNew = 1 Then txtBudget.BackColor = &H80000005
End Sub

Private Sub computePpsf()
On Error GoTo ErrHandler
   If Trim(txtBudget.text) <> "" And Trim(txtTotalArea.text) <> "" Then
      Dim ppsf As Double
      
      ppsf = CDbl(txtBudget.text) / CDbl(Trim(txtTotalArea.text))
      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
   End If
   Exit Sub
ErrHandler:
   ShowMsgInTaskBar "Please ensure that the values for Total Budget and Total Area are valid before continuing", , "N"
End Sub

Private Sub txtTotalArea_LostFocus()
   computePpsf
End Sub

Private Sub updateGrid()
   If flgEdit = 0 Then
      flxSCBudgetDetails.AddItem ""
      If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) <> "" Then flxSCBudgetDetails.AddItem ""
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) = txtBudgetId.text
'      If frmMMain.IsRibbonVersion Then
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 1) = cboProperty.Value
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 2) = CInt(cboSCFund.Value)
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 3) = cboSCFund.text
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 5) = txtTotalArea.text
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 6) = txtPpsf.text
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 8) = nextRow
   Else
'      If frmMMain.IsRibbonVersion Then
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = cboProperty.Value
'      Else
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = frmGlobal1.cboProperty.value
'      End If
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2) = CInt(cboSCFund.Value)
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 3) = cboSCFund.text
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5) = txtTotalArea.text
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6) = txtPpsf.text
   End If
   flgEdit = 0
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
