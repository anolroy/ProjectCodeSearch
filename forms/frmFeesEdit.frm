VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFeesEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit - Fees"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFeesEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCostCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtProjRef 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Canc&el"
      Height          =   405
      Index           =   2
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   405
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "&Save"
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtVat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtDetails 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin MSForms.ComboBox cboTaxCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   26
      Top             =   3360
      Width           =   1695
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2990;503"
      TextColumn      =   1
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1058;1058"
   End
   Begin MSForms.ComboBox cboDemandCategory 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   480
      Width           =   3495
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6165;503"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1234"
   End
   Begin MSForms.ComboBox cboFund 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Top             =   1560
      Width           =   3495
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6165;503"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1234"
   End
   Begin MSForms.ComboBox cboNC 
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   1200
      Width           =   3495
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6165;503"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1234"
   End
   Begin MSForms.ComboBox cboSupplier 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   3495
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6165;503"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1234"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Category"
      Height          =   300
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal Code"
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Project Reference"
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Code"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      Height          =   225
      Index           =   9
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   660
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      Height          =   225
      Index           =   10
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   225
      Index           =   12
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   225
      Index           =   13
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   420
   End
End
Attribute VB_Name = "frmFeesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private bChangesMade As Boolean
'
'Private Sub cboDemandCategory_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub cboFund_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub cboNC_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub cboSupplier_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub cboTaxCode_Click()
'   txtVat.text = txtNet.text * (cboTaxCode.Column(1) / 100)
'   txtTotal.text = Val(txtVat.text) + Val(txtNet.text)
'   bChangesMade = True
'End Sub
'
'Private Sub cmdClose_Click(Index As Integer)
'   If Index = 1 And bChangesMade Then
'      If MsgBox("Do you wish to save all changes before close the window?", vbQuestion + vbYesNo, "Fees Edit") = vbYes Then
'         cmdName_Click
'      End If
'   End If
'   If Index = 2 And bChangesMade Then
'      If MsgBox("Do you wish to discard all changes?", vbQuestion + vbYesNo, "Fees Edit") = vbNo Then Exit Sub
'   End If
'
'   Unload Me
'End Sub
'
'Private Sub cmdName_Click()
'   If bChangesMade Then
'      If MsgBox("Do you wish to save?", vbQuestion + vbYesNo, "Fees Edit") = vbNo Then Exit Sub
'
'      With frmFees
'         .flxMngFees.TextMatrix(.flxMngFees.row, 0) = txtDate.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 1) = cboDemandCategory.Value
'         .flxMngFees.TextMatrix(.flxMngFees.row, 2) = cboSupplier.Value
'         .flxMngFees.TextMatrix(.flxMngFees.row, 4) = cboNC.Value
'         .flxMngFees.TextMatrix(.flxMngFees.row, 5) = cboFund.Value
'         .flxMngFees.TextMatrix(.flxMngFees.row, 6) = txtProjRef.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 7) = txtCostCode.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 8) = txtDetails.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 9) = txtNet.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 10) = cboTaxCode.Value
'         .flxMngFees.TextMatrix(.flxMngFees.row, 11) = txtVat.text
'         .flxMngFees.TextMatrix(.flxMngFees.row, 12) = txtTotal.text
'      End With
'   End If
'   Unload Me
'End Sub
'
'Private Sub Form_Load()
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   Me.Top = 0
'   Me.Left = 0
'
'   Me.Width = 5265
'   Me.Height = 5670
'   Me.BackColor = MODULEBACKCOLOR
'
'   LoadDemandCategory adoConn
'   LoadAllSupplier adoConn
'   LoadNC adoConn
'   LoadFund adoConn
'   LoadTaxCode adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   bChangesMade = False
'End Sub
'
'Private Sub LoadTaxCode(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   szSQL = "SELECT VAT_CODE, VAT_RATE FROM TLBVATCODE"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To adoRst.RecordCount - 1
'       For j = 0 To adoRst.Fields.count - 1
'           Data(j, i) = adoRst.Fields(j)
'       Next j
'       adoRst.MoveNext
'   Next i
'   cboTaxCode.Clear
'   cboTaxCode.Column() = Data()
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
'
'Private Sub LoadDemandCategory(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   szSQL = "SELECT Code, Value " & _
'           "FROM SecondaryCode " & _
'           "WHERE PrimaryCode = 'DCTG';"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.RecordCount < 1 Then
'      adoRst.Close
'      Set adoRst = Nothing
'      Exit Sub
'   End If
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To adoRst.RecordCount - 1
'       For j = 0 To adoRst.Fields.count - 1
'           Data(j, i) = adoRst.Fields(j)
'       Next j
'       adoRst.MoveNext
'   Next i
'   cboDemandCategory.Clear
'   cboDemandCategory.Column() = Data()
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
'
'Private Sub LoadAllSupplier(ByVal adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, iTotalRow As Integer, j As Integer
'   Dim i As Integer, iTotalCol As Integer, Data() As String
'
'   On Error GoTo ErrorHandler
'
''   szSQL = "SELECT SupplierID, SupplierName  " & _
''           "FROM Supplier " & _
''           "ORDER BY SupplierName;"
'   szSQL = "SELECT AgentID, AgentName  " & _
'           "FROM Agent " & _
'           "ORDER BY AgentName;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
'
'   ReDim Data(iTotalCol, iTotalRow - 1) As String
'
'   For i = 0 To iTotalRow
'       For j = 0 To iTotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboSupplier.Clear
'   cboSupplier.Column() = Data()
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox Err.description & "::" & Err.Number
'
'   Set adoRst = Nothing
'End Sub
'
'Private Sub LoadNC(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, TotalRow As Integer
'   Dim Data() As String, i As Integer
'
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   TotalRow = adoRst.RecordCount
'   ReDim Data(2, TotalRow) As String
'   cboNC.Clear
'   i = 0
'
'   While Not adoRst.EOF
'      Data(0, i) = adoRst.Fields.Item("Code").Value
'      Data(1, i) = adoRst.Fields.Item("Name").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cboNC.Column() = Data()
'   ' Destroy Objects
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub
'
'Private Sub LoadFund(adoConn As ADODB.Connection)
'   Dim rRow As Integer, iRec As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT FundID, FundName FROM Fund;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Fund has not been setup.", vbExclamation, "Load Fund in Global"
'   Else
'      ReDim Data(2, adoRst.RecordCount) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
'         Data(1, rRow) = adoRst.Fields.Item("FundName").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'      cboFund.Clear
'      cboFund.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub
'
'Private Sub txtCostCode_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub txtDate_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub txtDetails_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub txtNet_Change()
'   bChangesMade = True
'End Sub
'
'Private Sub txtProjRef_Change()
'   bChangesMade = True
'End Sub
