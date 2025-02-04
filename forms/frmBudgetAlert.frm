VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBudgetAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Alerts"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBudgetAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12600
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Copy from lease"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Copy from lease"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CheckBox chkSMS 
      Caption         =   "by SMS"
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   1720
      Width           =   975
   End
   Begin VB.CheckBox chkAlarm 
      Caption         =   "by Alarm"
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   1720
      Width           =   975
   End
   Begin VB.CheckBox chkEmail 
      Caption         =   "by Email"
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   1720
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Copy from lease"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetAlert 
      Caption         =   "Set Alert"
      Height          =   375
      Left            =   11280
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7515
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtBudget 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8865
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBudgetAlert 
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   13
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
      _Band(0).Cols   =   13
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Assign To"
      Height          =   195
      Index           =   9
      Left            =   8640
      TabIndex        =   41
      Top             =   2160
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      Height          =   195
      Index           =   8
      Left            =   7800
      TabIndex        =   40
      Top             =   2160
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "SMS"
      Height          =   195
      Index           =   12
      Left            =   11640
      TabIndex        =   39
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm"
      Height          =   195
      Index           =   11
      Left            =   11040
      TabIndex        =   38
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Index           =   7
      Left            =   6960
      TabIndex        =   37
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Index           =   10
      Left            =   10440
      TabIndex        =   36
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Value Type"
      Height          =   195
      Index           =   6
      Left            =   5880
      TabIndex        =   35
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
      Height          =   195
      Index           =   5
      Left            =   4800
      TabIndex        =   34
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
      Height          =   195
      Index           =   4
      Left            =   3960
      TabIndex        =   33
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal Code"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   32
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   31
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Alert Reference"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Assign To"
      Height          =   195
      Index           =   7
      Left            =   10440
      TabIndex        =   29
      Top             =   1080
      Width           =   690
   End
   Begin MSForms.ComboBox cboAssign 
      Height          =   285
      Left            =   10440
      TabIndex        =   11
      Top             =   1320
      Width           =   2040
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3598;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   195
      Index           =   6
      Left            =   7515
      TabIndex        =   28
      Top             =   1080
      Width           =   405
   End
   Begin MSForms.ComboBox cboValueType 
      Height          =   285
      Left            =   5955
      TabIndex        =   8
      Top             =   1320
      Width           =   1560
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2752;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Value Type"
      Height          =   195
      Index           =   5
      Left            =   5955
      TabIndex        =   27
      Top             =   1080
      Width           =   780
   End
   Begin MSForms.ComboBox cboOperator 
      Height          =   285
      Left            =   4395
      TabIndex        =   7
      Top             =   1320
      Width           =   1560
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2752;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Operators"
      Height          =   195
      Index           =   4
      Left            =   4395
      TabIndex        =   26
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      Height          =   195
      Index           =   3
      Left            =   8865
      TabIndex        =   25
      Top             =   1080
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
      Height          =   195
      Index           =   2
      Left            =   3045
      TabIndex        =   24
      Top             =   1080
      Width           =   450
   End
   Begin MSForms.ComboBox cboNC 
      Height          =   285
      Left            =   1695
      TabIndex        =   5
      Top             =   1320
      Width           =   1320
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2328;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal Code"
      Height          =   195
      Index           =   1
      Left            =   1700
      TabIndex        =   23
      Top             =   1080
      Width           =   990
   End
   Begin MSForms.ComboBox cboType 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1560
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2752;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year:"
      Height          =   195
      Index           =   66
      Left            =   4605
      TabIndex        =   21
      Top             =   480
      Width           =   1005
   End
   Begin MSForms.ComboBox cmbFinancialYear 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   480
      Width           =   2880
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5080;503"
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
      Object.Width           =   "0;1940;0;0;0"
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "SC Fund"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   585
   End
   Begin MSForms.ComboBox cboFund 
      Height          =   285
      Left            =   870
      TabIndex        =   2
      Top             =   480
      Width           =   3135
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5530;503"
      TextColumn      =   2
      ColumnCount     =   6
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "705"
   End
   Begin MSForms.ComboBox cboProperty 
      Height          =   285
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7435;503"
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   6
      Left            =   4605
      TabIndex        =   18
      Top             =   120
      Width           =   645
   End
   Begin MSForms.ComboBox cboClient 
      Height          =   285
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   3120
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5503;503"
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "2116"
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   0
      Left            =   120
      Top             =   2160
      Width           =   12375
   End
End
Attribute VB_Name = "frmBudgetAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dtStartPnL    As Date
Public dtStartBS     As Date
Public dtEnd         As Date

Private Sub LoadProperty(adoConn As ADODB.Connection)
   Dim rRow       As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String

   On Error GoTo ErrorHandler
   
   If cboClient.text <> "" Then
      szSQL = "SELECT PropertyID, PropertyName, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboClient.Value & "' " & _
              "ORDER BY PropertyID;"
   Else
      szSQL = "SELECT PropertyID, PropertyName, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   End If
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   ReDim szaProperty(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           szaProperty(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cboProperty.Column() = szaProperty()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub cboClient_Click()
   Dim adoConn    As New ADODB.Connection
   Dim iRow       As Integer
   Dim K          As Integer

   adoConn.Open getConnectionString

   LoadCmbFinancialYear adoConn
   LoadProperty adoConn

   adoConn.Close
   Set adoConn = Nothing

'   K = 0
'   For iRow = 1 To flxBudget.Rows - 1
'      If flxBudget.TextMatrix(iRow, 5) = cboClient.Value Then
'         flxBudget.RowHeight(iRow) = 240
'         K = K + 1
'      End If
'   Next

   If cboProperty.ListCount > 0 Then cboProperty.ListIndex = 0
   If cmbFinancialYear.ListCount > 0 Then cmbFinancialYear.ListIndex = 0
'   If cmbPeriodFrom.ListCount > 0 Then cmbPeriodFrom.ListIndex = 0
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub LoadCmbFinancialYear(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer              'Open Flag index
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status " & _
           "FROM   FinancialYear " & _
           "WHERE  ClientID = '" & cboClient.Value & "' " & _
           "ORDER BY FY_StDate;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1
   ReDim Data(TotalCol, TotalRow) As String

   K = -1
   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
         If K = -1 And j = 4 Then
            If adoRst.Fields("Status").Value Then
               K = i
               dtStartPnL = CDate(adoRst.Fields("FY_StDate").Value)
               dtStartBS = CDate("01 January 2000")
            End If
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cmbFinancialYear.Column() = Data()
   cmbFinancialYear.ListIndex = K

   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Financial year has not been created for the client", "Y", "N"
   Exit Sub
End Sub

Private Sub LoadCboClient(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

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
      ReDim Data(2, adoRst.RecordCount - 1) As String

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

Private Sub cmdSetAlert_Click()
'MsgBox cboType.Value
End Sub

Private Sub Form_Load()
   Me.Width = 12690
   Me.Height = 7770
   Me.BackColor = MODULEBACKCOLOR
   chkEmail.BackColor = MODULEBACKCOLOR
   chkAlarm.BackColor = MODULEBACKCOLOR
   chkSMS.BackColor = MODULEBACKCOLOR

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadCboClient adoConn
   If cboClient.ListCount > 0 Then cboClient.ListIndex = 0
   LoadFund adoConn
   If cboFund.ListCount > 0 Then cboFund.ListIndex = 0

   adoConn.Close
   Set adoConn = Nothing

   ConfigFlxBudgetAlert
   LoadOptions
   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadOptions()
   With cboType
      .AddItem ""
      .Column(0, 0) = "T"
      .Column(1, 0) = "Total"
      .AddItem ""
      .Column(0, 1) = "L"
      .Column(1, 1) = "Line"
      .ListIndex = -1
   End With
   
   With cboNC
      .AddItem ""
      .Column(0, 0) = "NA"
      .Column(1, 0) = "N/A"
      .AddItem ""
      .Column(0, 1) = "5000"
      .Column(1, 1) = "Code"
      .ListIndex = -1
   End With
   
   With cboOperator
      .AddItem ""
      .Column(0, 0) = "E"
      .Column(1, 0) = "Equal to"
      .AddItem ""
      .Column(0, 1) = "L"
      .Column(1, 1) = "Less than"
      .AddItem ""
      .Column(0, 2) = "G"
      .Column(1, 2) = "Greater than"
      .ListIndex = -1
   End With

   With cboValueType
      .AddItem ""
      .Column(0, 0) = "A"
      .Column(1, 0) = "Amount"
      .AddItem ""
      .Column(0, 1) = "P"
      .Column(1, 1) = "Percentage"
      .ListIndex = -1
   End With
End Sub

Public Sub ConfigFlxBudgetAlert()
   Dim i As Integer, szHeader As String

   With flxBudgetAlert
      szHeader$ = "|<AlertRef|<Type|<NC|>Actual|>Operator|<ValueType" & _
                  "|>Value|>Budget|<Assign|<Email|<Alarm|<SMS|ClientID|Property|Fund|FY"
      .FormatString = szHeader
      .Clear
      .Rows = 2
      .Cols = 17

      .RowHeight(0) = 0
      .ColWidth(0) = 0
      For i = 1 To .Cols - 6
         .ColWidth(i) = Label2(i + 1).Left - Label2(i).Left
      Next i

      .ColWidth(i) = .Width + .Left - Label2(i).Left - 300
      .ColWidth(i + 1) = 0
      .ColWidth(i + 2) = 0
      .ColWidth(i + 3) = 0
      .ColWidth(i + 4) = 0
      
      .TextMatrix(1, 1) = "BL01-131002-1"
      .TextMatrix(1, 2) = "Budget Line"
      .TextMatrix(1, 3) = "5000"
      .TextMatrix(1, 4) = "828.00"
      .TextMatrix(1, 5) = "Equal to"
      .TextMatrix(1, 6) = "Percent"
      .TextMatrix(1, 7) = "50.00"
      .TextMatrix(1, 8) = "1800.00"
      .TextMatrix(1, 9) = "Peter Jones"
      .TextMatrix(1, 10) = "YES"
      .TextMatrix(1, 11) = "NO"
      .TextMatrix(1, 12) = "NO"
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
   Call WheelUnHook(Me.hWnd)
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
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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
