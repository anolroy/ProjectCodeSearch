VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRRR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income Received Report"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRRR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8160
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   6000
      Begin VB.Frame fraCategory 
         Caption         =   "Fund Categories:"
         Height          =   2070
         Left            =   90
         TabIndex        =   28
         Top             =   2655
         Width           =   5835
         Begin VB.CheckBox chkallcategories 
            Caption         =   "All Categories"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2025
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCategories 
            Height          =   1365
            Left            =   90
            TabIndex        =   30
            Top             =   630
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   2408
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
            BackColorFixed  =   13553358
            ForeColorFixed  =   -2147483634
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483630
            BackColorBkg    =   16777215
            GridColor       =   14737632
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
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
      End
      Begin VB.Frame fraProperties 
         Caption         =   "Properties:"
         Height          =   2115
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Width           =   5835
         Begin VB.CheckBox chkAllProperties 
            Caption         =   "All Properties"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2025
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
            Height          =   1455
            Left            =   90
            TabIndex        =   27
            Top             =   585
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   2566
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
            BackColorFixed  =   13553358
            ForeColorFixed  =   -2147483634
            BackColorSel    =   12648447
            ForeColorSel    =   -2147483630
            BackColorBkg    =   16777215
            GridColor       =   14737632
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
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
      End
      Begin VB.CommandButton cmdLandLord 
         Caption         =   ". ."
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
         Left            =   5385
         TabIndex        =   10
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   4095
         TabIndex        =   20
         Top             =   7560
         Width           =   1575
      End
      Begin VB.CommandButton cmdGenReport 
         Caption         =   "&Generate Report"
         Default         =   -1  'True
         Height          =   375
         Left            =   135
         TabIndex        =   19
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transaction Date - "
         Height          =   735
         Left            =   135
         TabIndex        =   14
         Top             =   6735
         Width           =   5595
         Begin VB.TextBox txtToDate 
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
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   16
            Top             =   285
            Width           =   1575
         End
         Begin VB.TextBox txtFromDate 
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
            Left            =   720
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "01/01/2000"
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   18
            Top             =   345
            Width           =   210
         End
         Begin VB.Label lblSpecifyDateRange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   390
         End
      End
      Begin VB.TextBox txtLlID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtLlName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2055
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   3375
      End
      Begin VB.CheckBox chkSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   4815
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFund 
         Height          =   1545
         Left            =   135
         TabIndex        =   24
         Top             =   5130
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2725
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   180
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "All "
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   22
         Top             =   4860
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Fund"
         Height          =   195
         Index           =   3
         Left            =   1245
         TabIndex        =   21
         Top             =   4860
         Width           =   4425
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   6345
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   4005
      Visible         =   0   'False
      Width           =   5295
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
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   2
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5900
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   6
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1875
         TabIndex        =   5
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   4
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   3
         Top             =   375
         Width           =   3555
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6271;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   15
         Left            =   0
         Top             =   80
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frmRRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name : Income received report
Option Explicit
Dim sText  As String
Dim szPropertyList As String
Private Sub cboCategory_Click()
    chkSelectAll.Value = 0
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    ConfigFlxFund
    LoadFlxFund adoconn
    adoconn.Close
    chkSelectAll.Value = 1
End Sub

Private Sub chkallcategories_Click()
      Dim iRow As Integer
   If chkallcategories.Value = 1 Then
        For iRow = 1 To flxCategories.Rows - 1
           flxCategories.TextMatrix(iRow, 0) = "X"
        Next iRow
   Else
        For iRow = 1 To flxCategories.Rows - 1
           flxCategories.TextMatrix(iRow, 0) = ""
        Next iRow
   End If
End Sub

Private Sub chkSelectAll_Click()
   Dim i As Integer

   If chkSelectAll.Value Then
      For i = 1 To flxFund.Rows - 1
         If flxFund.RowHeight(i) > 0 And flxFund.TextMatrix(i, 0) = "" Then
            'SelectFlxGridRow 0, flxFund, i
            flxFund.TextMatrix(i, 0) = "X"
         End If
      Next i
   Else
      For i = 1 To flxFund.Rows - 1
         If flxFund.RowHeight(i) > 0 And flxFund.TextMatrix(i, 0) = "X" Then
            'SelectFlxGridRow 0, flxFund, i
            flxFund.TextMatrix(i, 0) = ""
         End If
      Next i
   End If
End Sub
Private Sub ConfigFlxProperties()
   Dim szHeader As String
   flxProperties.Clear
   flxProperties.Rows = 2
   szHeader$ = "<|<|<|<"
   With flxProperties
      .FormatString = szHeader
      .Cols = 4
      .RowHeight(0) = 0
      .ColWidth(0) = 200 'Label2(0).Left - .Left '200                 '"X"
      .ColWidth(1) = 2000 'Label2(1).Left - Label2(0).Left 'Property ID
      .ColWidth(2) = 3500 'Label2(2).Left - Label2(1).Left 'Property Name
      .ColWidth(3) = 0 '.Width + .Left - Label2(2).Left - 300 'Client ID
   End With
End Sub
Private Sub loadflxProperties()
    Dim szSQL   As String
    Dim r       As Integer
    Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    adoconn.Open getConnectionString
     szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY where ClientID='" & txtLlName.text & "'" & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   ConfigFlxProperties
   r = 1

   While Not adoRst.EOF
      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      flxProperties.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value
      flxProperties.RowHeight(r) = 240
      r = r + 1
      
      adoRst.MoveNext
      If Not adoRst.EOF Then flxProperties.AddItem ""
   Wend
    Debug.Print r
   adoRst.Close
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
   flxProperties.row = 0
End Sub

Private Function LoadflxCategory(adoconn As ADODB.Connection)
'   Dim szSQL As String
'   szSQL = "SELECT   SC.Code AS CatID,SC.Value AS CategoryCode FROM SecondaryCode AS SC WHERE  SC.PrimaryCode = 'DCTG';"
'   Dim adoRst As New ADODB.Recordset
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   Dim iTotalRow, iTotalCol, i, j As Integer
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
'   ReDim Data(iTotalCol - 1, iTotalRow) As String
'
'   ReDim szNominal(iTotalCol - 1, iTotalRow - 1) As String
'   Data(0, 0) = "0"
'   Data(1, 0) = "ALL CATEGORIES"
'   For i = 1 To iTotalRow
'      For j = 0 To iTotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboCategory.Clear
'   cboCategory.Column() = Data()
   
   '***********************
   'Dim adoconn As New Adodb.Connection
   Dim rstClient As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer
   'adoconn.Open getConnectionString
   'On Error GoTo ErrorHandler
    szSQL = "SELECT   SC.Code AS CatID,SC.Value AS CategoryCode FROM SecondaryCode AS SC WHERE  SC.PrimaryCode = 'DCTG';"

   rstClient.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
  
   
   flxCategories.Rows = rstClient.RecordCount + 2
   iRow = 1
        flxCategories.TextMatrix(iRow, 0) = "X"
        flxCategories.TextMatrix(iRow, 1) = "ALL"
        flxCategories.TextMatrix(iRow, 2) = "ALL Category"
   iRow = 2
   flxCategories.RowHeight(0) = 0
   While Not rstClient.EOF
      flxCategories.TextMatrix(iRow, 1) = rstClient!CatID
      flxCategories.TextMatrix(iRow, 2) = rstClient!CategoryCode
      flxCategories.TextMatrix(iRow, 3) = ""
      rstClient.MoveNext
      iRow = iRow + 1
   Wend
   rstClient.Close
'   adoconn.Close
'   Set adoconn = Nothing
NoRes:
   Set rstClient = Nothing
   Exit Function

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set rstClient = Nothing
   
   
   
End Function
Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Function CreateListOfProp() As Integer
   Dim i As Integer

   szPropertyList = ""
   
   For i = 0 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" And flxProperties.RowHeight(i) > 0 Then
         szPropertyList = "'" & flxProperties.TextMatrix(i, 1) & "'" & ", " & szPropertyList
      End If
   Next i
   If Len(szPropertyList) > 2 Then
      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
      CreateListOfProp = Len(szPropertyList)
      Exit Function
   End If
   CreateListOfProp = 0
End Function

Private Sub MarkPropOfSelection(adoconn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = '';"
   adoconn.Execute szSQL

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = 'X' " & _
           "WHERE  PropertyID IN (" & szPropertyList & ");"
'Debug.Print szSQL
   adoconn.Execute szSQL
End Sub
Private Sub cmdGenReport_Click()
   Dim strCategory  As String
   Dim adoconn As New ADODB.Connection
   If txtLlID.text = "" Then
      MsgBox "Please a landlord.", vbInformation + vbOKOnly, "Landlord"
      cmdLandLord.SetFocus
      Exit Sub
   End If
   If Not IsDTSel Then
      MsgBox "Please select fund.", vbInformation + vbOKOnly, "Fund"
      Exit Sub
   End If
   If IsCatSel = "" Then
      MsgBox "Please select Category.", vbInformation + vbOKOnly, "Category"
      Exit Sub
   Else
      strCategory = IsCatSel
   End If
   
   If CreateListOfProp = 0 Then
      MsgBox "Please select Property from the grid."
      flxProperties.SetFocus
      Exit Sub
   End If
   
   adoconn.Open getConnectionString
   MarkPropOfSelection adoconn
   adoconn.Close
   Set adoconn = Nothing

   If IsDate(txtFromDate.text) = False Then
      MsgBox "Please enter a valid from date.", vbInformation + vbOKOnly, "From Date"
      txtFromDate.SetFocus
      Exit Sub
   End If
   If IsDate(txtToDate.text) = False Then
      MsgBox "Please enter a valid to date.", vbInformation + vbOKOnly, "To Date"
      txtToDate.SetFocus
      Exit Sub
   End If
   If strCategory = "ALL" Then
        strCategory = "ALL Fund Category "
   End If
   cmdGenReport.Enabled = False
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RRR.rpt")

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CStr(txtLlID.text)
   Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
   Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
   Report.ParameterFields(4).AddCurrentValue strCategory 'going as string to show in the header caption anol 2019-06-22
   Load frmReport
   frmReport.LoadReportViewer Report
   cmdGenReport.Enabled = True
End Sub
'Private Function CreateListOfProp() As Integer
'   Dim i As Integer
'
'   szPropertyList = ""
'
'   For i = 0 To flxProperties.Rows - 1
'      If flxProperties.TextMatrix(i, 0) = "X" And flxProperties.RowHeight(i) > 0 Then
'         szPropertyList = "'" & flxProperties.TextMatrix(i, 1) & "'" & ", " & szPropertyList
'      End If
'   Next i
'   If Len(szPropertyList) > 2 Then
'      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
'      CreateListOfProp = Len(szPropertyList)
'      Exit Function
'   End If
'   CreateListOfProp = 0
'End Function
'
'Private Sub MarkPropOfSelection(adoconn As ADODB.Connection)
'   Dim szSQL As String
'
'   szSQL = "UPDATE PROPERTY " & _
'           "SET    RAS = '';"
'   adoconn.Execute szSQL
'
'   szSQL = "UPDATE PROPERTY " & _
'           "SET    RAS = 'X' " & _
'           "WHERE  PropertyID IN (" & szPropertyList & ");"
''Debug.Print szSQL
'   adoconn.Execute szSQL
'End Sub

Private Function IsDTSel() As Boolean
   Dim i As Integer
   Dim szFunds As String

   szFunds = ""

   For i = 1 To flxFund.Rows - 1
      If flxFund.TextMatrix(i, 0) = "X" And flxFund.RowHeight(i) > 0 Then
         szFunds = flxFund.TextMatrix(i, 1) & ", " & szFunds
      End If
   Next i

   If szFunds = "" Then
      IsDTSel = False
   Else
      IsDTSel = True
      szFunds = Left(szFunds, Len(szFunds) - 2)

      Dim conLandlord As New ADODB.Connection

      conLandlord.Open getConnectionString

      conLandlord.Execute "UPDATE Fund SET RRR='';"
      conLandlord.Execute "UPDATE Fund SET RRR='RRR' WHERE FundID IN (" & szFunds & ");"

      conLandlord.Close
      Set conLandlord = Nothing
   End If
End Function
Private Function IsCatSel() As String
   Dim i As Integer
   For i = 1 To flxCategories.Rows - 1
      If flxCategories.TextMatrix(i, 0) = "X" Then
         IsCatSel = flxCategories.TextMatrix(i, 1)
         Exit Function
      End If
   Next i
End Function
Private Function IsPripertySel() As String
   Dim i As Integer
   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" Then
         IsPripertySel = flxProperties.TextMatrix(i, 1)
         Exit Function
      End If
   Next i
End Function
Private Sub cmdPicCLose_Click()
   picClient.Visible = False
   Frame1.Enabled = True
End Sub

Private Sub cmdLandlord_Click()
    sText = "1"
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    ConfigflxClient
    Call LoadflxClient(adoconn, "")
    adoconn.Close
    Set adoconn = Nothing
    picClient.Top = 135
    picClient.Left = 180
    picClient.Visible = True
    picClient.ZOrder 0
   
    Frame1.Enabled = False
End Sub
Private Sub chkAllProperties_Click()
'   If bCallingFromGrid Then
'      bCallingFromGrid = False
'      Exit Sub
'   End If

   Dim iRow As Integer
   If chkAllProperties.Value = 1 Then
        For iRow = 1 To flxProperties.Rows - 1
           flxProperties.TextMatrix(iRow, 0) = "X"
        Next iRow
   Else
        For iRow = 1 To flxProperties.Rows - 1
           flxProperties.TextMatrix(iRow, 0) = ""
        Next iRow
   End If
End Sub

Private Sub flxCategories_Click()
    Dim i As Integer
    Call SelectOnly1RowFlxGrid(flxCategories, flxCategories.row, 0)
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    ConfigFlxFund
    LoadFlxFund adoconn
    adoconn.Close
End Sub

Private Sub flxProperties_Click()
    If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
        flxProperties.TextMatrix(flxProperties.row, 0) = ""
     Else
        flxProperties.TextMatrix(flxProperties.row, 0) = "X"
     End If
End Sub
'Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
'   Dim iRow       As Integer
'   Dim iCol       As Integer
'   Dim iColPaint  As Integer
'
'   iColPaint = IIf(iColID = 0, 1, 0)
'
'   For iRow = 1 To conFlxGrid.Rows - 1
'      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
'         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then Exit Sub
'         conFlxGrid.TextMatrix(iRow, iColID) = ""
'         conFlxGrid.row = iRow
'         For iCol = iColPaint To conFlxGrid.Cols - 1
'            conFlxGrid.col = iCol
'            conFlxGrid.CellBackColor = vbWhite
'         Next iCol
'      End If
'   Next iRow
'
'   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
'   conFlxGrid.row = iNewRow
'
'   For iCol = iColPaint To conFlxGrid.Cols - 1
'      conFlxGrid.col = iCol
'      conFlxGrid.CellBackColor = RGB(174, 179, 233)
'   Next iCol
'End Sub
Private Sub flxFund_Click()
    If flxFund.TextMatrix(flxFund.row, 0) = "X" Then
        flxFund.TextMatrix(flxFund.row, 0) = ""
     Else
        flxFund.TextMatrix(flxFund.row, 0) = "X"
     End If
End Sub

Private Sub flxFund_RowColChange()
   'SelectFlxGridRow 0, flxFund, flxFund.row
'     If flxFund.TextMatrix(flxFund.row, 0) = "X" Then
'        flxFund.TextMatrix(flxFund.row, 0) = ""
'     Else
'        flxFund.TextMatrix(flxFund.row, 0) = "X"
'     End If
End Sub

Private Sub flxClient_Click()
   If flxClient.TextMatrix(flxClient.row, 1) = "" Then Exit Sub
   If sText = "1" Then
        txtLlID.text = flxClient.TextMatrix(flxClient.row, 2)
        txtLlName.text = flxClient.TextMatrix(flxClient.row, 1)
        Call loadflxProperties
        chkAllProperties.Value = 1
        chkAllProperties_Click
  
   End If
'   Dim i     As Integer
'   Dim bFlag As Boolean

'   bFlag = False
   chkSelectAll.Value = False

'   For i = 1 To flxFund.Rows - 1
'      If flxFund.TextMatrix(i, 3) = txtLlID.text Then
'         flxFund.RowHeight(i) = 240
'         bFlag = True
'      Else
'         flxFund.RowHeight(i) = 0
'      End If
'   Next i

   picClient.Visible = False
   Frame1.Enabled = True
'   If Not bFlag Then ShowMsgInTaskBar "Landlord has not been assigned against any property.", "Y", "N"
End Sub

Private Sub Form_Load()
    Dim adoconn As New ADODB.Connection
    Dim MODULEBACKCOLOR1 As String
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
    Me.Height = 8685
    Me.Width = 6075
    MODULEBACKCOLOR1 = 14737632
    Me.BackColor = MODULEBACKCOLOR1
    Frame2.BackColor = MODULEBACKCOLOR1
    Frame1.BackColor = MODULEBACKCOLOR1
    Label1(3).BackColor = MODULEBACKCOLOR1
    Label1(2).BackColor = MODULEBACKCOLOR1
    chkSelectAll.BackColor = MODULEBACKCOLOR1
    fraProperties.BackColor = MODULEBACKCOLOR1
    fraCategory.BackColor = MODULEBACKCOLOR1
    chkAllProperties.BackColor = MODULEBACKCOLOR1
    chkallcategories.BackColor = MODULEBACKCOLOR1
    txtToDate.text = Format(Now, "dd/mm/yyyy")
    
    adoconn.Open getConnectionString
    
'    ConfigflxClient
'    LoadflxClient adoConn
    'added by anol 27 Sep 2015
    'loadcombocategory adoconn
    loadfirstclient adoconn
    Call loadflxProperties
    chkAllProperties.Value = 1
    ConfigflxCategory
    LoadflxCategory adoconn
'    loadfirstcategory adoconn
    'End of Addition
    ConfigFlxFund
    LoadFlxFund adoconn
    
    adoconn.Close
    Set adoconn = Nothing
    
    Call WheelHook(Me.hWnd)
End Sub
'Private Function loadfirstcategory(adoconn As ADODB.Connection)
'   Dim rstClient As New ADODB.Recordset
'   Dim szSQL As String
'   Dim iRow As Integer
'
'   'On Error GoTo ErrorHandler
'    szSQL = "SELECT   SC.Code AS CatID,SC.Value AS CategoryCode FROM SecondaryCode AS SC WHERE  SC.PrimaryCode = 'DCTG';"
'
'   rstClient.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   flxClient.Rows = rstClient.RecordCount + 1
'   iRow = 1
'   flxClient.RowHeight(0) = 0
'   If Not rstClient.EOF Then
'      txtCategory.Tag = rstClient!CatID
'      txtCategory.text = rstClient!CategoryCode
'   End If
'   rstClient.Close
'
'NoRes:
'   Set rstClient = Nothing
'   Exit Function
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   Set rstClient = Nothing
'
'End Function
Private Function loadfirstclient(adoconn As ADODB.Connection)
    Dim rstClient As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer
   
   On Error GoTo ErrorHandler
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM   CLIENT " & _
           "ORDER BY CLIENTNAME;"

   rstClient.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   
   
   flxClient.Rows = rstClient.RecordCount + 1
   iRow = 1
   flxClient.RowHeight(0) = 0
   If Not rstClient.EOF Then
      txtLlName.text = rstClient!clientID
      txtLlID.text = rstClient!ClientName
     
   End If
   rstClient.Close

NoRes:
   Set rstClient = Nothing
   Exit Function

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set rstClient = Nothing
   
End Function
'Private Sub loadcombocategory(adoconn As ADODB.Connection)
'    'This loads fund category
'   Dim szSQL As String
'   szSQL = "SELECT   SC.Code AS CatID,SC.Value AS CategoryCode FROM SecondaryCode AS SC WHERE  SC.PrimaryCode = 'DCTG';"
'   Dim adoRst As New ADODB.Recordset
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'   Dim iTotalRow, iTotalCol, i, j As Integer
'   If adoRst.EOF Then GoTo NoRes
'
'   iTotalRow = adoRst.RecordCount
'   iTotalCol = adoRst.Fields.count
'   ReDim Data(iTotalCol - 1, iTotalRow) As String
'
'   ReDim szNominal(iTotalCol - 1, iTotalRow - 1) As String
'   Data(0, 0) = "0"
'   Data(1, 0) = "ALL CATEGORIES"
'   For i = 1 To iTotalRow
'      For j = 0 To iTotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboCategory.Clear
'   cboCategory.Column() = Data()
'NoRes:
'End Sub
Private Sub ConfigFlxFund()
   flxFund.Cols = 3
   flxFund.Clear
   flxFund.ColWidth(0) = Label1(2).Left - flxFund.Left            'Solid column
   flxFund.ColWidth(1) = Label1(3).Left - Label1(2).Left                'FundID
   flxFund.ColAlignment(1) = vbLeftJustify
   flxFund.ColWidth(2) = 4000 'Label1(4).Left - Label1(3).Left                'FundName
   flxFund.Rows = 2
   flxFund.RowHeight(0) = 0
End Sub

Private Sub ConfigflxClient()
   Dim szHeader As String

   flxClient.Cols = 4
   flxClient.Clear
   szHeader$ = "|<ID|<Name|<Type"
   flxClient.FormatString = szHeader$
   flxClient.ColWidth(0) = 80        'Solid column
   flxClient.ColWidth(1) = 1400        'Landlord ID
   flxClient.ColWidth(2) = 3500       'Landlord Name
   flxClient.ColWidth(3) = 0        'Post Code
   flxClient.Rows = 2
   flxClient.RowHeight(0) = 0
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
'   flxClient.RowHeightMin = 255
End Sub
Private Sub ConfigflxCategory()
   Dim szHeader As String

   flxCategories.Cols = 5
   flxCategories.Clear
   szHeader$ = "|<ID|<Name|<Type"
   flxCategories.FormatString = szHeader$
   flxCategories.ColWidth(0) = 270        'Solid column
   flxCategories.ColWidth(1) = 1400        'Landlord ID
   flxCategories.ColAlignment(1) = vbRightJustify
   flxCategories.ColWidth(2) = 3500       'Landlord Name
   flxCategories.ColWidth(3) = 0        'Post Code
   flxCategories.Rows = 2
   flxCategories.RowHeight(0) = 0
'   lblClientID.Caption = "Category ID"
'   lblClientName.Caption = "Category Name"
'   flxCategories.RowHeightMin = 255
End Sub
Private Sub LoadFlxFund(adoconn As ADODB.Connection)
   Dim rstF As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer
'Modified by anol 07 Oct 2015
'   szSQL = "SELECT F.FundID, F.FundName " & _
'           "FROM Fund AS F  " & _
'           "WHERE CategoryCode = 1 " & _
'           "ORDER BY F.FundID;"
        
        If IsCatSel = "" Then
             Exit Sub
        End If
        If IsCatSel = "ALL" Then
                 szSQL = "SELECT F.FundID, F.FundCode, F.FundName, SC.Value AS CategoryCode, SC.Code AS CatID " & _
                "FROM FUND AS F, SecondaryCode AS SC " & _
                "WHERE F.CategoryCode = CINT(SC.Code) AND " & _
                    "SC.PrimaryCode = 'DCTG' " & _
                "ORDER BY F.FundID;"
        Else
                szSQL = "SELECT F.FundID, F.FundCode, F.FundName, SC.Value AS CategoryCode, SC.Code AS CatID " & _
                   "FROM FUND AS F, SecondaryCode AS SC " & _
                   "WHERE F.CategoryCode = CINT(SC.Code) AND F.CategoryCode =" & IsCatSel & " AND " & _
                       "SC.PrimaryCode = 'DCTG' " & _
                   "ORDER BY F.FundID;"
        End If
   rstF.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   flxFund.Rows = rstF.RecordCount + 1
   iRow = 1

   While Not rstF.EOF
      flxFund.TextMatrix(iRow, 1) = rstF!fundID
      flxFund.TextMatrix(iRow, 2) = rstF!FundName
      rstF.MoveNext
      iRow = iRow + 1
   Wend

   rstF.Close
   Set rstF = Nothing
   flxFund.row = 0
End Sub
Private Sub txtSearchClientName_Change()
   Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Dim tempstr As String
    If Trim(txtSearchClientName.text) = "" Then
         Call LoadflxClient(adoconn, "")
         Exit Sub
    Else
         txtSearchClientID.text = ""
    End If
    tempstr = txtSearchClientName.text
    tempstr = Replace(tempstr, "'", "''")
    Call LoadflxClient(adoconn, "ClientName Like '%" & tempstr & "%'")
    adoconn.Close
    Set adoconn = Nothing
End Sub
Private Sub txtSearchClientID_Change()
    Dim tempstr As String
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    If Trim(txtSearchClientID.text) = "" Then
         Call LoadflxClient(adoconn, "")
         Exit Sub
    End If
    tempstr = txtSearchClientID.text
    tempstr = Replace(tempstr, "'", "''")
    Call LoadflxClient(adoconn, "ClientID Like '%" & tempstr & "%'")
    adoconn.Close
    Set adoconn = Nothing
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyDown Then
            flxClient.SetFocus
     End If
End Sub
Private Sub LoadflxClient(adoconn As ADODB.Connection, Filter As String)
   Dim rstClient As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow As Integer
   
   On Error GoTo ErrorHandler
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM   CLIENT " & _
           "ORDER BY CLIENTNAME;"

   rstClient.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Filter <> "" Then
            rstClient.Filter = Filter
   End If
   
   flxClient.Rows = rstClient.RecordCount + 1
   iRow = 1
   flxClient.RowHeight(0) = 0
   While Not rstClient.EOF
      flxClient.TextMatrix(iRow, 1) = rstClient!clientID
      flxClient.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClient.TextMatrix(iRow, 3) = "Client"
      rstClient.MoveNext
      iRow = iRow + 1
   Wend
   rstClient.Close

NoRes:
   Set rstClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set rstClient = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
   SelTxtInCtrl txtFromDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtToDate_GotFocus()
   SelTxtInCtrl txtToDate
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate
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
