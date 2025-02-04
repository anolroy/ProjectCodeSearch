VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSecondaryCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codes"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   Icon            =   "frmSecondaryCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleMode       =   0  'User
   ScaleWidth      =   9974.854
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCodes 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   480
      TabIndex        =   11
      Top             =   480
      Width           =   7695
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   2
         Top             =   135
         Width           =   2775
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   135
         Width           =   2595
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Clos&e"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5205
         TabIndex        =   8
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2670
         TabIndex        =   5
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1395
         TabIndex        =   6
         Top             =   5760
         Width           =   1035
      End
      Begin VB.TextBox txtDescription 
         Height          =   855
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   600
         Width           =   6435
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3930
         TabIndex        =   7
         Top             =   5760
         Width           =   1035
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMain 
         Bindings        =   "frmSecondaryCode.frx":0442
         Height          =   3435
         Left            =   120
         TabIndex        =   12
         Top             =   2100
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   6059
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   12632256
         BackColorSel    =   4210752
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   375
         Index           =   3
         Left            =   4800
         TabIndex        =   20
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   19
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Name"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Short Name:"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   165
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackColor       =   &H009F0258&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F9C6CD&
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   7395
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1860
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   6360
      Top             =   7920
      Visible         =   0   'False
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoPrimaryCode1 
      Height          =   375
      Left            =   2880
      Top             =   7200
      Visible         =   0   'False
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
      Caption         =   "PrimaryCode"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cboPrimaryCode_ 
      Bindings        =   "frmSecondaryCode.frx":0460
      DataSource      =   "adoPrimaryCode1"
      Height          =   315
      Left            =   1440
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Value"
      BoundColumn     =   "Code"
      Text            =   ""
   End
   Begin VB.Label lblPrimaryCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmSecondaryCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PRIMARY_CODE_SHOW         As String

Dim SECONDARY_CODE_NEW_ENTRY_    As Boolean
Dim GRID_BROWSE_                 As Boolean
Private CURRENT_SECONDARY_CODE   As String

Private Sub cmdCancel_Click()
   ComponentEnableMode Me, NewEntryMode, gridMain
   ComponentEnableMode Me, DefaultMode, gridMain
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   If gridMain.row = 0 Then
       ShowMsgInTaskBar "Please select a secondary code to continue.", , "N"
       Exit Sub
   End If
    If UCase(txtCode.text) = "LL" And gridMain.TextMatrix(gridMain.row, 0) = "LANDLORD" Then '
         MsgBox "Can not delete landlord type."
         Exit Sub
   End If
   If UCase(txtCode.text) = "MA" And gridMain.TextMatrix(gridMain.row, 0) = "MANAGING AGENT" Then '
        MsgBox "Can not delete managing agent type"
        Exit Sub
   End If
    If UCase(txtCode.text) = UCase("Supplier") And gridMain.TextMatrix(gridMain.row, 0) = "SUPPLIER CODE" Then '
        MsgBox "Can not delete Supplier type"
        Exit Sub
   End If
   
   If MsgBox("Are you sure to delete the code?", vbQuestion + vbYesNo, "Delete Secondary Code") = vbNo Then Exit Sub

   Dim Conn1 As New ADODB.Connection
   Dim sSQLStr As String

   On Error GoTo errHnd

   Conn1.Open getConnectionString

   sSQLStr = "DELETE * " & _
             "FROM SecondaryCode " & _
             "WHERE PrimaryCode = '" & PRIMARY_CODE_SHOW & "' AND " & _
                  "Code= '" & gridMain.TextMatrix(gridMain.row, 1) & "';"

   Conn1.Execute sSQLStr
   Conn1.Close
   Set Conn1 = Nothing

   LoadGrid
   ComponentEnableMode Me, DefaultMode, gridMain
   
   ShowMsgInTaskBar "Secondary code has been deleted successfully."
   Exit Sub
   
errHnd:
   ShowMsgInTaskBar Err.Number & " - " & Err.description & "; Secondary code has not been deleted successfully.", , "N"
   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub cmdEdit_Click()
   If gridMain.TextMatrix(gridMain.row, 0) = "" Then
       ShowMsgInTaskBar "Please select a record to continue"
       Exit Sub
   End If
   If txtCode.text = "LL" Then
         MsgBox "Can not edit landlord type."
         Exit Sub
   End If
   If txtCode.text = "MA" Then
        MsgBox "Can not edit managing agent type"
        Exit Sub
   End If
   SECONDARY_CODE_NEW_ENTRY_ = False
   CURRENT_SECONDARY_CODE = txtCode.text
   ComponentEnableMode Me, EditMode, gridMain
   txtValue.SetFocus
End Sub

Private Sub cmdNew_Click()
   If PRIMARY_CODE_SHOW = "" Then
    ShowMsgInTaskBar "Please select a primary code to continue"
    Exit Sub
   End If
   SECONDARY_CODE_NEW_ENTRY_ = True
   ComponentEnableMode Me, NewEntryMode, gridMain
   txtValue.SetFocus
End Sub

Private Sub cmdSave_Click()
   If txtValue.text = "" Then
      ShowMsgInTaskBar "Please Input the " & Label3.Caption & ".", , "N"
      txtValue.SetFocus
      Exit Sub
   End If
   If txtCode.text = "" Then
      ShowMsgInTaskBar "Please Input the " & Label2.Caption & ".", , "N"
      txtCode.SetFocus
      Exit Sub
   End If
   If PRIMARY_CODE_SHOW = "MNTJOB" And Not ValidateEmail(txtDescription.text) Then
      ShowMsgInTaskBar "Please enter a valid email address in the description field.", "Y", "N"
      txtDescription.SetFocus
      Exit Sub
   End If

   Dim sSQLQuery_ As String
   Dim iRow       As Integer

   For iRow = 1 To gridMain.Rows - 1
      If SECONDARY_CODE_NEW_ENTRY_ Then
         If UCase(txtCode.text) = UCase(gridMain.TextMatrix(iRow, 1)) Then
            ShowMsgInTaskBar "Same code has been found in the grid.", "Y", "N"
            Exit Sub
         End If
         If UCase(txtValue.text) = UCase(gridMain.TextMatrix(iRow, 2)) Then
            ShowMsgInTaskBar "Same Name has been found in the grid.", "Y", "N"
            Exit Sub
         End If
      End If
   Next iRow

   adoMain.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT * " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = '" & PRIMARY_CODE_SHOW & "' " & _
                     "AND CODE = '" & CURRENT_SECONDARY_CODE & "';"
'Debug.Print sSQLQuery_
   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If SECONDARY_CODE_NEW_ENTRY_ Then
      SaveCodes adoMain
   Else
      UpdateCodes adoMain
   End If

   ComponentEnableMode Me, DefaultMode, gridMain
   LoadGrid
   CURRENT_SECONDARY_CODE = ""
End Sub

Public Function SaveCodes(ByVal adoConnector As Adodc) As Boolean
   If adoConnector.Recordset.RecordCount > 0 Then
       ShowMsgInTaskBar "The ID specified already exists."
       SaveCodes = False
       Exit Function
   End If

   On Error GoTo ErrorHandler

   adoConnector.Recordset.AddNew
   
   adoConnector.Recordset.Fields(0).Value = PRIMARY_CODE_SHOW
   adoConnector.Recordset.Fields(1).Value = txtCode.text
   adoConnector.Recordset.Fields(2).Value = txtValue.text
   adoConnector.Recordset.Fields(3).Value = txtDescription.text

   adoConnector.Recordset.Update
   adoConnector.Refresh
   SaveCodes = True

   ShowMsgInTaskBar "The record saved successfully."

   Exit Function
ErrorHandler:
   ShowMsgInTaskBar Err.Number & " " & Err.description, , "N"
End Function

Public Function UpdateCodes(ByVal adoConnector As Adodc) As Boolean
On Error GoTo Err
   If adoConnector.Recordset.RecordCount = 0 Then
       If MsgBox("The record does not exists. Do you want to create a new record ", vbYesNo, "Update Record") = vbYes Then
           If SaveCodes(adoConnector) Then
               UpdateCodes = True
               Exit Function
           End If
       Else
           UpdateCodes = False
           Exit Function
       End If
   End If

   adoConnector.Recordset.Fields(0).Value = PRIMARY_CODE_SHOW
   adoConnector.Recordset.Fields(1).Value = txtCode.text
   adoConnector.Recordset.Fields(2).Value = txtValue.text
   adoConnector.Recordset.Fields(3).Value = txtDescription.text

   adoConnector.Recordset.Update
   adoConnector.Refresh

   UpdateUserRecords

   ShowMsgInTaskBar "The record updated successfully."
   Exit Function
Err:
   If Err.Number = -2147217900 Then
      MsgBox "Duplicate Short Name, update failed.", vbInformation, "Not Saved"
   End If
End Function

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 6705
   Me.Width = 7725
   Me.BackColor = MODULEBACKCOLOR
   fraCodes.BackColor = Me.BackColor
   fraCodes.Top = -15
   fraCodes.Left = 0
'
'   PopulateCode
'   ComponentEnableMode Me, DefaultMode, gridMain
'   adoMain.ConnectionString = getConnectionString
   
   LoadGrid

   gridMain.ColWidth(0) = Label5(1).Left - Label5(0).Left
   gridMain.ColWidth(1) = Label5(2).Left - Label5(1).Left
   gridMain.ColWidth(2) = Label5(3).Left - Label5(2).Left
   gridMain.ColWidth(3) = gridMain.Width + gridMain.Left - Label5(3).Left - 360
   
   If PRIMARY_CODE_SHOW = "MNTJOB" Then
      Label4.Caption = "Email:"
   Else
      Label4.Caption = "Description:"
   End If
   Label5(3).Caption = Label4.Caption
End Sub

Public Function PopulateCode()
   Dim sSQLQuery_ As String

   adoPrimaryCode1.ConnectionString = getConnectionString

   If PRIMARY_CODE_SHOW <> "" Then

      sSQLQuery_ = "SELECT CODE, VALUE " & _
                   "FROM PRIMARYCODE " & _
                   "WHERE Code = '" & PRIMARY_CODE_SHOW & "' AND " & _
                   "Flexible = TRUE " & _
                   "ORDER BY VALUE;"

   Else
      lblPrimaryCode.Visible = True
      fraCodes.Top = 480
      frmSecondaryCode.Height = frmSecondaryCode.Height + fraCodes.Top + 100


      sSQLQuery_ = "SELECT CODE, VALUE " & _
                   "FROM PRIMARYCODE " & _
                   "WHERE Flexible = TRUE " & _
                   "ORDER BY VALUE;"
   End If
   
   adoPrimaryCode1.RecordSource = sSQLQuery_
   adoPrimaryCode1.CommandType = adCmdText
   adoPrimaryCode1.Refresh
   
   If PRIMARY_CODE_SHOW <> "" Then
      frmSecondaryCode.Caption = adoPrimaryCode1.Recordset(1)
      lblCaption.Caption = frmSecondaryCode.Caption
   End If
End Function

Public Function LoadGrid()
   Dim sSQLQuery_ As String

   If PRIMARY_CODE_SHOW <> "" Then
      sSQLQuery_ = "SELECT PRIMARYCODE.VALUE, SECONDARYCODE.CODE, " & _
                    "SECONDARYCODE.VALUE, SECONDARYCODE.DESCRIPTION " & _
                    "FROM PRIMARYCODE, SECONDARYCODE " & _
                    "WHERE Flexible = TRUE AND PRIMARYCODE.CODE = '" & PRIMARY_CODE_SHOW & "' " & _
                    "AND PRIMARYCODE.CODE = SECONDARYCODE.PRIMARYCODE ORDER BY SECONDARYCODE.CODE"
   Else
      sSQLQuery_ = "SELECT PRIMARYCODE.VALUE, SECONDARYCODE.CODE, " & _
                    "SECONDARYCODE.VALUE, SECONDARYCODE.DESCRIPTION " & _
                    "FROM PRIMARYCODE, SECONDARYCODE " & _
                   "WHERE Flexible = TRUE AND PRIMARYCODE.CODE = SECONDARYCODE.PRIMARYCODE " & _
                   "ORDER BY SECONDARYCODE.CODE;"
   
   End If

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   populateGrid adoConn, sSQLQuery_, gridMain

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   PRIMARY_CODE_SHOW = ""
   lblPrimaryCode.Visible = False
   fraCodes.Top = 0
   frmSecondaryCode.Height = fraCodes.Height
   Unload Me
End Sub

Private Sub gridMain_Click()
   GRID_BROWSE_ = True
   ComponentEnableMode Me, GridRowOnSelection, gridMain
   populateControl frmSecondaryCode, gridMain
End Sub

Private Sub gridMain_GotFocus()
   GRID_BROWSE_ = True
End Sub

Private Sub gridMain_LostFocus()
   GRID_BROWSE_ = False
End Sub

Private Sub gridMain_RowColChange()
   GRID_BROWSE_ = True
   populateControl frmSecondaryCode, gridMain

End Sub

Private Sub UpdateUserRecords()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If PRIMARY_CODE_SHOW = "MNTJOB" Then
      adoConn.Execute "UPDATE PropertyMaintHistory " & _
                      "SET    TaskOwner = '" & txtCode.text & "' " & _
                      "WHERE  TaskOwner = '" & CURRENT_SECONDARY_CODE & "';"
      adoConn.Execute "UPDATE PropertyMaintHistory " & _
                      "SET    AssignedTo = '" & txtCode.text & "' " & _
                      "WHERE  AssignedTo = '" & CURRENT_SECONDARY_CODE & "';"
      adoConn.Execute "UPDATE PropertyMaintHistory " & _
                      "SET    ReportedBy = '" & txtCode.text & "' " & _
                      "WHERE  ReportedBy = '" & CURRENT_SECONDARY_CODE & "';"
   End If

   If PRIMARY_CODE_SHOW = "MTYP" Then
      adoConn.Execute "UPDATE PropertyMaintHistory " & _
                      "SET    MaintenanceType = '" & txtCode.text & "' " & _
                      "WHERE  MaintenanceType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "LTYP" Then
      adoConn.Execute "UPDATE LeaseDetails " & _
                      "SET    TypeOfStore = '" & txtValue.text & "' " & _
                      "WHERE  TypeOfStore = '" & gridMain.TextMatrix(gridMain.row, 2) & "';"
   End If
'-----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "UUSE" Then
      adoConn.Execute "UPDATE LeaseDetails " & _
                      "SET    Usage = '" & txtValue.text & "' " & _
                      "WHERE  Usage = '" & gridMain.TextMatrix(gridMain.row, 2) & "';"
      adoConn.Execute "UPDATE PropertyInsurance " & _
                      "SET    Usage = '" & txtCode.text & "' " & _
                      "WHERE  Usage = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "IRER" Then
      adoConn.Execute "UPDATE PropertyInsurance " & _
                      "SET    Insurer = '" & txtCode.text & "' " & _
                      "WHERE  Insurer = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "ITYP" Then
      adoConn.Execute "UPDATE PropertyInsurance " & _
                      "SET    InsuranceType = '" & txtCode.text & "' " & _
                      "WHERE  InsuranceType = '" & CURRENT_SECONDARY_CODE & "';"
   End If

   If PRIMARY_CODE_SHOW = "BTYP" Then
      adoConn.Execute "UPDATE LeaseBreaches " & _
                      "SET    BreachType = '" & txtCode.text & "' " & _
                      "WHERE  BreachType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'-----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "ATYP" Then
      adoConn.Execute "UPDATE UnitAnalysis " & _
                      "SET    AnalysisType = '" & txtCode.text & "' " & _
                      "WHERE  AnalysisType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'-----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "TNTYPE" Then
      adoConn.Execute "UPDATE Units " & _
                      "SET    Management = '" & txtCode.text & "' " & _
                      "WHERE  Management = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'-----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "IPT" Then
      adoConn.Execute "UPDATE UnitSafety " & _
                      "SET    InspectedBy = '" & txtCode.text & "' " & _
                      "WHERE  InspectedBy = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "STYP" Then
      adoConn.Execute "UPDATE UnitSafety " & _
                      "SET    SafetyType = '" & txtCode.text & "' " & _
                      "WHERE  SafetyType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'-----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "UTIL" Then
      adoConn.Execute "UPDATE UnitUtilities " & _
                      "SET    UtilitiesType = '" & txtCode.text & "' " & _
                      "WHERE  UtilitiesType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "UTIL" Then
      adoConn.Execute "UPDATE PropertyUtilities " & _
                      "SET    UtilitiesType = '" & txtCode.text & "' " & _
                      "WHERE  UtilitiesType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "USTA" Then
      adoConn.Execute "UPDATE UnitUtilities " & _
                      "SET    Status = '" & txtCode.text & "' " & _
                      "WHERE  Status = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "USTA" Then
      adoConn.Execute "UPDATE PropertyUtilities " & _
                      "SET    Status = '" & txtCode.text & "' " & _
                      "WHERE  Status = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "ACCT" Then
      adoConn.Execute "UPDATE Supplier " & _
                      "SET    AccountType = '" & txtCode.text & "' " & _
                      "WHERE  AccountType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "RAT" Then
      adoConn.Execute "UPDATE Supplier " & _
                      "SET    PaymentType = '" & txtCode.text & "' " & _
                      "WHERE  PaymentType = '" & CURRENT_SECONDARY_CODE & "';"
      adoConn.Execute "UPDATE tlbPayment " & _
                      "SET    PayAmtType = '" & txtCode.text & "' " & _
                      "WHERE  PayAmtType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
   If PRIMARY_CODE_SHOW = "SCODE" Then
      adoConn.Execute "UPDATE Supplier " & _
                      "SET    SupplierType = '" & txtCode.text & "' " & _
                      "WHERE  SupplierType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'----------------------------------------------------------------------------------------
   If PRIMARY_CODE_SHOW = "DPTYP" Or PRIMARY_CODE_SHOW = "RTYP" Or PRIMARY_CODE_SHOW = "EXPTYP" Then
      adoConn.Execute "UPDATE TenantDeposit " & _
                      "SET    DptType = '" & txtCode.text & "' " & _
                      "WHERE  DptType = '" & CURRENT_SECONDARY_CODE & "';"
   End If
'
   adoConn.Close
   Set adoConn = Nothing
End Sub
