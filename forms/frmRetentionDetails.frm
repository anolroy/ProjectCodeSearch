VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRetentionDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retention Details"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   14670
   Begin VB.Frame Frame2 
      Height          =   9510
      Left            =   14580
      TabIndex        =   1
      Top             =   8055
      Visible         =   0   'False
      Width           =   14505
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   2250
         TabIndex        =   9
         Top             =   8505
         Width           =   14370
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   420
            Left            =   12690
            TabIndex        =   11
            Top             =   405
            Width           =   1365
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   10125
            TabIndex        =   10
            Top             =   450
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6720
         Left            =   1530
         TabIndex        =   7
         Top             =   2205
         Width           =   14370
         Begin VB.CommandButton cmdEditLine 
            Caption         =   "&Edit Line"
            Height          =   375
            Left            =   6390
            TabIndex        =   16
            Top             =   5490
            Width           =   2400
         End
         Begin VB.CommandButton cmdDeleteLine 
            Caption         =   "Delete Line"
            Height          =   375
            Left            =   3960
            TabIndex        =   15
            Top             =   5490
            Width           =   2400
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetails 
            Height          =   4485
            Left            =   135
            TabIndex        =   8
            Top             =   720
            Width           =   14070
            _ExtentX        =   24818
            _ExtentY        =   7911
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   15329508
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483630
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   195
            Left            =   12150
            TabIndex        =   14
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   2250
            TabIndex        =   13
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SL"
            Height          =   195
            Left            =   945
            TabIndex        =   12
            Top             =   315
            Width           =   195
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1545
         Left            =   945
         TabIndex        =   2
         Top             =   495
         Width           =   14415
         Begin VB.TextBox txtRetensionAmount1 
            Alignment       =   1  'Right Justify
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
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   11205
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1080
            Width           =   1125
         End
         Begin VB.TextBox txtRetentionDescriptions 
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
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2295
            MaxLength       =   50
            TabIndex        =   0
            Top             =   1080
            Width           =   8820
         End
         Begin VB.CommandButton cmdAddToGrid 
            Caption         =   "OK"
            Height          =   420
            Left            =   12555
            TabIndex        =   3
            Top             =   990
            Width           =   1275
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   210
            Left            =   11250
            TabIndex        =   6
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   2340
            TabIndex        =   5
            Top             =   810
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   7710
      Left            =   45
      TabIndex        =   17
      Top             =   0
      Width           =   14595
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2475
         TabIndex        =   30
         Top             =   6345
         Width           =   1275
      End
      Begin VB.CommandButton cmadAddRetention1 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1125
         TabIndex        =   29
         Top             =   6345
         Width           =   1275
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "&Clear All"
         Height          =   375
         Left            =   4500
         TabIndex        =   28
         Top             =   6345
         Width           =   1275
      End
      Begin VB.CommandButton cmdClearSelected 
         Caption         =   "&Clear Selected"
         Height          =   375
         Left            =   5895
         TabIndex        =   27
         Top             =   6345
         Width           =   1275
      End
      Begin VB.CommandButton cmdUnclear 
         Caption         =   "&Unclear"
         Height          =   375
         Left            =   7335
         TabIndex        =   26
         Top             =   6345
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   420
         Left            =   13005
         TabIndex        =   25
         Top             =   6840
         Width           =   1365
      End
      Begin VB.TextBox txtBalance2 
         Alignment       =   1  'Right Justify
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
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13095
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   23
         Top             =   6345
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetailsMain 
         Height          =   5610
         Left            =   45
         TabIndex        =   18
         Top             =   540
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   9895
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483630
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SL No"
         Height          =   195
         Left            =   1485
         TabIndex        =   31
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "Balance"
         Height          =   285
         Left            =   12060
         TabIndex        =   24
         Top             =   6390
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   13230
         TabIndex        =   22
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Left            =   11610
         TabIndex        =   21
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   2835
         TabIndex        =   20
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Statement ID"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   225
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmRetentionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public szCurrentStatementID As String
Public szClienyIDForRetention As String
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
Private Sub cmadAddRetention1_Click()
'    Frame2.Top = 0
'    Frame2.Left = 1575
'    Frame2.Visible = True
     frmAddRetention.Show
     frmAddRetention.Top = Me.Top + 500
End Sub

Private Sub cmdAddToGrid_Click()
            If Val(txtRetensionAmount1.text) = 0 Then
                    MsgBox "Please enter amount greater than zero", vbInformation, "Warning"
                    Exit Sub
            End If
            flxRetensionDetails.Enabled = True
            'Enter data into grid only memory version
            'statementId you shall generate it when you finally save the statement
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = "" 'IIf(Option1.Value = True, "+", "-")
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
            flxRetensionDetails.AddItem ""
           ' txtRetensionAmount1.Visible = False
            
            txtRetensionAmount1.text = "0.00"
            FocusControl txtRetentionDescriptions
            txtRetensionAmount1.SelStart = 0
            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
            Call MakeSummaryRetention
End Sub

Private Sub cmdClearAll_Click()
    Dim iRow As Integer
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    For iRow = 1 To flxRetensionDetailsMain.Rows - 1
        If flxRetensionDetailsMain.TextMatrix(iRow, 1) <> "" Then
                adoConn.Execute "Update RetentionDetails  set isCleared =true where  StatementID=" & Mid(flxRetensionDetailsMain.TextMatrix(iRow, 1), 3, Len(flxRetensionDetailsMain.TextMatrix(iRow, 1)) - 2) & ""
        End If
    Next
    Call LoadflxRetensionDetailsMain
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdClearSelected_Click()
    Dim iRow As Integer
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    For iRow = 1 To flxRetensionDetailsMain.Rows - 1
        If flxRetensionDetailsMain.TextMatrix(iRow, 0) = "X" Then
                adoConn.Execute "Update RetentionDetails  set isCleared =true where  statementID=" & Mid(flxRetensionDetailsMain.TextMatrix(iRow, 1), 3, Len(flxRetensionDetailsMain.TextMatrix(iRow, 1)) - 2) & ""
        End If
    Next
    Call LoadflxRetensionDetailsMain
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdClose_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdDeleteLine_Click()
    If flxRetensionDetails.row = 0 Then
        MsgBox "Please select a row to delete"
        Exit Sub
    End If
    flxRetensionDetails.RemoveItem flxRetensionDetails.row
End Sub

Private Sub cmdSave_Click()
    Dim iCol As Integer
    Dim iRow As Integer
    
    frmRentPayableNew.ConfigflxRetensionDetails
    frmRentPayableNew.flxRetensionDetails.Cols = flxRetensionDetails.Cols
    frmRentPayableNew.flxRetensionDetails.Rows = flxRetensionDetails.Rows
    For iRow = 1 To flxRetensionDetails.Rows - 1
        For iCol = 1 To flxRetensionDetails.Cols - 1
             frmRentPayableNew.flxRetensionDetails.TextMatrix(iRow, iCol) = flxRetensionDetails.TextMatrix(iRow, iCol)
        Next
    Next
    MsgBox "Retention Details data has been saved temporarily", vbInformation, "Data saved"
    'Unload Me
End Sub
Public Sub ConfigflxRetensionDetailsMain()
        flxRetensionDetailsMain.Clear
        Dim szHeader As String
        szHeader$ = "|<StatementID|<Description|<SlNumber|<Amount"
        flxRetensionDetailsMain.FormatString = szHeader$

        flxRetensionDetailsMain.Cols = 6
        flxRetensionDetailsMain.Rows = 2
        flxRetensionDetailsMain.RowHeight(0) = 0
        flxRetensionDetailsMain.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxRetensionDetailsMain.ColWidth(1) = 1200 'This is statementId
        flxRetensionDetailsMain.ColWidth(2) = 1200 'This is slNumber
        flxRetensionDetailsMain.ColWidth(3) = 8900  'This is Description
        flxRetensionDetailsMain.ColWidth(4) = 1200  'This is amount
        flxRetensionDetailsMain.ColWidth(5) = 1200  'This is status
        flxRetensionDetailsMain.ColAlignment(3) = vbLeftJustify
End Sub
Public Sub ConfigflxRetensionDetails()
        flxRetensionDetails.Clear
        Dim szHeader As String
        szHeader$ = "|<StatementID|<SL No|<Description|<SlNumber|<Amount"
        flxRetensionDetails.FormatString = szHeader$

        flxRetensionDetails.Cols = 5
        flxRetensionDetails.Rows = 2
        flxRetensionDetails.RowHeight(0) = 0
        flxRetensionDetails.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxRetensionDetails.ColWidth(1) = 1200 'This is statementId
        flxRetensionDetails.ColWidth(2) = 1200 'This is slNumber
        flxRetensionDetails.ColWidth(3) = 1200  'This is Description
        flxRetensionDetails.ColWidth(4) = 1200  'This is amount
        flxRetensionDetails.ColAlignment(3) = vbLeftJustify
End Sub

Private Sub cmdUnclear_Click()
    Dim iRow As Integer
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    For iRow = 1 To flxRetensionDetailsMain.Rows - 1
        If flxRetensionDetailsMain.TextMatrix(iRow, 0) = "X" Then
                adoConn.Execute "Update RetentionDetails  set isCleared =false  where  statementID=" & Mid(flxRetensionDetailsMain.TextMatrix(iRow, 1), 3, Len(flxRetensionDetailsMain.TextMatrix(iRow, 1)) - 2) & ""
        End If
    Next
    Call LoadflxRetensionDetailsMain
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub flxRetensionDetailsMain_Click()
     SelectFlxGridRow 0, flxRetensionDetailsMain, flxRetensionDetailsMain.row
End Sub

Private Sub txtRetensionAmount1_GotFocus()
     txtRetensionAmount1.SelStart = 0
     txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
End Sub
Private Sub MakeSummaryRetention()
    Dim iRow As Long
    Dim dblAmt As Double
    For iRow = 1 To flxRetensionDetails.Rows - 1
            If flxRetensionDetails.TextMatrix(iRow, 2) <> "" Then
'                    If flxRetensionDetails.TextMatrix(iRow, 0) = "+" Then
                            dblAmt = dblAmt + Val(flxRetensionDetails.TextMatrix(iRow, 4))
'                    Else
'                            dblAmt = dblAmt - flxRetensionDetails.TextMatrix(iRow, 4)
'                    End If
            End If
        Next
    frmRentPayableNew.txtRetention.text = dblAmt
End Sub
Private Sub txtRetensionAmount1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
'            flxRetensionDetails.Enabled = True
'            'Enter data into grid only memory version
'            'statementId you shall generate it when you finally save the statement
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = IIf(Option1.Value = True, "+", "-")
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
'            flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
'            flxRetensionDetails.AddItem ""
'           ' txtRetensionAmount1.Visible = False
'
'            txtRetensionAmount1.text = "0.00"
'            txtRetensionAmount1.SelStart = 0
'            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
'            Call MakeSummaryRetention
        FocusControl cmdAddToGrid
     End If
'     If KeyAscii = 27 Then 'escape ascii key
'        txtRetensionAmount1.Visible = False
'     End If
     DigitTextKeyPress txtRetensionAmount1, KeyAscii
End Sub
Private Sub txtRetensionAmount1_LostFocus()
    txtRetensionAmount1.text = Format(txtRetensionAmount1.text, "0.00")
End Sub
Private Sub Form_Load()
    Frame4.BackColor = MODULEBACKCOLOR
    Label8.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Frame6.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR
    Frame3.BackColor = MODULEBACKCOLOR
    
    Label1.BackColor = MODULEBACKCOLOR
    Label2.BackColor = MODULEBACKCOLOR
    Label3.BackColor = MODULEBACKCOLOR
    Label14.BackColor = MODULEBACKCOLOR
    Label15.BackColor = MODULEBACKCOLOR
    
    Label4.BackColor = MODULEBACKCOLOR
    Label5.BackColor = MODULEBACKCOLOR
    Label6.BackColor = MODULEBACKCOLOR
    Label7.BackColor = MODULEBACKCOLOR
'    Label8.BackColor = MODULEBACKCOLOR
    Label9.BackColor = MODULEBACKCOLOR
    Call ConfigflxRetensionDetails
    Call LoadflxRetensionDetailsMain
End Sub
Private Sub LoadflxRetensionDetailsMain()
     Dim adoConn As New ADODB.Connection
     adoConn.Open getConnectionString
     Dim rsRetensionDetails As New ADODB.Recordset
     Dim iRow As Integer
     Call ConfigflxRetensionDetailsMain
     If szClienyIDForRetention = "" Then Exit Sub
     iRow = 1
     rsRetensionDetails.Open "Select * from RetentionDetails D,RentSummaryStatement R where R.StatementID =D.StatementID and ClientIDLandlordID='" & szClienyIDForRetention & "' order by D.statementID", adoConn, adOpenStatic, adLockReadOnly
     While Not rsRetensionDetails.EOF
            flxRetensionDetailsMain.AddItem ""
            flxRetensionDetailsMain.TextMatrix(iRow, 1) = "SS" & rsRetensionDetails("statementID").Value
            flxRetensionDetailsMain.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
            flxRetensionDetailsMain.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
            flxRetensionDetailsMain.TextMatrix(iRow, 4) = Format(rsRetensionDetails("Amount").Value, "0.00")
            flxRetensionDetailsMain.TextMatrix(iRow, 5) = IIf(rsRetensionDetails("isCleared").Value = False, "", "Cleared")
            iRow = iRow + 1
            rsRetensionDetails.MoveNext
     Wend
     rsRetensionDetails.Close
     Set rsRetensionDetails = Nothing
     rsRetensionDetails.Open "Select sum(amount) as amt from RetentionDetails D,RentSummaryStatement R where isCleared=false AND R.StatementID =D.StatementID and ClientIDLandlordID='" & szClienyIDForRetention & "' ", adoConn, adOpenStatic, adLockReadOnly
     If Not rsRetensionDetails.EOF Then
'        txtBalance1.text = Format(rsRetensionDetails("amt").Value, "0.00")
        txtBalance2.text = Format(rsRetensionDetails("amt").Value, "0.00")
     End If
     rsRetensionDetails.Close
     Set rsRetensionDetails = Nothing
     adoConn.Close
     Set adoConn = Nothing
End Sub

'Private Sub ConfigflxRetensionDetails()
'        flxRetensionDetails.Clear
'        Dim szHeader As String
'        szHeader$ = "|<StatementID|<SlNumber|<Amount"
'        flxRetensionDetails.FormatString = szHeader$
'
'        flxRetensionDetails.Cols = 5
'        flxRetensionDetails.Rows = 2
'        flxRetensionDetails.RowHeight(0) = 0
'        flxRetensionDetails.ColWidth(0) = 250   'Selection Row put plus or minus sign
'        flxRetensionDetails.ColWidth(1) = 0 'This is statementId
'        flxRetensionDetails.ColWidth(2) = 1200 'This is slNumber
'        flxRetensionDetails.ColWidth(3) = 7000  'This is Description
'        flxRetensionDetails.ColWidth(4) = 1200  'This is amount
'        flxRetensionDetails.ColAlignment(3) = vbLeftJustify
'
'End Sub

Private Sub LoadflxRetensionDetails()
     Dim adoConn As New ADODB.Connection
     adoConn.Open getConnectionString
     Dim rsRetensionDetails As New ADODB.Recordset
     Dim iRow As Integer
     iRow = 1
     rsRetensionDetails.Open "Select * from RetentionDetails where statementID=" & szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly
     While Not rsRetensionDetails.EOF
            flxRetensionDetails.AddItem ""
            flxRetensionDetails.TextMatrix(iRow, 1) = rsRetensionDetails("statementID").Value
            flxRetensionDetails.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
            flxRetensionDetails.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
            flxRetensionDetails.TextMatrix(iRow, 4) = rsRetensionDetails("Amount").Value
            iRow = iRow + 1
            rsRetensionDetails.MoveNext
     Wend
     rsRetensionDetails.Close
     Set rsRetensionDetails = Nothing
     adoConn.Close
     Set adoConn = Nothing
End Sub

Private Sub txtRetentionDescriptions_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtRetensionAmount1
    End If
End Sub
