VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAddRetention 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Retention"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14550
   Begin VB.Frame Frame2 
      Height          =   8520
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14505
      Begin VB.Frame Frame6 
         Height          =   1185
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   14415
         Begin VB.CommandButton cmdAddToGrid 
            Caption         =   "Add a New Retention Item"
            Height          =   420
            Left            =   11475
            TabIndex        =   2
            Top             =   495
            Width           =   2355
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
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   0
            Top             =   585
            Width           =   8820
         End
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
            Left            =   10125
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "0.00"
            Top             =   585
            Width           =   1125
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   1260
            TabIndex        =   16
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   210
            Left            =   10170
            TabIndex        =   15
            Top             =   315
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6045
         Left            =   135
         TabIndex        =   10
         Top             =   1305
         Width           =   14370
         Begin VB.CommandButton cmdDeleteLine 
            Caption         =   "Delete Line"
            Height          =   375
            Left            =   4590
            TabIndex        =   4
            Top             =   5265
            Width           =   2400
         End
         Begin VB.CommandButton cmdEditLine 
            Caption         =   "&Edit Line"
            Height          =   375
            Left            =   7020
            TabIndex        =   5
            Top             =   5265
            Width           =   2400
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRetensionDetails 
            Height          =   4485
            Left            =   135
            TabIndex        =   3
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SL"
            Height          =   195
            Left            =   810
            TabIndex        =   13
            Top             =   315
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   1350
            TabIndex        =   12
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   195
            Left            =   10260
            TabIndex        =   11
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   90
         TabIndex        =   9
         Top             =   7380
         Width           =   14370
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   375
            Left            =   10125
            TabIndex        =   6
            Top             =   450
            Width           =   1365
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   420
            Left            =   12690
            TabIndex        =   7
            Top             =   405
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frmAddRetention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public szCurrentStatementID As String
Public szClienyIDForRetention As String
Private bEditMode As Boolean
Private iSelectedRow As Integer
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
                    FocusControl txtRetensionAmount1
                    Exit Sub
            End If
            If bEditMode = False Then
                 flxRetensionDetails.Enabled = True
                 'Enter data into grid only memory version
                 'statementId you shall generate it when you finally save the statement
                 flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 0) = "" 'IIf(Option1.Value = True, "+", "-")
                 flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber
                 flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 3) = txtRetentionDescriptions.text 'This is Description
                 flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
                 flxRetensionDetails.AddItem ""
             Else
                 flxRetensionDetails.TextMatrix(iSelectedRow, 0) = "" 'IIf(Option1.Value = True, "+", "-")
                 'flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber we are not changing this
                 flxRetensionDetails.TextMatrix(iSelectedRow, 3) = txtRetentionDescriptions.text 'This is Description
                 flxRetensionDetails.TextMatrix(iSelectedRow, 4) = Format(Val(txtRetensionAmount1.text), "0.00") 'This is amount
             End If
                 
            txtRetensionAmount1.text = "0.00"
            FocusControl txtRetentionDescriptions
            txtRetensionAmount1.SelStart = 0
            txtRetensionAmount1.SelLength = Len(txtRetensionAmount1.text)
            Call MakeSummaryRetention
            txtRetentionDescriptions.text = ""
            txtRetensionAmount1.text = "0.00"
            bEditMode = False
End Sub




Private Sub cmdClose_Click()
    bEditMode = False
    Unload Me
End Sub

Private Sub cmdDeleteLine_Click()
    If flxRetensionDetails.row = 0 Then
        MsgBox "Please select a row to delete"
        Exit Sub
    End If

    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    adoconn.Execute "Delete from RetentionDetailsPrev where slnumber=" & _
            flxRetensionDetails.TextMatrix(flxRetensionDetails.row, 2) & " and StatementID=1"
            adoconn.Close
    Call loadflxRetensionDetails
End Sub

Private Sub cmdEditLine_Click()
    bEditMode = True
    'flxRetensionDetails.TextMatrix(iSelectedRow, 0) = "" 'IIf(Option1.Value = True, "+", "-")
    'flxRetensionDetails.TextMatrix(flxRetensionDetails.Rows - 1, 2) = flxRetensionDetails.Rows - 1 'This is slNumber we are not changing this
    txtRetentionDescriptions.text = flxRetensionDetails.TextMatrix(flxRetensionDetails.row, 3)   'This is Description
    txtRetensionAmount1.text = Format(Val(flxRetensionDetails.TextMatrix(flxRetensionDetails.row, 4)), "0.00") 'This is amount
End Sub

Private Sub cmdSave_Click()
    Dim iCol As Integer
    Dim iRow As Integer
'    bEditMode = False
'    frmRentPayableNew.ConfigflxRetensionDetails
'    frmRentPayableNew.flxRetensionDetails.Cols = flxRetensionDetails.Cols
'    frmRentPayableNew.flxRetensionDetails.Rows = flxRetensionDetails.Rows
'    For iRow = 1 To flxRetensionDetails.Rows - 1
'        For iCol = 1 To flxRetensionDetails.Cols - 1
'             frmRentPayableNew.flxRetensionDetails.TextMatrix(iRow, iCol) = flxRetensionDetails.TextMatrix(iRow, iCol)
'        Next
'    Next
'    MsgBox "Retention Details data has been saved temporarily", vbInformation, "Data saved"
'    Unload Me
    If frmRentPayableNew.szSelectedClient = "" Then
        MsgBox "Please select a client", vbInformation, "Select a client"
        Exit Sub
    End If
    Dim adoconn As New ADODB.Connection
    Dim rsRetentionPrev As New ADODB.Recordset
    Dim rsRetentionPrevMAX As New ADODB.Recordset
    Dim intMaxSlNumber As Integer
    adoconn.Open getConnectionString
    rsRetentionPrevMAX.Open "Select max(SlNumber) as SL from RetentionDetailsPrev", adoconn, adOpenKeyset, adLockOptimistic
    If IsNull(rsRetentionPrevMAX) Then
            intMaxSlNumber = 1
    ElseIf rsRetentionPrevMAX.EOF Then
            intMaxSlNumber = 1
    ElseIf Not rsRetentionPrevMAX.EOF Then
            intMaxSlNumber = rsRetentionPrevMAX("SL").Value
    End If
    
    
    rsRetentionPrev.Open "Select * from RetentionDetailsPrev", adoconn, adOpenKeyset, adLockOptimistic
    rsRetentionPrev.AddNew
    rsRetentionPrev!statementID = "1"
    rsRetentionPrev!slnumber = intMaxSlNumber
    rsRetentionPrev!description = txtRetentionDescriptions.text
    rsRetentionPrev!amount = Val(txtRetensionAmount1.text)
    rsRetentionPrev!isCleared = False
    rsRetentionPrev!ClientID = frmRentPayableNew.szSelectedClient
    rsRetentionPrev.Update
    adoconn.Close
    Call loadflxRetensionDetails

End Sub

Public Sub ConfigflxRetensionDetails()
        flxRetensionDetails.Clear
        Dim szHeader As String
        szHeader$ = "|<StatementID|<Description|<SlNumber|<Amount"
        flxRetensionDetails.FormatString = szHeader$

        flxRetensionDetails.Cols = 6
        flxRetensionDetails.Rows = 2
        flxRetensionDetails.RowHeight(0) = 0
        flxRetensionDetails.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxRetensionDetails.ColWidth(1) = 0 'This is statementId
        flxRetensionDetails.ColWidth(2) = 700 'This is slNumber
        flxRetensionDetails.ColWidth(3) = txtRetentionDescriptions.Width  'This is Description
        flxRetensionDetails.ColWidth(4) = 1200  'This is amount
        flxRetensionDetails.ColAlignment(4) = vbRightJustify
        flxRetensionDetails.ColWidth(5) = 1200
        flxRetensionDetails.ColAlignment(3) = vbLeftJustify
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub flxRetensionDetails_Click()
    iSelectedRow = flxRetensionDetails.row
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
    frmRentPayableNew.txtRetention.text = Format(dblAmt, "0.00")
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
'    Frame4.BackColor = MODULEBACKCOLOR
'    Frame5.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    Frame6.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR
    Frame3.BackColor = MODULEBACKCOLOR
    
    Label1.BackColor = MODULEBACKCOLOR
    Label2.BackColor = MODULEBACKCOLOR
    Label3.BackColor = MODULEBACKCOLOR
    Label14.BackColor = MODULEBACKCOLOR
    Label15.BackColor = MODULEBACKCOLOR
    
    'Label4.BackColor = MODULEBACKCOLOR
    'Label5.BackColor = MODULEBACKCOLOR
'    Label6.BackColor = MODULEBACKCOLOR
'    Label7.BackColor = MODULEBACKCOLOR
'    Label8.BackColor = MODULEBACKCOLOR
'    Label9.BackColor = MODULEBACKCOLOR
   ' Call ConfigflxRetensionDetails
    Call loadflxRetensionDetails
End Sub
Private Sub loadflxRetensionDetails()
     Dim adoconn As New ADODB.Connection
     adoconn.Open getConnectionString
     Dim rsRetensionDetails As New ADODB.Recordset
     Dim iRow As Integer
     Call ConfigflxRetensionDetails
     'If szClienyIDForRetention = "" Then Exit Sub
     iRow = 1
     rsRetensionDetails.Open "Select * from RetentionDetailsPrev D  order by D.statementID", adoconn, adOpenStatic, adLockReadOnly
            
'     rsRetensionDetails.Open "Select * from RetentionDetailsPrev D where ClientID='" & _
'            szClienyIDForRetention & "' order by D.statementID", adoConn, adOpenStatic, adLockReadOnly
     While Not rsRetensionDetails.EOF
            flxRetensionDetails.AddItem ""
            flxRetensionDetails.TextMatrix(iRow, 1) = rsRetensionDetails("statementID").Value
            flxRetensionDetails.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
            flxRetensionDetails.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
            flxRetensionDetails.TextMatrix(iRow, 4) = Format(rsRetensionDetails("Amount").Value, "0.00")
            flxRetensionDetails.TextMatrix(iRow, 5) = IIf(rsRetensionDetails("isCleared").Value = False, "", "Cleared")
            iRow = iRow + 1
            rsRetensionDetails.MoveNext
     Wend
     rsRetensionDetails.Close
     Set rsRetensionDetails = Nothing
     
'     rsRetensionDetails.Open "Select sum(amount) as amt from RetentionDetailsPrev D where " & _
'     "ClientID='" & szClienyIDForRetention & "' ", adoConn, adOpenStatic, adLockReadOnly
'     If Not rsRetensionDetails.EOF Then
'            txtBalance2.text = Format(rsRetensionDetails("amt").Value, "0.00")
'     End If
'     rsRetensionDetails.Close
'     Set rsRetensionDetails = Nothing
     adoconn.Close
     Set adoconn = Nothing
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

'Private Sub LoadflxRetensionDetails()
'     Dim adoConn As New ADODB.Connection
'     adoConn.Open getConnectionString
'     Dim rsRetensionDetails As New ADODB.Recordset
'     Dim iRow As Integer
'     iRow = 1
'     rsRetensionDetails.Open "Select * from RetentionDetails where statementID=" & szCurrentStatementID & "", adoConn, adOpenStatic, adLockReadOnly
'     While Not rsRetensionDetails.EOF
'            flxRetensionDetails.AddItem ""
'            flxRetensionDetails.TextMatrix(iRow, 1) = rsRetensionDetails("statementID").Value
'            flxRetensionDetails.TextMatrix(iRow, 2) = rsRetensionDetails("SLNumber").Value
'            flxRetensionDetails.TextMatrix(iRow, 3) = rsRetensionDetails("Description").Value
'            flxRetensionDetails.TextMatrix(iRow, 4) = rsRetensionDetails("Amount").Value
'            iRow = iRow + 1
'            rsRetensionDetails.MoveNext
'     Wend
'     rsRetensionDetails.Close
'     Set rsRetensionDetails = Nothing
'     adoConn.Close
'     Set adoConn = Nothing
'End Sub

Private Sub txtRetentionDescriptions_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtRetensionAmount1
    End If
End Sub

