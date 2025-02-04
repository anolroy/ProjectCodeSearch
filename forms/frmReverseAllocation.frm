VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReverseAllocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reverse Allocation - xxx\ABC"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
   Icon            =   "frmReverseAllocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   15015
   Begin VB.TextBox txtPayDt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   10305
      TabIndex        =   29
      Top             =   5130
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAmendAllocationDate 
      Caption         =   "&Apply New  Allocation Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9135
      TabIndex        =   24
      Top             =   7470
      Width           =   2460
   End
   Begin VB.CheckBox chkCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   495
      Width           =   195
   End
   Begin VB.CommandButton cmdRevereseAllocation 
      Caption         =   "&Reverse Allocation"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7605
      Width           =   1700
   End
   Begin VB.CommandButton cmdSPClose 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSPayment 
      Height          =   2325
      Left            =   120
      TabIndex        =   1
      Top             =   4965
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCrPoA 
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   5900
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   12648447
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   10215
      TabIndex        =   30
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   450
      X2              =   14985
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "transactionID"
      Height          =   195
      Left            =   4545
      TabIndex        =   28
      Top             =   90
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   660
      TabIndex        =   22
      Top             =   510
      Width           =   240
   End
   Begin MSForms.TextBox txtTenantID 
      Height          =   315
      Left            =   2040
      TabIndex        =   26
      Top             =   9105
      Width           =   2175
      VariousPropertyBits=   746604575
      BackColor       =   15858158
      Size            =   "3836;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Amt £"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   13935
      TabIndex        =   25
      Top             =   4725
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "  Reverse Allocation View  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   2
      Left            =   4680
      TabIndex        =   3
      Top             =   4140
      Width           =   3300
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   1770
      TabIndex        =   21
      Top             =   510
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   2970
      TabIndex        =   20
      Top             =   510
      Width           =   495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   4590
      TabIndex        =   19
      Top             =   510
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   6150
      TabIndex        =   18
      Top             =   510
      Width           =   225
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   8190
      TabIndex        =   17
      Top             =   510
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   11760
      TabIndex        =   16
      Top             =   510
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O/S Amt. £"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   12735
      TabIndex        =   15
      Top             =   4725
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount £"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   11505
      TabIndex        =   14
      Top             =   4725
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   6600
      TabIndex        =   13
      Top             =   4725
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   4980
      TabIndex        =   12
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3525
      TabIndex        =   11
      Top             =   4725
      Width           =   675
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   2625
      TabIndex        =   10
      Top             =   4725
      Width           =   495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   1335
      TabIndex        =   9
      Top             =   4725
      Width           =   345
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No."
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   390
      TabIndex        =   8
      Top             =   4725
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "credit row no"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   4
      Left            =   12390
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "allocating row no"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   3
      Left            =   11895
      TabIndex        =   6
      Top             =   4305
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   75
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   675
      X2              =   14940
      Y1              =   375
      Y2              =   405
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4395
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   2
      Index           =   3
      X1              =   675
      X2              =   13060
      Y1              =   435
      Y2              =   435
   End
End
Attribute VB_Name = "frmReverseAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szLesseeID As String
Dim iCurRow As Integer
Private Sub chkCredits_Click()
        Dim iCount As Integer
        If chkCredits = 1 Then
            For iCount = 1 To flxCrPoA.Rows - 1
                If flxCrPoA.TextMatrix(iCount, 2) = "+" Then
                    flxCrPoA.TextMatrix(iCount, 1) = "X"
                End If
             Next
        Else
             For iCount = 1 To flxCrPoA.Rows - 1
                   flxCrPoA.TextMatrix(iCount, 1) = ""
              Next
        End If
End Sub

Private Function checkOnlyOneSelected() As Boolean
     Dim iCount As Integer
     Dim iSel As Integer
     For iCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(iCount, 1) = "X" Then
            iSel = iSel + 1
        End If
     Next
     If iSel = 1 Then
        checkOnlyOneSelected = True
     Else
        checkOnlyOneSelected = False
     End If
        
End Function
Private Sub cmdAmendAllocationDate_Click()
    On Error GoTo Err
    Dim szDate As String
    Dim iCount As Integer
    Dim adoConn As New ADODB.Connection
    Dim rsAllocation As New ADODB.Recordset
    Dim szAllocdate As Date
    Dim szTransactionDate As Date
    Dim TransactionID As Long
    Dim szAllocationTransactionID As Long
    If checkOnlyOneSelected = False Then
        MsgBox "Please Select one Transaction from the list", vbOKOnly, "Please Select"
        FocusControl flxCrPoA
        Exit Sub
    End If
    If MsgBox("Do you wish to amend Allocation Date?", vbYesNo, "Please confirm") = vbNo Then Exit Sub
    For iCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(iCount, 1) = "X" Then
            TransactionID = flxCrPoA.TextMatrix(iCount, 0)
            szDate = flxCrPoA.TextMatrix(iCount, 5)
            Exit For
        End If
    Next
    adoConn.Open getConnectionString
'    rsAllocation.Open "Select allocdate,P.RDate,T.TransactionID from RptTransactions T,tlbReceipt P where P.transactionId=T.FromTran AND " & _
'            "P.transactionID=" & TransactionID & "", adoConn, adOpenKeyset, adLockReadOnly
'    If Not rsAllocation.EOF Then
'         szAllocdate = rsAllocation("allocdate").Value
'         szTransactionDate = rsAllocation("Rdate").Value
'         frmPostingDate.szAllocationTransactionID = rsAllocation("TransactionID").Value
'    Else
'        MsgBox "Allocation date not found for this transaction", vbInformation, "Warning"
'        Exit Sub
'    End If
'    rsAllocation.Close
    For iCount = 1 To flxSPayment.Rows - 1
            szAllocationTransactionID = flxSPayment.TextMatrix(iCount, 0)
            szDate = flxSPayment.TextMatrix(iCount, 8)
            adoConn.Execute "Update RptTransactions T set allocdate=#" & Format(szDate, "dd MMM yyyy") & "# where T.FromTran=" & TransactionID & "" & _
           "and T.ToTran=" & szAllocationTransactionID & " and T.DeleteFlag=False"
           adoConn.Execute "Update RptTransactionsSplit T set allocdate=#" & Format(szDate, "dd MMM yyyy") & "# where T.FromTran=" & TransactionID & "" & _
           "and T.ToTran=" & szAllocationTransactionID & " and T.DeleteFlag=False"
            
    Next
   
    
    adoConn.Close
    Set adoConn = Nothing
    MsgBox "Allocation Date has been updated", vbInformation, "Completed"
    Exit Sub
Err:
    MsgBox Err.description
End Sub

Private Sub cmdSPClose_Click()
   Unload Me

   frmDemands3.Show
'   frmDemands3.Refresh_Receipt_Rev_Aloc
End Sub

Private Sub flxCrPoA_Click() 'Here we are loading child grid
        Dim szSQL As String, iRow As Integer
        Dim adoConn As New ADODB.Connection
        Dim adoRst As New ADODB.Recordset
        Call SquezeExpand
        If flxCrPoA.TextMatrix(flxCrPoA.row, 1) = "" And flxCrPoA.TextMatrix(flxCrPoA.row, 0) <> "" Then
           flxCrPoA.TextMatrix(flxCrPoA.row, 1) = "X"
                
               If flxCrPoA.TextMatrix(flxCrPoA.row, 0) = "" Then Exit Sub
               adoConn.Open getConnectionString
            
            'FromTran  is receipt
            ' flxCrPoA.TextMatrix(flxCrPoA.row, 0) is the receipt header transaction ID
               szSQL = "SELECT DISTINCT T.DESCRIPTION, R.DemandRef, R.UnitID, R.DDate, R.Details, " & _
                                       "R.Amount, R.OSAmount, RT.ReceiptAmount, R.RDate, R.TransactionID,RT.AllocDate, " & _
                                       "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, R.SlNumber " & _
                       "FROM RptTransactions AS RT, tlbTransactionTypes AS T, tlbReceipt AS R " & _
                       "Where RT.DeleteFlag=false and RT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " And " & _
                                       "RT.ToTran = R.TransactionID And R.Type = t.TYPE_ID " & _
                       "ORDER BY R.TransactionID;"
            '   Here we are loading child grid
            
               adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
               Call ConfigFlxSPayment
               'ConfigFlxSPayment flxSPayment, 11, 9, 10
               'flxSPayment.Cols = 12
               ''flxSPayment.ColWidth(11) = 0
            
               iRow = 1
               While Not adoRst.EOF
                  flxSPayment.TextMatrix(iRow, 0) = adoRst.Fields.Item("TransactionID").Value
                  flxSPayment.TextMatrix(iRow, 2) = adoRst.Fields.Item("PF").Value & adoRst.Fields.Item("SlNumber").Value
                  flxSPayment.TextMatrix(iRow, 3) = adoRst.Fields.Item("DESCRIPTION").Value
                  flxSPayment.TextMatrix(iRow, 4) = adoRst.Fields.Item("UnitID").Value
                  flxSPayment.TextMatrix(iRow, 5) = adoRst.Fields.Item("DDate").Value
                  flxSPayment.TextMatrix(iRow, 6) = adoRst.Fields.Item("RDate").Value '20171107
                  flxSPayment.TextMatrix(iRow, 7) = adoRst.Fields.Item("Details").Value
                  flxSPayment.TextMatrix(iRow, 8) = Format(adoRst.Fields.Item("AllocDate").Value, "dd/MM/yyyy")
                  
                  flxSPayment.TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
                  flxSPayment.TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("OSAmount").Value, "0.00")
                  flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("ReceiptAmount").Value, "0.00")
            '      flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("Fund").Value, "0.00")
            
                  adoRst.MoveNext
                  If Not adoRst.EOF Then
                     iRow = iRow + 1
                     flxSPayment.AddItem ""
                  End If
               Wend
            
               adoRst.Close
               Set adoRst = Nothing
               adoConn.Close
        Else
           flxCrPoA.TextMatrix(flxCrPoA.row, 1) = ""
           Call ConfigFlxSPayment
           'ConfigFlxSPayment flxSPayment, 11, 9, 10
        End If
  If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
        Label1.Caption = "allocated SI transactionID :" & flxSPayment.TextMatrix(flxSPayment.row, 0)
       Label1.Visible = True
   Else
       Label1.Visible = False
   End If
End Sub

Private Sub flxCrPoA_RowColChange()
'   If flxCrPoA.TextMatrix(flxCrPoA.row, 0) = "" Then Exit Sub
'
'   Dim szSQL As String, iRow As Integer
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'
'   adoConn.Open getConnectionString
'
'
'   szSQL = "SELECT DISTINCT T.DESCRIPTION, R.DemandRef, R.UnitID, R.DDate, R.Details, " & _
'                           "R.Amount, R.OSAmount, RT.ReceiptAmount, R.Ref, R.TransactionID, " & _
'                           "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, R.SlNumber " & _
'           "FROM RptTransactions AS RT, tlbTransactionTypes AS T, tlbReceipt AS R " & _
'           "Where RT.FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & " And " & _
'                           "RT.ToTran = R.TransactionID And R.Type = t.TYPE_ID " & _
'           "ORDER BY R.TransactionID;"
''Debug.Print szSQL
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ConfigFlxSPayment flxSPayment, 11, 9, 10
'   flxSPayment.Cols = 12
'   flxSPayment.ColWidth(11) = 0
'
'   iRow = 1
'   While Not adoRst.EOF
'      flxSPayment.TextMatrix(iRow, 0) = adoRst.Fields.Item("TransactionID").Value
'      flxSPayment.TextMatrix(iRow, 2) = adoRst.Fields.Item("PF").Value & adoRst.Fields.Item("SlNumber").Value
'      flxSPayment.TextMatrix(iRow, 3) = adoRst.Fields.Item("DESCRIPTION").Value
'      flxSPayment.TextMatrix(iRow, 4) = adoRst.Fields.Item("UnitID").Value
'      flxSPayment.TextMatrix(iRow, 5) = adoRst.Fields.Item("DDate").Value
'      flxSPayment.TextMatrix(iRow, 6) = adoRst.Fields.Item("Ref").Value
'      flxSPayment.TextMatrix(iRow, 7) = adoRst.Fields.Item("Details").Value
'      flxSPayment.TextMatrix(iRow, 8) = Format(adoRst.Fields.Item("Amount").Value, "0.00")
'      flxSPayment.TextMatrix(iRow, 9) = Format(adoRst.Fields.Item("OSAmount").Value, "0.00")
'      flxSPayment.TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("ReceiptAmount").Value, "0.00")
''      flxSPayment.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("Fund").Value, "0.00")
'
'      adoRst.MoveNext
'      If Not adoRst.EOF Then
'         iRow = iRow + 1
'         flxSPayment.AddItem ""
'      End If
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

Private Sub flxSPayment_Click()
    txtPayDt.Visible = False
End Sub

Private Sub Form_Load()
   Me.Width = 15105
   Me.Height = 8880
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   chkCredits.BackColor = MODULEBACKCOLOR
'   ConfigFlxCrPoA flxCrPoA, 10, 7, 20, "|<No|<Type|<UnitID||<Date|<Ref|Details|>Amt"
   
   Call ConfigFlxSPayment
   'ConfigFlxSPayment flxSPayment, 12, 9, 10, "|<No|<Type|<UnitID|<Date|<Ref||Details|>Amt||"

   Me.Caption = "Reverse Allocation - " & frmDemands3.txtTenantID.text

   Dim szaLessee() As String
   Dim adoConn As New ADODB.Connection

'      connect to database
   adoConn.Open getConnectionString

   szaLessee = Split(frmDemands3.txtTenantID.text, "\")
   szLesseeID = szaLessee(0)

   LoadFlxCrPoA adoConn

   flxCrPoA.row = 0
   flxCrPoA.col = 0
   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxCrPoA(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer
   Dim strWh As String
   
   Dim adoRst As New ADODB.Recordset
   Call ConfigFlxCrPoA
   If frmDemands3.chkExlessee.Value = True Then
        strWh = " AND LeaseDetails.Status = False "
   Else
        strWh = " AND LeaseDetails.Status = True "
   End If
      szSQL = " Select sign,TransactionI,SplitID,Description, INVNO, FundCode, ReceiptAmount, SageAccountNumber, RDate,  ReceiptAmount,ExtRef, " & _
            " TYPE, Name1, Property.PropertyID, UnitID, BankCode, Name1,PostingDate,RptAmtType,FundID,Details,ClientID,TransactionID,ReconNow " & _
            " from (SELECT  '+' as sign, Mid(tlbTransactionTypes.CONSTANT, 4, (Len(tlbTransactionTypes.CONSTANT)-3)) & tlbReceipt.SlNumber as INVNO,TransactionID as TransactionI,'0' as SplitID,'' as" & _
            " Description, '' as FundCode, tlbReceipt.SageAccountNumber, tlbReceipt.RDate, tlbReceipt.Amount as  ReceiptAmount, " & _
            " tlbReceipt.ExtRef, RIGHT(tlbTransactionTypes.CONSTANT, 2) AS TYPE, Tenants.Name as Name1, Property.PropertyID, tlbReceipt.UnitID, tlbReceipt.BankCode, " & _
            " Tenants.Name,tlbReceipt.PostingDate,tlbReceipt.RptAmtType,tlbReceipt.FundID,tlbReceipt.Details,tlbReceipt.ClientID,tlbReceipt.TransactionID,tlbReceipt.ReconNow " & _
            " FROM tlbReceipt, tlbTransactionTypes, Tenants," & _
            " LeaseDetails, Units, Property WHERE (tlbReceipt.Type = 3 Or tlbReceipt.Type = 4" & _
            " OR tlbReceipt.Type = 2 ) AND tlbReceipt.Type = tlbTransactionTypes.TYPE_ID  AND" & _
            " tlbReceipt.SageAccountNumber = Tenants.SageAccountNumber AND Tenants.SageAccountNumber  = LeaseDetails.SageAccountNumber AND LeaseDetails.UnitNumber =" & _
            " Units.UnitNumber " & strWh & " AND Units.PropertyID = Property.PropertyID AND tlbReceipt.Amount > tlbReceipt.OSAmount and tlbReceipt.SageAccountNumber = '" & szLesseeID & "'" & _
            " UNION ALL Select  '-' as sign,'' as INVNO,RptHeader as TransactionI,SplitID,tlbReceiptSplit.Description,F.FundCode as" & _
            " FundCode,tlbReceipt.SageAccountNumber, tlbReceipt.RDate,tlbReceiptSplit.Amount as ReceiptAmount, " & _
            " tlbReceipt.ExtRef, RIGHT(tlbTransactionTypes.CONSTANT, 2) AS TYPE, Tenants.Name as Name1, Property.PropertyID, tlbReceipt.UnitID, tlbReceipt.BankCode, " & _
            " Tenants.Name,tlbReceipt.PostingDate,tlbReceipt.RptAmtType,tlbReceipt.FundID,tlbReceipt.Details,tlbReceipt.ClientID,tlbReceipt.TransactionID,tlbReceipt.ReconNow " & _
            " from  tlbReceiptSplit,tlbReceipt, tlbTransactionTypes,FUND F, Tenants, LeaseDetails," & _
            " Units, Property WHERE (tlbReceipt.Type = 3 Or tlbReceipt.Type = 4 OR" & _
            " tlbReceipt.Type = 23 ) AND tlbReceipt.Type = tlbTransactionTypes.TYPE_ID AND F.FundID=tlbReceiptSplit.FundID AND" & _
            " tlbReceipt.SageAccountNumber = Tenants.SageAccountNumber AND Tenants.SageAccountNumber  = LeaseDetails.SageAccountNumber AND tlbReceiptSplit.RptHeader=tlbReceipt.TransactionID AND" & _
            " LeaseDetails.UnitNumber = Units.UnitNumber AND LeaseDetails.Status = True AND Units.PropertyID = Property.PropertyID AND tlbReceipt.Amount > tlbReceipt.OSAmount  and tlbReceipt.SageAccountNumber = '" & szLesseeID & "')"
            szSQL = szSQL & " order by TransactionI,SplitID "
''AND LeaseDetails.Status = True
'   szSQL = "SELECT R.TransactionID, R.SlNumber, R.Amount, R.RDate, R.UnitID, R.Details, " & _
'               "T.DESCRIPTION, R.ExtRef, " & _
'                  "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
'           "FROM tlbReceipt AS R, tlbTransactionTypes AS T " & _
'           "WHERE " & _
'               "R.SageAccountNumber = '" & szLesseeID & "' AND " & _
'               "(R.Type = 2 OR R.Type = 3 OR R.Type = 4) AND " & _
'               "R.Type = T.TYPE_ID AND R.Amount > R.OSAmount " & _
'           "ORDER BY R.TransactionID, R.Type;"
' If WS_Name = "PCM-DEV2" Then
'   szSQL = "SELECT R.TransactionID, R.SlNumber, R.Amount, R.RDate, R.UnitID, R.Details, " & _
'               "T.DESCRIPTION, R.ExtRef, " & _
'                  "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF " & _
'           "FROM tlbReceipt AS R, tlbTransactionTypes AS T " & _
'           "WHERE " & _
'               "(R.Type = 2 OR R.Type = 3 OR R.Type = 4) AND " & _
'               "R.Type = T.TYPE_ID AND R.Amount > R.OSAmount " & _
'           "ORDER BY R.TransactionID, R.Type;"
'End If
'2:  Sales Credit, 3:  Sales Receipt, 4:  Sales Receipt on Account
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   i = 1
   While Not adoRst.EOF
      flxCrPoA.TextMatrix(i, 2) = adoRst.Fields.Item("SIGN").Value
      If flxCrPoA.TextMatrix(i, 2) = "+" Then
            flxCrPoA.TextMatrix(i, 0) = adoRst.Fields.Item("TransactionI").Value 'colwidth for this is zero
            flxCrPoA.TextMatrix(i, 3) = adoRst.Fields.Item("INVNO").Value ' adoRst.Fields.Item("PF").Value & adoRst.Fields.Item("SlNumber").Value
            flxCrPoA.TextMatrix(i, 4) = IIf(IsNull(adoRst.Fields.Item("DESCRIPTION").Value), "", adoRst.Fields.Item("DESCRIPTION").Value)
            flxCrPoA.TextMatrix(i, 5) = IIf(IsNull(adoRst.Fields.Item("UnitID").Value), "", adoRst.Fields.Item("UnitID").Value)
            flxCrPoA.TextMatrix(i, 6) = adoRst.Fields.Item("RDate").Value
            flxCrPoA.TextMatrix(i, 7) = IIf(IsNull(adoRst.Fields.Item("ExtRef").Value), "", adoRst.Fields.Item("ExtRef").Value)
            flxCrPoA.TextMatrix(i, 8) = IIf(IsNull(adoRst.Fields.Item("Details").Value), "", adoRst.Fields.Item("Details").Value)
            flxCrPoA.TextMatrix(i, 9) = Format(adoRst.Fields.Item("ReceiptAmount").Value, "0.00")
            flxCrPoA.RowHeight(i) = 280
      Else
            flxCrPoA.TextMatrix(i, 7) = IIf(IsNull(adoRst!FundCode), "", adoRst!FundCode) ' null problem fixed by anol 20160705
            flxCrPoA.TextMatrix(i, 8) = IIf(IsNull(adoRst!description), "", adoRst!description)
            flxCrPoA.TextMatrix(i, 9) = Format(adoRst!receiptAmount, "0.00")
            flxCrPoA.RowHeight(i) = 0
      End If
      adoRst.MoveNext

      If Not adoRst.EOF Then
         i = i + 1
         flxCrPoA.AddItem ""
      End If
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub SquezeExpand()
    On Error GoTo Err
       Dim i As Integer, iCurRowHeight As Integer

  iCurRowHeight = 280
   

   If flxCrPoA.TextMatrix(flxCrPoA.row, 2) = "+" Then           'Expanding the grid 'flxCrPoA.col = 1 And
      flxCrPoA.TextMatrix(flxCrPoA.row, 2) = ">"
      iCurRowHeight = flxCrPoA.RowHeight(flxCrPoA.row)
      i = 1

      While flxCrPoA.TextMatrix(flxCrPoA.row + i, 2) = "-"
         flxCrPoA.RowHeight(flxCrPoA.row + i) = iCurRowHeight
         i = i + 1
         If (flxCrPoA.row + i) = flxCrPoA.Rows Then Exit Sub
      Wend
      Exit Sub
   End If

   If flxCrPoA.TextMatrix(flxCrPoA.row, 2) = ">" Then          'Squeezing the grid 'flxCrPoA.col = 1 And
      flxCrPoA.TextMatrix(flxCrPoA.row, 2) = "+"
      i = 1
      While flxCrPoA.TextMatrix(flxCrPoA.row + i, 2) = "-"
         flxCrPoA.RowHeight(flxCrPoA.row + i) = 0
         i = i + 1
         If (flxCrPoA.row + i) = flxCrPoA.Rows Then Exit Sub
      Wend
      Exit Sub
   End If
   Exit Sub
Err:
   'HighLightRowFlxGridA flxCrPoA, flxCrPoA.row
End Sub
'
'Private Sub UnAllocateRest(adoConn As ADODB.Connection)
'   Dim szSQL As String
'   Dim adoSplit1 As New ADODB.Recordset
'   Dim szaAllocTransID() As String, i As Integer, j As Integer, szaTemp() As String
'
'   szSQL = "SELECT * " & _
'           "FROM RptTransactions " & _
'           "WHERE FromTran = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & ";"
'   adoSplit1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoSplit1.EOF
'
'
'
'
'
'   Wend
'
''     Mark the visibility only the Receipt header
'   adoConn.Execute _
'           "UPDATE tlbReceipt AS R " & _
'           "SET R.OSAmount = R.Amount, R.ReceiptView = TRUE " & _
'           "WHERE R.TransactionID = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & ";"
'
''     Update the OSAmt = Amt of all Receipt splits
''     At this stage save the AllocTranID in an array @szaAllocTransID
'   szSQL = "SELECT * " & _
'           "FROM tlbReceiptSplit " & _
'           "WHERE RptHeader = " & flxCrPoA.TextMatrix(flxCrPoA.row, 0) & ";"
'   adoSplit1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
''   Debug.Print szSQL
'   ReDim szaAllocTransID(RecordCount(adoSplit1) - 1) As String
'   ReDim daSIRptID(RecordCount(adoSplit1) - 1, 1) As Double
'
'   While Not adoSplit1.EOF
'      With adoSplit1
'         .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
'         .Update
'
'         szaAllocTransID(i) = .Fields.Item("AllocTranID").Value & _
'                              "-#-" & CStr(.Fields.Item("Amount").Value)
'         i = i + 1
'         .MoveNext
'      End With
'   Wend
'   adoSplit1.Close
'
''     Update all SI splits in the receipt split table
'   szSQL = "SELECT * " & _
'           "FROM tlbReceiptSplit;"
'
'   adoSplit1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'   For j = 0 To i - 1
'      szaTemp = Split(szaAllocTransID(j), "-#-")
'
'      If Len(szaTemp(0)) > 10 Then
'         adoSplit1.Find ("TransactionID = '" & szaTemp(0) & "'"), , , 1
'      Else
'         adoSplit1.Find ("RptHeader = " & CLng(szaTemp(0)) & ""), , , 1
'      End If
'      adoSplit1.Fields.Item("OSAmount").Value = CDbl(adoSplit1.Fields.Item("OSAmount").Value) + _
'                                                CDbl(szaTemp(1))
'      adoSplit1.Update
'      daSIRptID(j, 0) = CDbl(adoSplit1.Fields.Item("RptHeader").Value)
'      daSIRptID(j, 1) = CDbl(szaTemp(1))
'   Next j
'   adoSplit1.Close
'
''  Update the header of all SIs in the Receipt table.
'   szSQL = "SELECT * " & _
'           "FROM tlbReceipt;"
'
'   adoSplit1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'   For j = 0 To i - 1
'      adoSplit1.Find "TransactionID = " & daSIRptID(j, 0) & "", , , 1
'
'      adoSplit1.Fields.Item("OSAmount").Value = CDbl(adoSplit1.Fields.Item("OSAmount").Value) + _
'                                                daSIRptID(j, 1)
'      adoSplit1.Update
'   Next j
'   adoSplit1.Close
'End Sub

Private Function UnAllocateReceipt(adoConn As ADODB.Connection, headerTransactionID As Long) As Boolean
   Dim adoSplit1 As New ADODB.Recordset

   Dim szSQL   As String
   Dim i       As Integer
   Dim j       As Integer

   Dim szaAllocTransID()   As String
   Dim daSIRptID()         As Double
   Dim szaTemp()           As String
   Dim dicPayHeader As New Dictionary
   Dim dicPayHeaderAmount As New Dictionary
   Dim dblPerInvPayAmount As Double
   Dim dblSIAmount As Double
   Dim strLesseeID As String
   Dim dblTotalPayAmount As Double
   On Error GoTo Err
        szSQL = "SELECT * " & _
         "FROM tlbReceipt " & _
         "WHERE transactionID = " & headerTransactionID & " "
        adoSplit1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not adoSplit1.EOF Then
            If adoSplit1("Type").Value = 2 Then
                MsgBox "Please check if this credit note has been included in a previous client statement as changing the allocation date may amend your previous client statement", vbInformation, "Warning"
            End If
        End If
        adoSplit1.Close
   szSQL = "SELECT * " & _
           "FROM tlbReceiptSplit " & _
           "WHERE RptHeader = " & headerTransactionID & " AND " & _
               "ISNULL(AllocTranID);"
   adoSplit1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoSplit1.EOF Then
     
      ShowMsgInTaskBar "Prestige cannot reverse this allocation. Please contact with PCM.", , "N"
      adoSplit1.Close
      Set adoSplit1 = Nothing
'      adoConn.Close
'      Set adoConn = Nothing

      Exit Function
   End If
   adoSplit1.Close

'  Mark the visibility only the Receipt header
   adoConn.Execute _
           "UPDATE tlbReceipt AS R " & _
           "SET R.OSAmount = R.Amount, R.ReceiptView = TRUE " & _
           "WHERE R.TransactionID = " & headerTransactionID & ";"

'  Update the OSAmt = Amt of all Receipt splits
'  At this stage save the AllocTranID in an array @szaAllocTransID
'@Example
'60 is receipt ID
'60 has one  split
'retreiving the allocation ID
'which is 15, Has one cell to store allocation ID
'saving allocation ID and amount in an array
   szSQL = "SELECT * " & _
           "FROM tlbReceiptSplit " & _
           "WHERE RptHeader = " & headerTransactionID & " Order by OSamount;"
   adoSplit1.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'Debug.Print szSQL
'Resolved by BOSL
'Below line added by anol 29 Mar 2015
'issue 549: Demand receipts not working. Note 1
   If Not IsNull(adoSplit1.Fields.Item("ClientStatementID").Value) Then
                MsgBox "This transaction cannot be unallocated because it is included in a Client statement", vbInformation, "Warning"
                Exit Function
   End If
   If adoSplit1.Fields.Item("ISMGTFeeS").Value = True Then
                MsgBox "This transaction cannot be unallocated because there is a management fee against it", vbInformation, "Warning"
                Exit Function
   End If
        
   If RecordCount(adoSplit1) > 0 Then
        ReDim szaAllocTransID(RecordCount(adoSplit1) - 1) As String
        ReDim daSIRptID(RecordCount(adoSplit1) - 1, 1) As Double
   End If
'here it is updating osamount of receipt-anol
   While Not adoSplit1.EOF
      With adoSplit1
         .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
         .Update

         szaAllocTransID(i) = .Fields.Item("AllocTranID").Value & _
                              "-#-" & CStr(.Fields.Item("Amount").Value)
         i = i + 1
         .MoveNext
      End With
   Wend
   adoSplit1.Close

'     Update all SI splits in the receipt split table
'here it is updating osamount of SI-anol
'Here relation is one SI :: two receipts for tran ID 60
' here vairiable i contains the split count of the receipt

'Left is SI | right is receipt
'Now you have to solve this problem
   
   For j = 0 To i - 1 'if a receipt has two lines this loop shall run twice  but now it has one line receipt tran ID 60
      szaTemp = Split(szaAllocTransID(j), "-#-")
    'szaTemp(0) contains the AllocTranID from tlbReceiptsplit table
     'szaTemp(1) contains the Amount from tlbReceiptsplit table
      If Len(szaTemp(0)) > 10 Then
          'modified by anol 2019-06-11
         adoSplit1.Open "SELECT * FROM tlbReceiptSplit where TransactionID = '" & szaTemp(0) & "' AND round(Amount,2)>=" & szaTemp(1) & " AND Amount<>OSamount;", adoConn, adOpenDynamic, adLockOptimistic
      Else
        'Modified by anol 2019-06-11
         adoSplit1.Open "SELECT * FROM tlbReceiptSplit where RptHeader = " & szaTemp(0) & " AND round(Amount,2)>=" & szaTemp(1) & " AND Amount<>OSamount;", adoConn, adOpenDynamic, adLockOptimistic 'herer szaTemp(0) =15 which have two split
      End If
      If RecordCount(adoSplit1) = 1 Then
            adoSplit1.Fields.Item("OSAmount").Value = CDbl(adoSplit1.Fields.Item("OSAmount").Value) + _
                                                      CDbl(szaTemp(1))
            dblSIAmount = dblSIAmount + CDbl(szaTemp(1))
            adoSplit1.Update
            If dicPayHeader.Exists(adoSplit1("RptHeader").Value) Then
                    'rsPaymentSplit("PayHeader").Value is the index of the collection
                    'MsgBox "This item Already exists ,cannot add again.But I am adding amount again "
                    'addAmount function add amount to an existing transation to the collection
                     dicPayHeader(adoSplit1("RptHeader").Value) = dicPayHeader(adoSplit1("RptHeader").Value) + CDbl(CDbl(szaTemp(1)))
            Else
                     dicPayHeader.Add adoSplit1("RptHeader").Value, CDbl(CDbl(szaTemp(1)))
                       
            End If
            
            
            If adoSplit1.Fields.Item("OSAmount").Value > adoSplit1.Fields.Item("Amount").Value Then
                adoSplit1.Close
                UnAllocateReceipt = False
                Exit Function
            End If
            daSIRptID(j, 0) = CDbl(adoSplit1.Fields.Item("RptHeader").Value)
            daSIRptID(j, 1) = CDbl(szaTemp(1))
      Else
            'No such record found as per allocation ID in receipt so exit this function and say failed
            adoSplit1.Close
            UnAllocateReceipt = False
            Exit Function
      End If
      adoSplit1.Close
   Next j

 Dim key
 For Each key In dicPayHeader.Keys
       adoSplit1.Open "SELECT * FROM tlbReceipt where TransactionID in (" & key & ") Order by OSAmount ;", adoConn, adOpenDynamic, adLockOptimistic
             dblPerInvPayAmount = dicPayHeader(key)
             If Not adoSplit1.EOF Then
                adoSplit1.Fields.Item("OSAmount").Value = CDbl(adoSplit1.Fields.Item("OSAmount").Value) + dblPerInvPayAmount
                adoSplit1.Fields.Item("ReceiptView").Value = True
                
                 dblTotalPayAmount = dblTotalPayAmount + dblPerInvPayAmount
                     
                 If strLesseeID = "" Then
                     strLesseeID = " '" & adoSplit1.Fields.Item("SageAccountNumber").Value & " '"
                 Else
                     strLesseeID = strLesseeID & ",'" & adoSplit1.Fields.Item("SageAccountNumber").Value & "'"
                 End If
                 If adoSplit1.Fields.Item("OSAmount").Value > adoSplit1.Fields.Item("Amount").Value Then
                             UnAllocateReceipt = False
                             Exit Function
                 End If
                 adoSplit1.Fields.Item("ReceiptView").Value = True
                 adoSplit1.Update
            End If
            adoSplit1.Close
            Set adoSplit1 = Nothing
 Next key
    If Round(dblTotalPayAmount, 2) = Round(dblSIAmount, 2) Then
    Else
        UnAllocateReceipt = False
        MsgBox "Amount mismatch at Receipt and SI while un allocation . UnAllocation failed."
        Exit Function
    End If
    adoConn.Execute "Update RptTransactions Set deleteflag=true WHERE RptTransactions.FromTran = " & headerTransactionID & ";"
    adoConn.Execute "Update RptTransactionsSplit Set deleteflag=true WHERE RptTransactionsSplit.FromTran = " & headerTransactionID & ";"
    Dim adoRst     As New ADODB.Recordset
    Dim szTran2Fix As String
  
    szSQL = "SELECT  R.TransactionID " & _
             "FROM tlbReceipt AS R, (" & _
                   "SELECT RptHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                   "From tlbReceiptSplit " & _
                   "Group by RptHeader " & _
                   ") AS Q " & _
             "WHERE R.TransactionID = Q.RptHeader AND R.Amount <> R.OSAmount AND " & _
                   "ROUND(R.Amount - R.OSAmount, 2) <> Q.T;"
    
    adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    
    While Not adoRst.EOF
       szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)
       adoRst.MoveNext
    Wend
    
    adoRst.Close
    Set adoRst = Nothing

   If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)
   If Len(szTran2Fix) > 0 Then
             MsgBox "The following transactions need fixing immediately: " & _
             Chr(13) & szTran2Fix & "." & _
             "Please contact PCM Consulting Support.", _
             vbInformation + vbOKOnly, " SI_Check "
             Exit Function
   Else
             UnAllocateReceipt = True
   End If
        Dim rsChecksum As New ADODB.Recordset
        rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions where deleteflag=false  group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID And R.sageaccountnumber in (" & strLesseeID & ") AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
        While Not rsChecksum.EOF
            szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
            rsChecksum.MoveNext
        Wend
        rsChecksum.Close
        Set rsChecksum = Nothing
        If Len(szTran2Fix) > 0 Then
             MsgBox "The following transactions need fixing immediately: " & _
             Chr(13) & szTran2Fix & "." & _
             "Please contact PCM Consulting Support.", _
             vbInformation + vbOKOnly, " SI_Check "
             Exit Function
   Else
             UnAllocateReceipt = True
   End If
   
   Exit Function
Err:
    UnAllocateReceipt = False
End Function
Private Function CountSelectedItem() As Long
  
   Dim rCount As Integer
   For rCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(rCount, 1) = "X" Then
            CountSelectedItem = CountSelectedItem + 1
        End If
   Next
   
End Function
Private Sub cmdRevereseAllocation_Click()
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rCount As Integer
   Dim bResult As Boolean
   Dim iCount As Integer
   Dim jCount As Integer
'   If flxCrPoA.TextMatrix(flxCrPoA.row, 0) = "" Then
'      ShowMsgInTaskBar "Please select a credit transaction first to reverse allocate.", , "N"
'      Exit Sub
'   End If


   If CountSelectedItem = 0 Then
      ShowMsgInTaskBar "Please select a credit transaction first to reverse allocate.", , "N"
      Exit Sub
   End If
   FocusControl cmdSPClose
   If MsgBox("Do you wish to reverse the allocation of selected credit transaction?", vbQuestion + vbYesNo, _
         "Reverse Allocation") = vbNo Then Exit Sub

'  Process of reversing starts -->
   
   For rCount = 1 To flxCrPoA.Rows - 1
        If flxCrPoA.TextMatrix(rCount, 1) = "X" Then
                adoConn.Open getConnectionString
                adoConn.BeginTrans
                ' flxCrPoA.TextMatrix(rCount, 0) is  headerTransactionID of tlbReceiptSplit
                bResult = UnAllocateReceipt(adoConn, flxCrPoA.TextMatrix(rCount, 0))
                If bResult = True Then
                     iCount = iCount + 1
                     adoConn.CommitTrans
                     flxCrPoA.RowHeight(rCount) = 0
                Else
                    adoConn.RollbackTrans
                    flxCrPoA.RowHeight(rCount) = 240
                    jCount = jCount + 1
                End If
               
                flxCrPoA.Refresh
                adoConn.Close
                Set adoConn = Nothing
        End If
    Next
'   configflxcrPoa flxCrPoA, 9, 7, 20, "|<No|<Type|<UnitID||<Date|<Ref|<Details|>Amt"
   Call ConfigFlxCrPoA
   Call ConfigFlxSPayment
   'ConfigFlxSPayment flxSPayment, 11, 9, 10
   adoConn.Open getConnectionString
   LoadFlxCrPoA adoConn
   flxCrPoA.row = 0

   frmDemands3.ConfigFlxSPayment
   frmDemands3.ConfigFlxCrPoA
   frmDemands3.LoadFlxSPayment adoConn
   frmDemands3.LoadFlxCrPoA adoConn

   adoConn.Close
   Set adoConn = Nothing
   chkCredits.Value = False
   If jCount = 0 Then
        ShowMsgInTaskBar iCount & " Transaction has been reveresed successfully."
   Else
        ShowMsgInTaskBar iCount & " Transaction has been reveresed successfully. Failed : " & jCount
   End If
'   MsgBox "Transaction has been reveresed successfully", vbInformation + vbOKOnly, "Reverese Allocation"
End Sub

'Private Sub ConfigFlxSPayment(ctrFlxGrid As MSHFlexGrid, iCols As Integer, iLabels As Integer, iLblFstIdx As Integer, Optional szHeader As String)
'   Dim iCol As Integer
'
'   ctrFlxGrid.Clear
'   ctrFlxGrid.Cols = iCols
'   ctrFlxGrid.Rows = 2
'   ctrFlxGrid.RowHeight(0) = 0
'
'   If szHeader <> "" Then _
'      ctrFlxGrid.FormatString = szHeader$
'
'   ctrFlxGrid.ColWidth(0) = 0                                           'ID of the transactions
'   ctrFlxGrid.ColWidth(1) = Label19(iLblFstIdx).Left - ctrFlxGrid.Left  'Sign -> X, +, -
'
'   For iCol = 2 To ctrFlxGrid.Cols - 3
'      ctrFlxGrid.ColWidth(iCol) = Label19(iCol + iLblFstIdx - 1).Left - Label19(iCol - 1 + iLblFstIdx - 1).Left
'   Next iCol
'   ctrFlxGrid.ColWidth(iCol) = ctrFlxGrid.Width + ctrFlxGrid.Left - Label19(iCol - 2 + iLblFstIdx).Left - 340
'End Sub
Private Sub ConfigFlxSPayment()
    Dim iCol As Integer
    
    flxSPayment.Clear
    flxSPayment.Cols = 12
    flxSPayment.Rows = 2
    flxSPayment.RowHeight(0) = 0
    flxSPayment.FormatString = "|<No|<Type|<UnitID||<Date|<Ref|Details|>Amt"
    flxSPayment.ColWidth(0) = 0                                           'ID of the transactions
    flxSPayment.ColWidth(1) = 250 'Sign -> X
    'flxSPayment.ColWidth(2) = 250   'Sign ->, +, -
    flxSPayment.ColWidth(2) = Label19(11).Left - Label19(10).Left
    flxSPayment.ColWidth(3) = Label19(12).Left - Label19(11).Left
    flxSPayment.ColWidth(4) = Label19(13).Left - Label19(12).Left
    flxSPayment.ColWidth(5) = Label19(14).Left - Label19(13).Left
    flxSPayment.ColWidth(6) = Label19(15).Left - Label19(14).Left
    flxSPayment.ColAlignment(6) = vbLeftJustify
    flxSPayment.ColWidth(7) = 3500 ' Label19(16).Left - Label19(15).Left
    flxSPayment.ColAlignment(7) = vbLeftJustify
    flxSPayment.ColWidth(8) = 1200 'Label19(27).Left - Label19(26).Left
     flxSPayment.ColWidth(9) = 1200
    flxSPayment.ColWidth(10) = 1200
    flxSPayment.ColWidth(11) = 1200
'    FlxSPayment.ColWidth(2) = 250
'    FlxSPayment.ColWidth(iCol) = 3500
'    FlxSPayment.ColWidth(iCol + 1) = 1300
End Sub
Private Sub ConfigFlxCrPoA()
    Dim iCol As Integer
    
    flxCrPoA.Clear
    flxCrPoA.Cols = 11
    flxCrPoA.Rows = 2
    flxCrPoA.RowHeight(0) = 0
    flxCrPoA.FormatString = "|<No|<Type|<UnitID||<Date|<Ref|Details|>Amt"
    flxCrPoA.ColWidth(0) = 0                                           'ID of the transactions
    flxCrPoA.ColWidth(1) = 250 'Sign -> X
    flxCrPoA.ColWidth(2) = 250   'Sign ->, +, -
    flxCrPoA.ColWidth(3) = Label19(21).Left - Label19(20).Left
    flxCrPoA.ColWidth(4) = Label19(22).Left - Label19(21).Left
    flxCrPoA.ColWidth(5) = Label19(23).Left - Label19(22).Left
    flxCrPoA.ColWidth(6) = Label19(24).Left - Label19(23).Left
    flxCrPoA.ColWidth(7) = Label19(25).Left - Label19(24).Left
    flxCrPoA.ColAlignment(7) = vbLeftJustify
    flxCrPoA.ColWidth(8) = Label19(26).Left - Label19(25).Left
    flxCrPoA.ColAlignment(8) = vbLeftJustify
    flxCrPoA.ColWidth(9) = 1500 'Label19(27).Left - Label19(26).Left
    flxCrPoA.ColWidth(10) = 0
    flxCrPoA.ColWidth(11) = 0
'    flxCrPoA.ColWidth(2) = 250
'    flxCrPoA.ColWidth(iCol) = 3500
'    flxCrPoA.ColWidth(iCol + 1) = 1300
End Sub
'Private Sub ConfigFlxCrPoA(ctrFlxGrid As MSHFlexGrid, iCols As Integer, iLabels As Integer, iLblFstIdx As Integer, Optional szHeader As String)
'   Dim iCol As Integer
'
'   ctrFlxGrid.Clear
'   ctrFlxGrid.Cols = iCols
'   ctrFlxGrid.Rows = 2
'   ctrFlxGrid.RowHeight(0) = 0
'
'   If szHeader <> "" Then _
'      ctrFlxGrid.FormatString = szHeader$
'
'   ctrFlxGrid.ColWidth(0) = 0                                           'ID of the transactions
'   ctrFlxGrid.ColWidth(1) = 250 'Label19(iLblFstIdx).Left - ctrFlxGrid.Left  'Sign -> X, +, -
'   ctrFlxGrid.ColWidth(2) = 250 'Label19(iLblFstIdx).Left - ctrFlxGrid.Left  'Sign -> X, +, -
'
'   For iCol = 2 To ctrFlxGrid.Cols - 3
'      ctrFlxGrid.ColWidth(iCol + 1) = Label19(iCol + iLblFstIdx - 1).Left - Label19(iCol - 1 + iLblFstIdx - 1).Left
'   Next iCol
'   ctrFlxGrid.ColWidth(iCol) = 3500 'ctrFlxGrid.Width + ctrFlxGrid.Left - Label19(iCol - 2 + iLblFstIdx).Left + 500
'   ctrFlxGrid.ColWidth(iCol + 1) = 1300 'ctrFlxGrid.Width + ctrFlxGrid.Left - Label19(iCol - 2 + iLblFstIdx).Left + 500
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   UnLoadForm Me
   Call WheelUnHook(Me.hWnd)
   frmDemands3.Show
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
Private Sub txtPayDt_Change()
   TextBoxChangeDate txtPayDt
End Sub
Private Sub txtPayDt_GotFocus()
   'iCurRow = StarFound

   If flxSPayment.col = 8 Then
      If flxSPayment.TextMatrix(flxSPayment.row, 8) = "" Then
         txtPayDt.text = Format(Now, "dd/mm/yyyy")
      Else
         txtPayDt.text = Format(flxSPayment.TextMatrix(flxSPayment.row, 8), "dd/mm/yyyy")
      End If
   End If
   SelTxtInCtrl txtPayDt
End Sub
Private Sub txtPayDt_LostFocus()
        If txtPayDt.text <> "" Then TextBoxFormatDate txtPayDt
        If txtPayDt.text = "" Then
            flxSPayment.TextMatrix(iCurRow, 8) = Format(Date, "dd/MM/yyyy")
        Else
            flxSPayment.TextMatrix(iCurRow, 8) = txtPayDt.text
        End If
        txtPayDt.Visible = False
'        If flxSPayment.TextMatrix(iCurRow, 8) <> "" Then
'            flxSPayment.TextMatrix(iCurRow, 8) = flxSPayment.TextMatrix(flxSPayment.row, 8)
'        End If
End Sub
Private Sub flxSPayment_DblClick()
    If flxSPayment.col = 8 Then
      iCurRow = flxSPayment.row
      'flxSPayment.col = iFlxSPayCol
      txtPayDt.Top = flxSPayment.CellTop + flxSPayment.Top
     ' iTop = txtSPayment.Top
      txtPayDt.Left = flxSPayment.CellLeft + flxSPayment.Left
      'iLeft = flxSPayment.CellLeft + flxSPayment.Left
      txtPayDt.Width = 1200 'flxSPayment.ColWidth(iFlxSPayCol)
      txtPayDt.Height = flxSPayment.RowHeight(flxSPayment.row) - 15
      txtPayDt.text = flxSPayment.TextMatrix(flxSPayment.row, 8)
      txtPayDt.Visible = True
      'txtPayDt.ScrollBars = flexScrollBarNone
      txtPayDt.SetFocus
   End If
End Sub
Private Sub txtPayDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      flxSPayment.TextMatrix(iCurRow, 8) = txtPayDt.text
      If IsDate(txtPayDt.text) = False Then
            ShowMsgInTaskBar "Date format is not correct!"
            Exit Sub 'added by anol 20160911
      End If
      txtPayDt.Visible = False
      flxSPayment.ScrollBars = flexScrollBarVertical
   End If
   TextBoxKeyPrsDate txtPayDt, KeyAscii
End Sub


