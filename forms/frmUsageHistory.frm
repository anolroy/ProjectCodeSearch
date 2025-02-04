VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUsageHistory 
   BackColor       =   &H00FFEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Usage"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsageHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12045
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -20
      ScaleHeight     =   825
      ScaleWidth      =   10665
      TabIndex        =   9
      Top             =   0
      Width           =   10695
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Utility Usage History"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   6360
      Width           =   10695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   375
         Left            =   8880
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C000C0&
         Index           =   1
         X1              =   0
         X2              =   10680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   10680
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5655
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   10695
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Height          =   315
         Left            =   9360
         TabIndex        =   5
         Top             =   560
         Width           =   1215
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Text            =   "01/01/2010"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeases 
         Height          =   4080
         Left            =   120
         TabIndex        =   19
         Top             =   1395
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7197
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   7
         Left            =   7560
         TabIndex        =   14
         Top             =   1170
         Width           =   570
         VariousPropertyBits=   276824083
         Caption         =   "To Date"
         Size            =   "1005;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Index           =   3
         Left            =   2565
         TabIndex        =   15
         Top             =   1140
         Width           =   1065
         VariousPropertyBits=   276824083
         Caption         =   "Description"
         Size            =   "1879;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   6
         Left            =   6630
         TabIndex        =   16
         Top             =   975
         Width           =   375
         VariousPropertyBits=   276824083
         Caption         =   "From Date"
         Size            =   "661;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   150
         Index           =   5
         Left            =   5805
         TabIndex        =   26
         Top             =   1215
         Width           =   615
         VariousPropertyBits=   8388627
         Caption         =   "Reading"
         Size            =   "1085;265"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   4
         Left            =   4800
         TabIndex        =   25
         Top             =   975
         Width           =   360
         VariousPropertyBits=   276824083
         Caption         =   "Bill Date"
         Size            =   "635;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   2
         Left            =   1725
         TabIndex        =   24
         Top             =   1170
         Width           =   315
         VariousPropertyBits=   276824083
         Caption         =   "Unit"
         Size            =   "556;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   1
         Left            =   705
         TabIndex        =   23
         Top             =   1170
         Width           =   510
         VariousPropertyBits=   276824083
         Caption         =   "Lessee"
         Size            =   "900;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   420
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   975
         Width           =   675
         VariousPropertyBits=   276824083
         Caption         =   "Utility Type"
         Size            =   "1191;741"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   8
         Left            =   8565
         TabIndex        =   21
         Top             =   1170
         Width           =   450
         VariousPropertyBits=   276824083
         Caption         =   "Usage"
         Size            =   "794;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   9
         Left            =   9405
         TabIndex        =   20
         Top             =   1170
         Width           =   570
         VariousPropertyBits=   276824083
         Caption         =   "Amount"
         Size            =   "1005;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         Height          =   195
         Index           =   1
         Left            =   7080
         TabIndex        =   18
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin MSForms.ComboBox cboType 
         Height          =   315
         Left            =   5880
         TabIndex        =   2
         Top             =   165
         Width           =   3015
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5318;556"
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
         Object.Width           =   "881"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Type"
         Height          =   195
         Index           =   33
         Left            =   4800
         TabIndex        =   13
         Top             =   165
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   555
         Width           =   615
      End
      Begin MSForms.ComboBox cboDmdPropertyList 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   555
         Width           =   3615
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6376;556"
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
         Left            =   960
         TabIndex        =   0
         Top             =   120
         Width           =   3615
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6376;556"
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
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   420
         Left            =   120
         Top             =   960
         Width           =   10455
      End
   End
End
Attribute VB_Name = "frmUsageHistory"
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
           As Long) As Long
Const SW_SHOW = 5

Private Const LD_ID = 9
Private Const UU_ID = 10

Private Const MAX_NUMBER_FRAME_INDEX = 7
Private bPreviewShowed As Boolean
Private szDemandCharges As String
Private iSelectedFlxRec As Integer
Private bDerivedFromSelAll As Boolean

Private bBackGotFocused As Boolean
Private bAllInputted As Boolean
Private iLeft As Integer, iTop As Integer, iCurRow As Integer
Private aListPropertyID() As String
Private aListDT_ID() As String

Private Sub cmdBack_GotFocus()
   bBackGotFocused = True
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Function MoveRightPosition() As Boolean
   flxLeases.col = flxLeases.col + 1
   If flxLeases.col > 7 Then
      MoveRightPosition = False
      Exit Function
   End If
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveRightPosition = True
End Function

Private Function MoveLeftPosition() As Boolean
   flxLeases.col = flxLeases.col - 1
   If flxLeases.col < 5 Then
      MoveLeftPosition = False
      Exit Function
   End If
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveLeftPosition = True
End Function

Private Function MoveUpPosition() As Boolean
   Dim iRow As Integer

   For iRow = iCurRow - 1 To 1 Step -1
      If flxLeases.TextMatrix(iRow, 0) = "X" Then
         flxLeases.row = iRow
         iCurRow = iRow
         Exit For
      End If
   Next iRow
   If iRow = 0 Then
      iCurRow = 0
      MoveUpPosition = False
      Exit Function
   End If

   If flxLeases.col < 5 Then flxLeases.col = 5
   If flxLeases.col > 7 Then flxLeases.col = 7
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveUpPosition = True
End Function

Private Function MoveDownPosition() As Boolean
   Dim iRow As Integer

   For iRow = iCurRow + 1 To flxLeases.Rows - 1
      If flxLeases.TextMatrix(iRow, 0) = "X" Then
         flxLeases.row = iRow
         iCurRow = iRow
         Exit For
      End If
   Next iRow
   If iRow = flxLeases.Rows Then
      iCurRow = 0
      MoveDownPosition = False
      Exit Function
   End If

   If flxLeases.col < 5 Then flxLeases.col = 5
   If flxLeases.col > 7 Then flxLeases.col = 7
   iLeft = flxLeases.CellLeft + flxLeases.Left
   iTop = flxLeases.CellTop + flxLeases.Top
   MoveDownPosition = True
End Function

Private Sub cmdFinish_Click()
   If MsgBox("Do you wish to import the utilities?", vbQuestion + vbYesNo, "Import Utility Usage") = vbNo Then Exit Sub

   Dim szSQL   As String
   Dim iRow    As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST  As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT * " & _
           "FROM LUtilityUsage;"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxLeases.Rows - 1
      adoRST.AddNew
      adoRST.Fields.Item(0).Value = UniqueID()
      adoRST.Fields.Item(1).Value = GetLeaseID(adoConn, flxLeases.TextMatrix(iRow, 2), flxLeases.TextMatrix(iRow, 3))

      If UCase(Left(flxLeases.TextMatrix(iRow, 1), 1)) = "E" Then
         adoRST.Fields.Item(2).Value = "ELC"
      End If
      If UCase(Left(flxLeases.TextMatrix(iRow, 1), 1)) = "G" Then
         adoRST.Fields.Item(2).Value = "GAS"
      End If
      If UCase(Left(flxLeases.TextMatrix(iRow, 1), 1)) = "W" Then
         adoRST.Fields.Item(2).Value = "WAT"
      End If

      adoRST.Fields.Item(3).Value = flxLeases.TextMatrix(iRow, 4)
      adoRST.Fields.Item(4).Value = flxLeases.TextMatrix(iRow, 6)
      adoRST.Fields.Item(5).Value = flxLeases.TextMatrix(iRow, 5)
      adoRST.Fields.Item(7).Value = flxLeases.TextMatrix(iRow, 9)
      adoRST.Fields.Item(8).Value = flxLeases.TextMatrix(iRow, 10)
      adoRST.Fields.Item(13).Value = cboType.Value
      adoRST.Fields.Item(15).Value = flxLeases.TextMatrix(iRow, 7)
      adoRST.Fields.Item(16).Value = flxLeases.TextMatrix(iRow, 8)
      adoRST.Update
   Next iRow

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing

   Form_Unload 0
   ShowMsgInTaskBar "Utility usage have been update successfully", "Y", "P"
End Sub

Private Sub SelectRowFlxLease(iRow As Integer, bSel As Boolean)
   Dim c As Integer

   flxLeases.TextMatrix(iRow, 0) = IIf(bSel, "X", "")
   flxLeases.row = iRow
   For c = 0 To flxLeases.Cols - 1
      flxLeases.col = c
      flxLeases.CellBackColor = IIf(bSel, &HE0E0E0, vbWhite)
   Next c
   Label1(11).Visible = IIf(iSelectedFlxRec > 0, True, False)
   Label1(11).Caption = iSelectedFlxRec
End Sub

Private Sub LoadFlxLeases(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim iRow    As Integer
   Dim adoRST  As New ADODB.Recordset

   szSQL = "SELECT U.UtilityType, U.Description, U.CurrentReading, U.ReadingDate, " & _
               "U.Usage, U.Amount, U.Balance, U.DateFrom, U.DateTo, " & _
               "L.SageAccountNumber, L.UnitNumber " & _
           "FROM (((LUtilityUsage AS U INNER JOIN LeaseDetails AS L ON U.LeaseID = L.LeaseID) " & _
               "INNER JOIN Units AS N ON L.UnitNumber = N.UnitNumber) " & _
               "INNER JOIN Property AS P ON N.PropertyID = P.PropertyID) " & _
           "WHERE U.DemandType = " & cboType.Value & " AND " & _
               "U.DateFrom >= #" & txtDateFrom.text & "# AND " & _
               "U.DateTo <= #" & txtDateTo.text & "# AND " & _
               "N.PropertyID = '" & cboDmdPropertyList.Value & "' AND " & _
               "P.ClientID = '" & cboDmdClientList.Value & "';"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   iRow = 1
   While Not adoRST.EOF
      flxLeases.TextMatrix(iRow, 1) = "ELC"
      flxLeases.TextMatrix(iRow, 2) = adoRST.Fields.Item("SageAccountNumber").Value
      flxLeases.TextMatrix(iRow, 3) = adoRST.Fields.Item("UnitNumber").Value
      flxLeases.TextMatrix(iRow, 4) = IIf(IsNull(adoRST.Fields.Item("Description").Value), "", adoRST.Fields.Item("Description").Value)
      flxLeases.TextMatrix(iRow, 5) = adoRST.Fields.Item("ReadingDate").Value
      flxLeases.TextMatrix(iRow, 6) = IIf(IsNull(adoRST.Fields.Item("CurrentReading").Value), _
                                             "", adoRST.Fields.Item("CurrentReading").Value)
      flxLeases.TextMatrix(iRow, 7) = IIf(IsNull(adoRST.Fields.Item("DateFrom").Value), "", adoRST.Fields.Item("DateFrom").Value)
      flxLeases.TextMatrix(iRow, 8) = IIf(IsNull(adoRST.Fields.Item("DateTo").Value), "", adoRST.Fields.Item("DateTo").Value)
      flxLeases.TextMatrix(iRow, 9) = adoRST.Fields.Item("Usage").Value
      flxLeases.TextMatrix(iRow, 10) = adoRST.Fields.Item("Amount").Value
      
      adoRST.MoveNext
      If Not adoRST.EOF Then flxLeases.AddItem ""
      iRow = iRow + 1
   Wend

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdOK_Click()
   If txtDateFrom.text = "" Then
      txtDateFrom.SetFocus
      ShowMsgInTaskBar "Please enter the From Date", "Y", "N"
      Exit Sub
   End If
   If txtDateTo.text = "" Then
      txtDateTo.SetFocus
      ShowMsgInTaskBar "Please enter the To Date", "Y", "N"
      Exit Sub
   End If

   ConfigureFlxLeases

   Dim adoConn As New ADODB.Connection
'   connect to database
   adoConn.Open getConnectionString
   
   LoadFlxLeases adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_Load()
   Me.Width = 10710
   Me.Height = 7545
   frmMMain.Arrange vbCascade
   Me.ZOrder 0

   Me.BackColor = MODULEBACKCOLOR
   Frame1(0).BackColor = Me.BackColor
   Frame1(8).BackColor = Me.BackColor
   txtDateTo.text = Format(Now, "dd/mm/yyyy")
   ConfigureFlxLeases

   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   PrepareList adoConn
   LoadFlxLeases adoConn

   adoConn.Close
   Set adoConn = Nothing

   bPreviewShowed = False
   bDerivedFromSelAll = False
   bBackGotFocused = False
   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount - 1
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i

   cboDmdClientList.Column() = Data()
   cboDmdClientList.ListIndex = 0
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & cboDmdClientList.Column(0) & "' " & _
           "ORDER BY PropertyID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount - 1
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   cboDmdPropertyList.Column() = Data()
   cboDmdPropertyList.ListIndex = 0
   adoRST.Close

'*************************************** DEMAND TYPES ******************************************
   szSQL = "SELECT ID, TYPE " & _
           "FROM DEMANDTYPES " & _
           "WHERE PropertyID = '" & cboDmdPropertyList.Value & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRST.RecordCount - 1
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
      Next j
      adoRST.MoveNext
      If adoRST.EOF Then Exit For
   Next i
   cboType.Column() = Data()
   cboType.ListIndex = 0
   adoRST.Close

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigureFlxLeases()
   Dim szHeader As String, iCol As Integer

   flxLeases.Clear
   flxLeases.Cols = 11
   flxLeases.Rows = 2
   flxLeases.RowHeight(0) = 0
   szHeader$ = "|<UtilityType|<Lessee<|Unit|<Decription|<BillDate|>CurrentReading|<FromDate|<ToDate|>Usage|>Amount"
   flxLeases.FormatString = szHeader$

   flxLeases.ColWidth(0) = 0
   For iCol = 1 To flxLeases.Cols - 2
      flxLeases.ColWidth(iCol) = Label1(iCol).Left - Label1(iCol - 1).Left
   Next iCol
   flxLeases.ColWidth(iCol) = flxLeases.Left + flxLeases.Width - Label1(iCol - 1).Left - 250

   flxLeases.row = 0
   flxLeases.col = 0

   iSelectedFlxRec = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
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
