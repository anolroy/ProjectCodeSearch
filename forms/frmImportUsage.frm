VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmImportUsage 
   BackColor       =   &H00FFEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Usage"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10620
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2775
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   4680
      Width           =   10695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeases 
         Height          =   2280
         Left            =   75
         TabIndex        =   17
         Top             =   435
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4022
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
         Index           =   9
         Left            =   9360
         TabIndex        =   27
         Top             =   210
         Width           =   570
         VariousPropertyBits=   276824083
         Caption         =   "Amount"
         Size            =   "1005;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   8
         Left            =   8520
         TabIndex        =   26
         Top             =   210
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
         Height          =   420
         Index           =   0
         Left            =   100
         TabIndex        =   25
         Top             =   15
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
         Index           =   1
         Left            =   660
         TabIndex        =   24
         Top             =   210
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
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   23
         Top             =   210
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
         Height          =   390
         Index           =   4
         Left            =   4755
         TabIndex        =   22
         Top             =   15
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
         Height          =   270
         Index           =   5
         Left            =   5760
         TabIndex        =   21
         Top             =   135
         Width           =   615
         VariousPropertyBits=   276824083
         Caption         =   "Reading"
         Size            =   "1085;476"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   6
         Left            =   6585
         TabIndex        =   20
         Top             =   15
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
         Height          =   225
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   180
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
         Height          =   195
         Index           =   7
         Left            =   7515
         TabIndex        =   18
         Top             =   210
         Width           =   570
         VariousPropertyBits=   276824083
         Caption         =   "To Date"
         Size            =   "1005;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   420
         Left            =   75
         Top             =   0
         Width           =   10455
      End
   End
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
         Caption         =   "Import Individual Utility Usage"
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
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   10695
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8655
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   7260
         TabIndex        =   4
         Top             =   360
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
      Height          =   2775
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   10695
      Begin VB.TextBox txtImportFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   2160
         Width           =   7455
      End
      Begin VB.CommandButton cmdClinetAddAtch 
         Height          =   315
         Left            =   1440
         Picture         =   "frmImportUsage.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select File"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   750
      End
      Begin MSForms.ComboBox cboType 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1485
         Width           =   8415
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "14843;556"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1485
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   2
         Left            =   360
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
         Left            =   360
         TabIndex        =   11
         Top             =   795
         Width           =   615
      End
      Begin MSForms.ComboBox cboDmdPropertyList 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   795
         Width           =   8415
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "14843;556"
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
         Left            =   1440
         TabIndex        =   0
         Top             =   120
         Width           =   8415
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "14843;556"
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
   End
End
Attribute VB_Name = "frmImportUsage"
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
   If MsgBox("Are you sure you want to cancel the routine?", vbQuestion + vbYesNo, "Global Lease Update") = vbNo Then Exit Sub
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

Private Sub cmdClinetAddAtch_Click()
   BrowseForFile
End Sub

Public Sub BrowseForFile()
   Dim ofn As OPENFILENAME
   Dim lHwnd As Long
   Const HKEY_LOCAL_MACHINE As Long = &H80000002
   Dim szOldFile_PathName As String
   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
   Dim fso As Object, adoConn As New ADODB.Connection

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

   txtImportFile.text = ofn.lpstrFile

   Exit Sub
FileError:
   ShowMsgInTaskBar "File not does not exists"
End Sub

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

Private Sub cmdNext_Click()
   If txtImportFile.text = "" Then
      ShowMsgInTaskBar "Please select the input file to import", "Y", "N"
      cmdClinetAddAtch.SetFocus
      Exit Sub
   End If

   Dim oXL As New Excel.Application
   Dim oWB As Workbook
   Dim oWS As Worksheet

   Set oWB = oXL.Workbooks.Open(txtImportFile.text)
   Set oWS = oWB.Worksheets(1) 'Specify your worksheet name

   Dim iRow As Integer
   Dim iChr As Integer

   If UCase(oWS.Range("A1").Value) <> "TYPE" Then
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
      For iChr = 65 To 74
         If iChr - 64 = 10 Then
            flxLeases.TextMatrix(iRow - 1, iChr - 64) = Format(oWS.Range(Chr$(iChr) & iRow).Value, "0.00")
         Else
            flxLeases.TextMatrix(iRow - 1, iChr - 64) = oWS.Range(Chr$(iChr) & iRow).Value
         End If
      Next iChr
      iRow = iRow + 1
      If oWS.Range("A" & CStr(iRow)).Value <> vbEmpty Then flxLeases.AddItem ""
   Wend

   oWB.Close
   oXL.Quit

   Set oWS = Nothing
   Set oWB = Nothing
   Set oXL = Nothing

   Frame1(3).Top = Frame1(0).Top
   Frame1(3).Left = Frame1(0).Left
   Frame1(3).ZOrder 0
   cmdNext.Enabled = False
   cmdFinish.Enabled = True
End Sub

Private Sub Form_Load()
   Me.Width = 10710
   Me.Height = 4950
   Me.Top = frmMMain.Height / 2 - Me.Height
   Me.Left = 0 'frmMMain.Width / 2 - Me.Width + 500

   Me.BackColor = MODULEBACKCOLOR
   Frame1(0).BackColor = Me.BackColor
   Frame1(8).BackColor = Me.BackColor

   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   PrepareList adoConn

   adoConn.Close
   Set adoConn = Nothing

   bPreviewShowed = False
   bDerivedFromSelAll = False
   bBackGotFocused = False
   ConfigureFlxLeases
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

   If Not adoRST.EOF Then
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
   End If
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & cboDmdClientList.Column(0) & "' " & _
           "ORDER BY PropertyID;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
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
   End If
   adoRST.Close

'*************************************** DEMAND TYPES ******************************************
   szSQL = "SELECT ID, TYPE " & _
           "FROM DEMANDTYPES " & _
           "WHERE PropertyID = '" & cboDmdPropertyList.Value & "';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
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
   End If
   adoRST.Close
   Set adoRST = Nothing
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
