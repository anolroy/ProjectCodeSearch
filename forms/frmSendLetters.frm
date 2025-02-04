VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendLetters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Letters"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendLetters.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   14625
   Begin VB.Frame Frame17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Attachments"
      ForeColor       =   &H00000000&
      Height          =   1680
      Index           =   0
      Left            =   675
      MousePointer    =   1  'Arrow
      TabIndex        =   27
      Top             =   1890
      Visible         =   0   'False
      Width           =   7260
      Begin VB.CommandButton cmdDeleteFile 
         Caption         =   "&Remove Attachment"
         Height          =   360
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   315
         Width           =   1830
      End
      Begin VB.CommandButton cmdClinetAddAtch 
         Caption         =   "&Add Attachment"
         Height          =   360
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   315
         Width           =   1605
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Open Attachment"
         Height          =   360
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   315
         Width           =   1695
      End
      Begin VB.CommandButton cmdAttachOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Send Email"
         Height          =   360
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   315
         Width           =   1155
      End
      Begin VB.CommandButton cmdCloseAttcah 
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
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   135
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAttachment 
         Height          =   870
         Left            =   90
         TabIndex        =   33
         Top             =   720
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   1535
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
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
   End
   Begin VB.Frame fraListTemplates 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11415
      Begin VB.CheckBox chkAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkLogo 
         Caption         =   "Logo,"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Next >>"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8640
         TabIndex        =   0
         Top             =   3800
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTemplates 
         Height          =   3405
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   6006
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   10
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
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
         _Band(0).Cols   =   10
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print company "
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   9720
         TabIndex        =   5
         Top             =   120
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified by"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         Top             =   120
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Template Name"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fraPropertyUnit 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   11415
      Begin MSComctlLib.ListView lvwSelectedLessee 
         Height          =   3345
         Left            =   7560
         TabIndex        =   26
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "UnitID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LesseeID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LesseeName"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.ListBox lstDisplay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3345
         ItemData        =   "frmSendLetters.frx":17D2A
         Left            =   7560
         List            =   "frmSendLetters.frx":17D2C
         TabIndex        =   11
         Top             =   4080
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.CommandButton cmdPrintWord 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9120
         Picture         =   "frmSendLetters.frx":17D2E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3800
         Width           =   495
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "Email Letters"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7560
         TabIndex        =   23
         Top             =   3800
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddUnitAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   18
         Top             =   960
         Width           =   415
      End
      Begin VB.CommandButton cmdRemUnitAll 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   17
         Top             =   2595
         Width           =   415
      End
      Begin VB.CommandButton cmdRemUnit 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   16
         Top             =   3240
         Width           =   415
      End
      Begin VB.CommandButton cmdAddUnit 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7080
         TabIndex        =   15
         Top             =   360
         Width           =   415
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Print Letters"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9840
         TabIndex        =   14
         Top             =   3800
         Width           =   1335
      End
      Begin VB.ListBox lstProperty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   3345
         ItemData        =   "frmSendLetters.frx":17EE8
         Left            =   120
         List            =   "frmSendLetters.frx":17EEA
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdback 
         Caption         =   "<< Back"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   3800
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvwUnit 
         Height          =   3345
         Left            =   3600
         TabIndex        =   25
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "UnitID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LesseeID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LesseeName"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.ListBox lstUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Columns         =   2
         Height          =   3345
         ItemData        =   "frmSendLetters.frx":17EEC
         Left            =   3600
         List            =   "frmSendLetters.frx":17EEE
         TabIndex        =   12
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Units/Lessees"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   10
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmSendLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bChoice   As Boolean
Private bSaved    As Boolean
Private bValidate As Boolean
Public szLetter   As String
Private szP_U()   As String
Private lTemplateID As Long
Dim attachmentframePlayed As Boolean
'Dim DataAttach() As String
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

Private Sub cmdAddUnit_Click()
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer
   Dim szULeft()  As String
   Dim szURight() As String

   If lstUnit.SelCount < 1 Then Exit Sub
   
   szULeft = Split(lvwUnit.SelectedItem, " \ ")

   For i = 0 To lstDisplay.ListCount
      If lstDisplay.List(i) = "" Then Exit For
      szURight = Split(lstDisplay.List(i), " \ ")
      If szULeft(0) = szURight(0) Then Exit For
   Next i
   If i = lstDisplay.ListCount Then
      lstDisplay.AddItem lstUnit.List(lstUnit.ListIndex)

      lvwSelectedLessee.ListItems.Add , , lvwUnit.SelectedItem
      lvwSelectedLessee.ListItems(lvwSelectedLessee.ListItems.Count).SubItems(1) = lvwUnit.ListItems(lvwUnit.SelectedItem.Index).SubItems(1)
      lvwSelectedLessee.ListItems(lvwSelectedLessee.ListItems.Count).SubItems(2) = lvwUnit.ListItems(lvwUnit.SelectedItem.Index).SubItems(2)
   End If
End Sub

Private Sub cmdAddUnitAll_Click()
   Dim i As Integer, j As Integer

   If lstUnit.ListCount = 0 Then Exit Sub

   lstDisplay.Clear
   lvwSelectedLessee.ListItems.Clear

   For i = 0 To lstUnit.ListCount - 1
      lstUnit.ListIndex = i

      lstDisplay.AddItem lstUnit.List(lstUnit.ListIndex)

      lvwSelectedLessee.ListItems.Add , , lvwUnit.ListItems(i + 1)
      lvwSelectedLessee.ListItems(lvwSelectedLessee.ListItems.Count).SubItems(1) = lvwUnit.ListItems(lvwUnit.ListItems(i + 1).Index).SubItems(1)
      lvwSelectedLessee.ListItems(lvwSelectedLessee.ListItems.Count).SubItems(2) = lvwUnit.ListItems(lvwUnit.ListItems(i + 1).Index).SubItems(2)
   Next i
End Sub

Private Sub cmdAttachOK_Click()
     Frame17(0).Visible = False
     attachmentframePlayed = True
     Call cmdEmail_Click
End Sub

Private Sub cmdBack_Click()
   fraListTemplates.Visible = True
   fraPropertyUnit.Visible = False
End Sub

Private Sub CreateRecords(adoConn As ADODB.Connection)
   Dim szSQL         As String
   Dim szaUnit()     As String
   Dim i             As Integer
   Dim rstReport     As New ADODB.Recordset
   Dim rstLessee     As New ADODB.Recordset

   adoConn.Execute "UPDATE tlbLetterReports SET isPrint = '';"
   szSQL = "SELECT * FROM tlbLetterReports"
   rstReport.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   For i = 0 To lstDisplay.ListCount - 1
      szaUnit = Split(lstDisplay.List(i), " \ ")

      szSQL = "SELECT T.* " & _
              "FROM Tenants T, LeaseDetails L " & _
              "WHERE T.SageAccountNumber = L.SageAccountNumber And " & _
                  "L.Status AND " & _
                  "L.UnitNumber = '" & szaUnit(0) & "';"

      rstLessee.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      rstReport.AddNew
      rstReport!SageAccountNumber = rstLessee!SageAccountNumber
      rstReport!UnitNo = szaUnit(0)
      rstReport!LesseeName = IIf(IsNull(rstLessee!Name), "", rstLessee!Name)
      If rstLessee!InvoiceTo = "H" Then
         rstReport!AddressLine1 = IIf(IsNull(rstLessee!HOAddressLine1), "", rstLessee!HOAddressLine1)
         rstReport!AddressLine2 = IIf(IsNull(rstLessee!HOAddressLine2), "", rstLessee!HOAddressLine2)
         rstReport!AddressLine3 = IIf(IsNull(rstLessee!HOAddressLine3), "", rstLessee!HOAddressLine3)
         rstReport!AddressLine4 = IIf(IsNull(rstLessee!HOAddressLine4), "", rstLessee!HOAddressLine4)
         rstReport!PostCode = IIf(IsNull(rstLessee!HOPostCode), "", rstLessee!HOPostCode)
         rstReport!TEmail = IIf(IsNull(rstLessee!Email1), "", rstLessee!Email1)
      Else
         rstReport!AddressLine1 = IIf(IsNull(rstLessee!BillAddressLine1), "", rstLessee!BillAddressLine1)
         rstReport!AddressLine2 = IIf(IsNull(rstLessee!BillAddressLine2), "", rstLessee!BillAddressLine2)
         rstReport!AddressLine3 = IIf(IsNull(rstLessee!BillAddressLine3), "", rstLessee!BillAddressLine3)
         rstReport!AddressLine4 = IIf(IsNull(rstLessee!BillAddressLine4), "", rstLessee!BillAddressLine4)
         rstReport!PostCode = IIf(IsNull(rstLessee!BillPostCode), "", rstLessee!BillPostCode)
         rstReport!TEmail = IIf(IsNull(rstLessee!Email2), "", rstLessee!Email2)
      End If

      rstReport!PrintDate = Format(Date, "dd/mm/yyyy")
      rstReport!Subject = flxTemplates.TextMatrix(flxTemplates.row, 3)
      rstReport!Body = flxTemplates.TextMatrix(flxTemplates.row, 6)
      rstReport!SenderName = flxTemplates.TextMatrix(flxTemplates.row, 7)
      rstReport!SenderPosition = flxTemplates.TextMatrix(flxTemplates.row, 8)
      rstReport!TemplateID = lTemplateID
      rstReport!isPrint = "Y"

      rstReport.Update
      rstLessee.Close
   Next i
   rstReport.Close
   Set rstReport = Nothing
   Set rstLessee = Nothing
End Sub

Private Sub cmdCloseAttcah_Click()
     attachmentframePlayed = False
     Frame17(0).Visible = False
     cmdEmail.Enabled = True
End Sub
Private Sub cmdClinetAddAtch_Click()
   Dim ofn As OPENFILENAME
   Dim lHwnd As Long
   Const HKEY_LOCAL_MACHINE As Long = &H80000002
   Dim szOldFile_PathName As String
   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
   Dim fso As Object, adoConn As New ADODB.Connection

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir
   ofn.lpstrTitle = "Select File to attach"
   ofn.Flags = 0
   Dim j As Integer
   Dim iRow As Integer
   If GetOpenFileName(ofn) = 0 Then Exit Sub

'   If cmbFiles.ListCount > 0 Then
'        ReDim Preserve DataAttach(2, cmbFiles.ListCount) As String
'   End If
'
'
'    For j = 0 To 2
'        If j = 0 Then
'            DataAttach(0, cmbFiles.ListCount) = ofn.lpstrFileTitle
'        ElseIf j = 1 Then
'            DataAttach(1, cmbFiles.ListCount) = ofn.lpstrFile
'        ElseIf j = 2 Then
'            DataAttach(2, cmbFiles.ListCount) = "c"
'        End If
'    Next j
'
'   cmbFiles.Column = DataAttach()
'   If UBound(DataAttach, 2) >= 0 Then
'          cmbFiles.ListIndex = UBound(DataAttach, 2)
'   End If
'   If cmbFiles.text <> "" Then
'        attachmentframePlayed = True
'   End If

   iRow = flxAttachment.Rows - 1
   flxAttachment.Cols = 3
   flxAttachment.ColWidth(0) = 250
   flxAttachment.ColWidth(1) = 2000
   flxAttachment.ColWidth(2) = 4500
   If iRow > 1 Or flxAttachment.TextMatrix(iRow, 1) <> "" Then
     flxAttachment.AddItem ""
     iRow = iRow + 1
   End If
   
   flxAttachment.TextMatrix(iRow, 1) = ofn.lpstrFileTitle
   flxAttachment.TextMatrix(iRow, 2) = ofn.lpstrFile
    
   If flxAttachment.TextMatrix(1, 1) <> "" Then
        attachmentframePlayed = True
   End If
End Sub
Private Sub cmdOpenFile_Click()
'    Call OpenFile(FileName_FilePath(cmbFiles.Column(1)), FileName_FilePath(cmbFiles.Column(1)))
    On Error GoTo Err
    Dim rCount As Integer
    Dim iSelectedattachRow As Integer
   iSelectedattachRow = 0
   For rCount = 1 To flxAttachment.Rows - 1
        If flxAttachment.TextMatrix(rCount, 0) = "X" Then
            iSelectedattachRow = iSelectedattachRow + 1
        End If
   Next
   If iSelectedattachRow = 0 Then
        MsgBox "Please select a file to open", vbInformation, "Warning"
        Exit Sub
   End If
    Dim i As Integer
'    Call OpenFile(FileName_FilePath(cmbFiles.Column(1)), FileName_FilePath(cmbFiles.Column(1)))
        For i = 1 To flxAttachment.Rows - 1
               If flxAttachment.TextMatrix(i, 0) = "X" Then
                        Call OpenFile(FileName_FilePath(flxAttachment.TextMatrix(i, 2)), FileName_FilePath(flxAttachment.TextMatrix(i, 2)))
              End If
       Next i
     Exit Sub
Err:
    MsgBox Err.description
End Sub
Private Sub cmdDeleteFile_Click()
'       Dim i As Integer
'       If cmbFiles.text = "" Then Exit Sub
'
'        For i = cmbFiles.ListIndex + 1 To UBound(DataAttach, 2)
'            DataAttach(0, i - 1) = DataAttach(0, i)
'            DataAttach(1, i - 1) = DataAttach(1, i)
'            DataAttach(2, i - 1) = DataAttach(2, i)
'        Next
'        If UBound(DataAttach, 2) > 0 Then
'            ReDim Preserve DataAttach(2, UBound(DataAttach, 2) - 1)
'        End If
'        cmbFiles.Column() = DataAttach()
'        cmbFiles.text = ""
'        If UBound(DataAttach, 2) > 0 Then
'            cmbFiles.ListIndex = 0
'        End If
        Dim rCount As Integer
        Dim iSelectedattachRow As Integer
        iSelectedattachRow = 0
        For rCount = 1 To flxAttachment.Rows - 1
            If flxAttachment.TextMatrix(rCount, 0) = "X" Then
                iSelectedattachRow = iSelectedattachRow + 1
            End If
        Next
        If iSelectedattachRow = 0 Then
            MsgBox "Please select a file to remove", vbInformation, "Warning"
            Exit Sub
        End If
   
       Dim i As Integer

       For i = 1 To flxAttachment.Rows - 1
               If flxAttachment.TextMatrix(flxAttachment.row, 0) = "X" Then
                        flxAttachment.RemoveItem flxAttachment.row
              End If
       Next i
       If flxAttachment.Rows = 1 Then
            flxAttachment.Rows = 2
       End If

End Sub

Private Sub cmdEmail_Click()
   Dim K As Integer
   If lstDisplay.ListCount < 1 Then Exit Sub
   If attachmentframePlayed = False Then
        If MsgBox("Do you wish to send this letter to the selected lessees?", _
           vbQuestion + vbYesNo, "General Letter") = vbNo Then Exit Sub
   End If

   Dim szSQL            As String, iEmailDmd    As Integer
   Dim adoConn          As New ADODB.Connection
   Dim adoRST           As New ADODB.Recordset
   Dim colAtt           As New Collection
   If IsLoadedAndVisible("frmReport") Then
      MsgBox "There are open reports found. Please must close all open reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
   Dim szTemp As String
   szTemp = Replace(FullDatabasePath, "mdb", "ldb")
   If FileExists(szTemp) Then
      MsgBox "There are open demand reports on another computer. Please close all open demand reports before sending an email.", vbCritical + vbOKOnly, "Sending email"
      Exit Sub
   End If
   If attachmentframePlayed = False Then
        If MsgBox("Do you wish to add additional attachments to this email?", vbYesNo, "Additional attachments") = vbYes Then
            flxAttachment.Cols = 3
            flxAttachment.Rows = 2
            flxAttachment.ColWidth(0) = 250
            flxAttachment.ColWidth(1) = 2000
            flxAttachment.ColWidth(2) = 4500
            flxAttachment.RowHeight(0) = 0
            
            Frame17(0).Top = 2730
            Frame17(0).Left = 100
            Frame17(0).Visible = True
            Frame17(0).ZOrder 0
            cmdEmail.Enabled = False
            Exit Sub
         Else
            Frame17(0).Visible = False
         End If
    End If
  
   adoConn.Open getConnectionString

   CreateRecords adoConn

   szSQL = "SELECT * FROM tlbLetterReports " & _
           "WHERE  isPrint = 'Y' AND TEmail<>'';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   EmailDelay 10

   Dim reportApp        As New CRAXDRT.Application
   Dim Report           As CRAXDRT.Report
   Dim bEmailResult     As Boolean
   
       While Not adoRST.EOF
          
                  Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LetterTemplate.rpt")
                  Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
                  Report.EnableParameterPrompting = False
                  If Report.HasSavedData Then Report.DiscardSavedData
            
            '    Passing the from and to date values to Crystal Reports
                  Report.ParameterFields(1).AddCurrentValue adoRST.Fields.Item("TemplateID").Value
                  Report.ParameterFields(2).AddCurrentValue adoRST.Fields.Item("UnitNo").Value
            
                  szSQL = adoRST.Fields.Item("sageAccountNumber").Value & "_" & UniqueID() & ".pdf"
                  Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
                  Report.ExportOptions.DestinationType = crEDTDiskFile
                  Report.ExportOptions.FormatType = crEFTPortableDocFormat
                  Report.ExportOptions.PDFExportAllPages = True
                  Report.Export False
                  Set Report = Nothing
            
                  If colAtt.Count > 0 Then colAtt.Remove (1)
                  colAtt.Add DB_PATH & "\AllStuff\Temp\" & szSQL
                  
         
         
          'here I Need to add extra attachements
''          If attachmentframePlayed = True And cmbFiles.ListCount > 0 Then
''                For k = 0 To UBound(DataAttach, 2)
''                    colAtt.Add cmbFiles.Column(1, k)
''                Next k
''          End If
''          attachmentframePlayed = False
'          cmbFiles.Clear
'          ReDim DataAttach(2, 0) As String
           If attachmentframePlayed = True Then
                   For K = 1 To flxAttachment.Rows - 1
                         colAtt.Add flxAttachment.TextMatrix(K, 2)
                   Next K
            End If
            attachmentframePlayed = False
        '   cmbFiles.Clear
        '   ReDim DataAttach(2, 0) As String
            flxAttachment.Cols = 3
            flxAttachment.Rows = 2
            flxAttachment.ColWidth(0) = 250
            flxAttachment.ColWidth(1) = 2000
            flxAttachment.ColWidth(2) = 4500
            flxAttachment.RowHeight(0) = 0
    
                    bEmailResult = SendEmail(szFromEmail, adoRST.Fields.Item("TEmail").Value, _
                                  "General Letter", _
                                  "Please find the letter in the attachment.", , , _
                                   colAtt, adoRST.Fields.Item("sageAccountNumber").Value, "General Letter")
          
                                   
          If bEmailResult Then
             ShowMsgInTaskBar "Email sent.", "Y", "P"
          Else
             ShowMsgInTaskBar "No email sent.", "Y", "N"
          End If
    
          adoRST.MoveNext
          EmailDelay 10
       Wend
   
   

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing
   cmdEmail.Enabled = True
   attachmentframePlayed = False
End Sub

Private Sub cmdNew_Click()
   If flxTemplates.Rows = 1 Then
      ShowMsgInTaskBar "No template has been created. Goto Tools --> Templates", "Y", "N"
      Exit Sub
   End If

   If flxTemplates.row < 1 Then
      ShowMsgInTaskBar "Please select a template from the grid.", "Y", "N"
      Exit Sub
   End If

   fraListTemplates.Visible = False
   fraPropertyUnit.Visible = True

   fraPropertyUnit.Top = 0
   fraPropertyUnit.Left = 0
   bChoice = True
   lTemplateID = flxTemplates.TextMatrix(flxTemplates.row, 0)
End Sub

Private Sub cmdNext_Click()
   Dim j As Integer
   If lstDisplay.ListCount < 1 Then Exit Sub

   On Error GoTo ErrHandler

   Dim adoConn       As New ADODB.Connection
   Dim rstRst        As New ADODB.Recordset
   Dim rstLessee     As New ADODB.Recordset, rstReport As New ADODB.Recordset
   Dim szSQL         As String
   Dim i             As Integer
   Dim szaUnit()     As String
   
'*****************************************************************************************************************
'                                         Save the letter in the lessee records
'*****************************************************************************************************************
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tlbLetterReports SET isPrint = '';"
   szSQL = "SELECT * FROM tlbLetterReports"

   With rstReport
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      For i = 0 To lstDisplay.ListCount - 1
         szaUnit = Split(lstDisplay.List(i), " \ ")

         szSQL = "SELECT T.* " & _
                 "FROM Tenants T, LeaseDetails L " & _
                 "WHERE T.SageAccountNumber = L.SageAccountNumber And " & _
                     "L.Status AND " & _
                     "L.UnitNumber = '" & szaUnit(0) & "';"
'Debug.Print szSQL
         rstLessee.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         .AddNew
         !SageAccountNumber = rstLessee!SageAccountNumber
         !UnitNo = szaUnit(0)
         !LesseeName = IIf(IsNull(rstLessee!Name), "", rstLessee!Name)
         If rstLessee!InvoiceTo = "H" Then
            !AddressLine1 = IIf(IsNull(rstLessee!HOAddressLine1), "", rstLessee!HOAddressLine1)
            !AddressLine2 = IIf(IsNull(rstLessee!HOAddressLine2), "", rstLessee!HOAddressLine2)
            !AddressLine3 = IIf(IsNull(rstLessee!HOAddressLine3), "", rstLessee!HOAddressLine3)
            !AddressLine4 = IIf(IsNull(rstLessee!HOAddressLine4), "", rstLessee!HOAddressLine4)
            !PostCode = IIf(IsNull(rstLessee!HOPostCode), "", rstLessee!HOPostCode)
            !TEmail = IIf(IsNull(rstLessee!Email1), "", rstLessee!Email1)
         Else
            !AddressLine1 = IIf(IsNull(rstLessee!BillAddressLine1), "", rstLessee!BillAddressLine1)
            !AddressLine2 = IIf(IsNull(rstLessee!BillAddressLine2), "", rstLessee!BillAddressLine2)
            !AddressLine3 = IIf(IsNull(rstLessee!BillAddressLine3), "", rstLessee!BillAddressLine3)
            !AddressLine4 = IIf(IsNull(rstLessee!BillAddressLine4), "", rstLessee!BillAddressLine4)
            !PostCode = IIf(IsNull(rstLessee!BillPostCode), "", rstLessee!BillPostCode)
            !TEmail = IIf(IsNull(rstLessee!Email2), "", rstLessee!Email2)
         End If

         !PrintDate = Format(Date, "dd mmmm yyyy")
         !Subject = flxTemplates.TextMatrix(flxTemplates.row, 3)
'Debug.Print flxTemplates.TextMatrix(flxTemplates.row, 6)
         !Body = flxTemplates.TextMatrix(flxTemplates.row, 6)
         !SenderName = flxTemplates.TextMatrix(flxTemplates.row, 7)
         !SenderPosition = flxTemplates.TextMatrix(flxTemplates.row, 8)
         !TemplateID = lTemplateID
         !isPrint = "Y"

         .Update
         rstLessee.Close
      Next i
      .Close
   End With
'**************************************************************************************************************

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LetterTemplate.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

'    Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue CInt(lTemplateID)
   Report.ParameterFields(2).AddCurrentValue ListOfUnits

   Load frmReport
   frmReport.LoadReportViewer Report
   Exit Sub

ErrHandler:
   MsgBox "Please contact with PCM Consulting support", vbCritical + vbOKOnly, "Error Ref: 365294"
'  In 'tlbLetterReports' table 'TemplateID' is INDEXED -> NO DUPLICATE. Set it to 'NO'
End Sub

Private Function ListOfUnits() As String
   Dim szTemp()      As String
   Dim j             As Integer
      
   For j = 0 To lstDisplay.ListCount - 1
      szTemp = Split(lstDisplay.List(j), " \ ")
      ListOfUnits = ListOfUnits & szTemp(0) & ", "
   Next j
   If Len(ListOfUnits) > 0 Then ListOfUnits = Left(ListOfUnits, Len(ListOfUnits) - 2)
End Function

Private Sub cmdPrintWord_Click()
   Dim j As Integer
   If lstDisplay.ListCount < 1 Then Exit Sub

   On Error GoTo ErrHandler

   Dim adoConn       As New ADODB.Connection
   Dim rstRst        As New ADODB.Recordset
   Dim rstLessee     As New ADODB.Recordset, rstReport As New ADODB.Recordset
   Dim szSQL         As String
   Dim i             As Integer
   Dim szaUnit()     As String
   
'*****************************************************************************************************************
'                                         Save the letter in the lessee records
'*****************************************************************************************************************
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tlbLetterReports SET isPrint = '';"
   szSQL = "SELECT * FROM  tlbLetterReports"

   With rstReport
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      For i = 0 To lstDisplay.ListCount - 1
         szaUnit = Split(lstDisplay.List(i), " \ ")

         szSQL = "SELECT T.* " & _
                 "FROM Tenants T, LeaseDetails L " & _
                 "WHERE T.SageAccountNumber = L.SageAccountNumber And " & _
                     "L.Status AND " & _
                     "L.UnitNumber = '" & szaUnit(0) & "';"
'Debug.Print szSQL
         rstLessee.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         .AddNew
         !SageAccountNumber = rstLessee!SageAccountNumber
         !UnitNo = szaUnit(0)
         !LesseeName = IIf(IsNull(rstLessee!Name), "", rstLessee!Name)
         If rstLessee!InvoiceTo = "H" Then
            !AddressLine1 = IIf(IsNull(rstLessee!HOAddressLine1), "", rstLessee!HOAddressLine1)
            !AddressLine2 = IIf(IsNull(rstLessee!HOAddressLine2), "", rstLessee!HOAddressLine2)
            !AddressLine3 = IIf(IsNull(rstLessee!HOAddressLine3), "", rstLessee!HOAddressLine3)
            !AddressLine4 = IIf(IsNull(rstLessee!HOAddressLine4), "", rstLessee!HOAddressLine4)
            !PostCode = IIf(IsNull(rstLessee!HOPostCode), "", rstLessee!HOPostCode)
            !TEmail = IIf(IsNull(rstLessee!Email1), "", rstLessee!Email1)
         Else
            !AddressLine1 = IIf(IsNull(rstLessee!BillAddressLine1), "", rstLessee!BillAddressLine1)
            !AddressLine2 = IIf(IsNull(rstLessee!BillAddressLine2), "", rstLessee!BillAddressLine2)
            !AddressLine3 = IIf(IsNull(rstLessee!BillAddressLine3), "", rstLessee!BillAddressLine3)
            !AddressLine4 = IIf(IsNull(rstLessee!BillAddressLine4), "", rstLessee!BillAddressLine4)
            !PostCode = IIf(IsNull(rstLessee!BillPostCode), "", rstLessee!BillPostCode)
            !TEmail = IIf(IsNull(rstLessee!Email2), "", rstLessee!Email2)
         End If

         !PrintDate = Format(Date, "dd mmmm yyyy")
         !Subject = flxTemplates.TextMatrix(flxTemplates.row, 3)
'Debug.Print flxTemplates.TextMatrix(flxTemplates.row, 6)
         !Body = flxTemplates.TextMatrix(flxTemplates.row, 6)
         !SenderName = flxTemplates.TextMatrix(flxTemplates.row, 7)
         !SenderPosition = flxTemplates.TextMatrix(flxTemplates.row, 8)
         !TemplateID = lTemplateID
         !isPrint = "Y"

         .Update
         rstLessee.Close
      Next i
      .Close
   End With
'**************************************************************************************************************
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LetterTemplate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   '    Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue CInt(lTemplateID)
   Report.ParameterFields(2).AddCurrentValue ListOfUnits

   szSQL = "Sending Batch Letters" & "_" & UniqueID() & ".DOC"
   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTWordForWindows
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing
   
   OpenFile szSQL, DB_PATH & "\AllStuff\Temp\"
   Exit Sub

ErrHandler:
   MsgBox "Please contact with PCM Consulting support", vbCritical + vbOKOnly, "Error Ref: 365294"
End Sub

Private Sub cmdRemUnit_Click()
   Dim j As Integer

   If lstDisplay.SelCount < 1 Then Exit Sub

   Do
      If lstDisplay.Selected(j) Then
         lstDisplay.RemoveItem lstDisplay.ListIndex
         
         lvwSelectedLessee.ListItems.Remove lvwSelectedLessee.SelectedItem.Index
         
         j = j - 1
      End If
      j = j + 1
   Loop While j <= lstDisplay.ListCount - 1
End Sub

Private Sub cmdRemUnitAll_Click()
   If lstDisplay.ListCount < 0 Then Exit Sub
   lstDisplay.Clear
   lvwSelectedLessee.ListItems.Clear
   lstUnit.ListIndex = -1
End Sub

Private Sub cmdSelectdiff_Click(Index As Integer)
   Dim sSQLQuery_ As String

   fraListTemplates.Visible = True
End Sub

Private Sub LoadRecords_P_U(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim i       As Integer
   Dim K       As Integer
   Dim adoRST  As New ADODB.Recordset

   szSQL = "SELECT Units.PropertyID, Property.PropertyName, Units.UnitNumber, " & _
                  "Units.UnitName, IIF(InvoiceTo = 'H', Email1, Email2) AS EM, " & _
                  "Tenants.SageAccountNumber, Name " & _
           "FROM Tenants INNER JOIN ((Property INNER JOIN " & _
                  "Units ON Property.PropertyID = Units.PropertyID) INNER JOIN " & _
                  "LeaseDetails ON Units.UnitNumber = LeaseDetails.UnitNumber) ON " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber " & _
           "WHERE (((LeaseDetails.Status)=True)) " & _
           "ORDER BY Name;"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'Debug.Print szSQL
   If Not adoRST.EOF Then
      ReDim szP_U(6, adoRST.RecordCount - 1) As String

      i = 0
      While Not adoRST.EOF
         szP_U(0, i) = IIf(IsNull(adoRST.Fields.Item(0).Value), "", adoRST.Fields.Item(0).Value)
         szP_U(1, i) = IIf(IsNull(adoRST.Fields.Item(1).Value), "", adoRST.Fields.Item(1).Value)
         szP_U(2, i) = IIf(IsNull(adoRST.Fields.Item(2).Value), "", adoRST.Fields.Item(2).Value)
         szP_U(3, i) = IIf(IsNull(adoRST.Fields.Item(3).Value), "", adoRST.Fields.Item(3).Value)
         szP_U(4, i) = IIf(IsNull(adoRST.Fields.Item(4).Value), "", adoRST.Fields.Item(4).Value)
         szP_U(5, i) = IIf(IsNull(adoRST.Fields.Item(5).Value), "", adoRST.Fields.Item(5).Value)
         szP_U(6, i) = IIf(IsNull(adoRST.Fields.Item(6).Value), "", adoRST.Fields.Item(6).Value)

         i = i + 1
         adoRST.MoveNext
      Wend
   End If

   adoRST.Close
   Set adoRST = Nothing

'  Load Property
   If lstProperty.ListCount = 0 Then
      K = 0
      While K <= UBound(szP_U, 2)
         For i = 0 To K
            If szP_U(0, i) = szP_U(0, K) Then Exit For
         Next i

         If i = K Then
            lstProperty.AddItem szP_U(0, K) & " \ " & szP_U(1, K)
         End If
         K = K + 1
      Wend
   End If
End Sub

Private Sub flxAttachment_Click()
    If flxAttachment.TextMatrix(flxAttachment.row, 0) = "X" Then
           flxAttachment.TextMatrix(flxAttachment.row, 0) = ""
     ElseIf flxAttachment.TextMatrix(flxAttachment.row, 1) <> "" Then
           flxAttachment.TextMatrix(flxAttachment.row, 0) = "X"
     End If
End Sub

Private Sub Form_Load()
   Me.Height = 4710
   Me.Width = 11580
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   attachmentframePlayed = False
'   If cmbFiles.ListCount = 0 Then
'        ReDim DataAttach(2, cmbFiles.ListCount) As String
'   End If
   flxAttachment.Clear
   flxAttachment.Rows = 2
   flxAttachment.Cols = 2
   flxAttachment.ColWidth(0) = 250
   flxAttachment.ColWidth(1) = 2000
   flxAttachment.ColWidth(2) = 4500
   flxAttachment.RowHeight(0) = 0
   
   Me.BackColor = MODULEBACKCOLOR
   fraListTemplates.BackColor = MODULEBACKCOLOR
   fraPropertyUnit.BackColor = MODULEBACKCOLOR
   bChoice = False
   bSaved = False
   bValidate = True

   If szLetter = "GL" Then Me.Caption = "General Letters"
   If szLetter = "RL" Then Me.Caption = "Reminder Letters"

   Dim sSQLQuery_    As String
   Dim listSQL       As String

   fraListTemplates.Top = 0
   fraListTemplates.Left = 0
   ConfigFlxTemplates

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadFlxTemplate adoConn
   LoadRecords_P_U adoConn

   adoConn.Close
   Set adoConn = Nothing

   flxTemplates.row = 0
   ListViewConfigure

   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxTemplate(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim iRow    As Integer
   Dim adoRST  As New ADODB.Recordset

   If szLetter = "RL" Then
      szSQL = "SELECT TemplateID, TemplateName, ModifiedBy, Description, " & _
                     "TemplateDate, Salutation, Body, SenderName, SenderPosition, TempType " & _
              "FROM Template " & _
              "WHERE TemplateName <> 'BACS Email Template' AND " & _
                    "TemplateName <> 'Demand Email Template' AND " & _
                    "TempType = 'RT';"
   Else
      szSQL = "SELECT TemplateID, TemplateName, ModifiedBy, Description, " & _
                     "TemplateDate, Salutation, Body, SenderName, SenderPosition, TempType " & _
              "FROM Template " & _
              "WHERE TemplateName <> 'BACS Email Template' AND " & _
                    "TemplateName <> 'Demand Email Template' AND " & _
                    "TempType <> 'RT';"
   End If
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   With flxTemplates
      While Not adoRST.EOF
         If Not adoRST.EOF Then .AddItem ""
         .TextMatrix(iRow, 0) = adoRST.Fields.Item("TemplateID").Value
         .TextMatrix(iRow, 1) = adoRST.Fields.Item("TemplateName").Value
         .TextMatrix(iRow, 2) = adoRST.Fields.Item("ModifiedBy").Value
         .TextMatrix(iRow, 3) = adoRST.Fields.Item("Description").Value
         .TextMatrix(iRow, 4) = Format(adoRST.Fields.Item("TemplateDate").Value, "dd/mm/yyyy")
         .TextMatrix(iRow, 5) = adoRST.Fields.Item("Salutation").Value
         .TextMatrix(iRow, 6) = adoRST.Fields.Item("Body").Value
         .TextMatrix(iRow, 7) = adoRST.Fields.Item("SenderName").Value
         .TextMatrix(iRow, 8) = adoRST.Fields.Item("SenderPosition").Value
         .TextMatrix(iRow, 9) = adoRST.Fields.Item("TempType").Value

         adoRST.MoveNext
         iRow = iRow + 1
      Wend
   End With
   
   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ListViewConfigure()
   lvwUnit.ColumnHeaders(1).Width = 0
   lvwUnit.ColumnHeaders(2).Width = 1000
   lvwUnit.ColumnHeaders(3).Width = lvwUnit.Width - 1000
   lvwSelectedLessee.ColumnHeaders(1).Width = 0
   lvwSelectedLessee.ColumnHeaders(2).Width = 1000
   lvwSelectedLessee.ColumnHeaders(3).Width = lvwSelectedLessee.Width - 1000
End Sub

Private Sub ConfigFlxTemplates()
   Dim szFlxHeader As String

   szFlxHeader$ = "|<|<|<|<|||||"

   With flxTemplates
      .FormatString = szFlxHeader$
      .Cols = 10
      .ColWidth(0) = 0
      .ColWidth(1) = 1500
      .ColWidth(2) = 1500
      .ColWidth(3) = 6500
      .ColWidth(4) = 1000
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(9) = 0
      .Rows = 1
      .RowHeight(0) = 0
   End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'If Not frmMMain.IsRibbonVersion() Then frmMMain.fraCmdButton.Enabled = True
   flxTemplates.Clear
   Unload Me
End Sub

Private Sub lstProperty_Click()
   Dim i          As String
   Dim K          As String
   Dim szTemp()   As String

   lstUnit.Clear
   szTemp = Split(lstProperty.text, " \ ")

   K = 0
   lvwUnit.ListItems.Clear
   While K <= UBound(szP_U, 2)
      If szTemp(0) = szP_U(0, K) Then
         lstUnit.AddItem szP_U(2, K) & " \ " & szP_U(3, K)

         lvwUnit.ListItems.Add , , szP_U(2, K) & " \ " & szP_U(3, K)
         lvwUnit.ListItems(lvwUnit.ListItems.Count).SubItems(1) = szP_U(5, K)
         lvwUnit.ListItems(lvwUnit.ListItems.Count).SubItems(2) = szP_U(6, K)
      End If
      K = K + 1
   Wend
End Sub

Private Sub lstUnit_DblClick()
   cmdAddUnit_Click
End Sub

Private Sub txtBody_Change()
   bSaved = False
End Sub

Private Sub txtPosition_Change()
   bSaved = False
End Sub

Private Sub txtSalutation_Change()
   bSaved = False
End Sub

Private Sub txtSender_Change()
   bSaved = False
End Sub

Private Sub cmdSelect_Click()
   If flxTemplates.Rows = 1 Then
      ShowMsgInTaskBar "There is no saved template.", "Y", "N"
      Exit Sub
   End If
   If flxTemplates.TextMatrix(flxTemplates.row, 0) = "" Then
      ShowMsgInTaskBar "Please select a record to edit.", "Y", "N"
      Exit Sub
   End If

   fraListTemplates.Visible = False

   bChoice = False
   bSaved = True
End Sub

Private Sub lvwSelectedLessee_Click()
    On Error GoTo Err
   If lvwSelectedLessee.ListItems.Count = 0 Then Exit Sub
   lstDisplay.ListIndex = lvwSelectedLessee.SelectedItem.Index - 1
Err:
End Sub

Private Sub lvwUnit_Click()
   On Error GoTo Err
   If lvwUnit.ListItems.Count = 0 Then Exit Sub
   lstUnit.ListIndex = lvwUnit.SelectedItem.Index - 1
Err:
End Sub

Private Sub lvwUnit_DblClick()
   cmdAddUnit_Click
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
