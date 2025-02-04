VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAttachment 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attachments"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAttachment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8265
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
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
         Left            =   120
         Picture         =   "frmAttachment.frx":058A
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
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
         Left            =   120
         Picture         =   "frmAttachment.frx":09CC
         TabIndex        =   3
         Top             =   920
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   120
         Picture         =   "frmAttachment.frx":1296
         TabIndex        =   2
         Top             =   1600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNew 
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
         Left            =   120
         Picture         =   "frmAttachment.frx":16D8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAttList 
      Height          =   2895
      Left            =   45
      TabIndex        =   4
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
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
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   45
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "New File Name"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   45
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
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
      Left            =   120
      TabIndex        =   6
      Top             =   40
      Width           =   705
   End
End
Attribute VB_Name = "frmAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OwnerID As String
Public CallerForm As String

Private Sub cmdAddNew_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If AddNewAttachment(CallerForm, OwnerID, adoConn) Then
      MsgBox "File has been attached successfull, thanks"
      LoadFlxAttList adoConn
      If InStr(CallerForm, "Unit") > 0 Then frmUnits2.HEALTH_N_SAFETY_ATTACH = True
      If InStr(CallerForm, "Property") > 0 Then frmProperty2.HEALTH_N_SAFETY_ATTACH = True
   Else
      MsgBox "File has not been saved successfully, try again"
   End If

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdClose_Click()
   Form_Unload -1
End Sub

Private Sub cmdDelete_Click()
   If flxAttList.TextMatrix(flxAttList.row, 0) = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & flxAttList.TextMatrix(flxAttList.row, 1) & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub

   DeleteAttachmentFileID flxAttList.TextMatrix(flxAttList.row, 0)

   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdOpen_Click()
   If flxAttList.TextMatrix(flxAttList.row, 0) = "" Then Exit Sub
   MousePointer = vbHourglass

   If OpenFile(flxAttList.TextMatrix(flxAttList.row, 3), App.Path & "\" & flxAttList.TextMatrix(flxAttList.row, 2)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation

   MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   ConfigureFlxAttList
   LoadFlxAttList adoConn

   adoConn.Close
   Set adoConn = Nothing

   flxAttList.row = 0
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   If InStr(CallerForm, "Unit") > 0 Then
      Me.Top = frmUnits2.Top + frmUnits2.Height / 2
      Me.Left = frmUnits2.Left + frmUnits2.Width / 2
   End If
   If InStr(CallerForm, "Property") > 0 Then
      Me.Top = frmProperty2.Top + frmProperty2.Height / 2
      Me.Left = frmProperty2.Left + frmProperty2.Width / 2
   End If
End Sub

Private Sub LoadFlxAttList(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim iRow As Integer, SQLStr As String

   SQLStr = "SELECT A.FileID, A.FileName, S.Value as NewFilePath, A.NewFileName " & _
            "FROM AttachedFile AS A, SecondaryCode AS S " & _
            "WHERE A.OwnerID = '" & OwnerID & "' And " & _
               "A.Entity = '" & CallerForm & "' And " & _
               "S.PrimaryCode = 'FPATH' And " & _
               "S.Code = A.NewFilePath;"
'Debug.Print SQLStr
   adoRST.Open SQLStr, adoConn, adOpenStatic, adLockOptimistic

   iRow = 1
   While Not adoRST.EOF
      flxAttList.TextMatrix(iRow, 0) = adoRST.Fields.Item(0).Value
      flxAttList.TextMatrix(iRow, 1) = adoRST.Fields.Item(1).Value
      flxAttList.TextMatrix(iRow, 2) = adoRST.Fields.Item(2).Value
      flxAttList.TextMatrix(iRow, 3) = adoRST.Fields.Item(3).Value
      adoRST.MoveNext
      If Not adoRST.EOF Then
         flxAttList.AddItem ""
         iRow = iRow + 1
      End If
   Wend
   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigureFlxAttList()
   Dim szHeader As String

   szHeader$ = "<ID|<File Name|<New File Name|>Path"

   flxAttList.Clear
   flxAttList.Cols = 4
   flxAttList.Rows = 2
   flxAttList.RowHeight(0) = 0

   flxAttList.ColWidth(0) = 0                   'ID
   flxAttList.ColWidth(1) = 1800
   flxAttList.ColWidth(2) = 2200
   flxAttList.ColWidth(3) = 2100
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If CallerForm = "Unit" Then frmUnits2.Enabled = True
   If CallerForm = "Unit_Insurance" Then frmUnits2.Enabled = True
   If CallerForm = "Property" Then frmProperty2.Enabled = True
   If CallerForm = "Property_Insurance" Then frmProperty2.Enabled = True

   Unload Me
End Sub
