VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRptCategory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Categories"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRptCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleMode       =   0  'User
   ScaleWidth      =   8884.932
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtClientID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6306
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraCodes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   7695
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   600
         Width           =   6075
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   1
         Top             =   135
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4080
         MaxLength       =   255
         TabIndex        =   2
         Top             =   135
         Width           =   3435
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   1500
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Clos&e"
         Height          =   375
         Left            =   6245
         TabIndex        =   7
         Top             =   5760
         Width           =   1275
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   1500
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "De&lete"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   5760
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRptCat 
         Height          =   3720
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   7400
         _ExtentX        =   13044
         _ExtentY        =   6562
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   17
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Category &Name:"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   165
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cate&gory Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   165
         Width           =   1155
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Client:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   465
   End
   Begin MSForms.ComboBox cmbClient 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5325
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "9393;503"
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
End
Attribute VB_Name = "frmRptCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RowIndex As Integer

Private Sub cmbClient_Click()
   txtClientID.text = cmbClient.Value

   Dim iRow As Integer

   For iRow = 1 To flxRptCat.Rows - 1
      flxRptCat.TextMatrix(iRow, 0) = ""
      If flxRptCat.TextMatrix(iRow, 5) = txtClientID.text Then
         flxRptCat.RowHeight(iRow) = 245
      Else
         If flxRptCat.TextMatrix(iRow, 1) <> "" Then flxRptCat.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   If txtCode = "" Then Exit Sub

   If MsgBox("Do you wish to delete the category code: " & flxRptCat.TextMatrix(flxRptCat.row, 2) & "", vbQuestion + vbYesNo, "Deliting category code") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   adoConn.Execute "DELETE * FROM ReportCategory WHERE RecordID = '" & flxRptCat.TextMatrix(flxRptCat.row, 1) & "';"
   
   adoConn.Close
   Set adoConn = Nothing

   flxRptCat.RemoveItem flxRptCat.row
   ShowMsgInTaskBar "Category code has been deleted", "Y", "P"
End Sub

Private Sub cmdEdit_Click()
   txtCode.text = ""
   txtName.text = ""
   txtDescription.text = ""
   flxRptCat.Enabled = True
   RowIndex = -1
End Sub

Private Sub cmdNew_Click()
   If txtClientID.text = "" Then
      ShowMsgInTaskBar "Please select a client to continue", "Y", "N"
      cmbClient.SetFocus
      Exit Sub
   End If
   If Trim(txtCode.text) = "" Then
      ShowMsgInTaskBar "Please enter the category code", "Y", "N"
      txtCode.SetFocus
      Exit Sub
   End If
   If Trim(txtName.text) = "" Then
      ShowMsgInTaskBar "Please enter the category name", "Y", "N"
      txtName.SetFocus
      Exit Sub
   End If
   If flxRptCat.Enabled And IsCodeExists Then
      ShowMsgInTaskBar "Code already exist", "Y", "N"
      Exit Sub
   End If

   If flxRptCat.Enabled Then
      If flxRptCat.TextMatrix(flxRptCat.Rows - 1, 1) <> "" Then flxRptCat.AddItem ""
      flxRptCat.TextMatrix(flxRptCat.Rows - 1, 1) = UniqueID()
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   If RowIndex = -1 Then               'Add New
      RowIndex = flxRptCat.Rows - 1

      szSQL = "SELECT S.* FROM ReportCategory AS S;"
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      szSQL = UniqueID()
      flxRptCat.TextMatrix(RowIndex, 1) = szSQL
      adoRst.AddNew
      adoRst.Fields.Item("RecordID").Value = szSQL
   Else
      szSQL = "SELECT S.* " & _
              "FROM ReportCategory AS S " & _
              "WHERE S.RecordID = '" & flxRptCat.TextMatrix(RowIndex, 1) & "';"
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   End If

   adoRst.Fields.Item("CategoryCode").Value = txtCode.text
   adoRst.Fields.Item("CategoryName").Value = txtName.text
   adoRst.Fields.Item("ClientID").Value = txtClientID.text
   adoRst.Fields.Item("CatDesc").Value = txtDescription.text
   adoRst.Update
   adoRst.Close

   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
   
   flxRptCat.TextMatrix(RowIndex, 2) = txtCode.text
   flxRptCat.TextMatrix(RowIndex, 3) = txtName.text
   flxRptCat.TextMatrix(RowIndex, 4) = cmbClient.text
   flxRptCat.TextMatrix(RowIndex, 5) = txtClientID.text
   flxRptCat.TextMatrix(RowIndex, 6) = txtDescription.text

   cmdEdit_Click
End Sub

Private Function IsCodeExists() As Boolean
   Dim iRow As Integer

   IsCodeExists = False

   For iRow = 1 To flxRptCat.Rows - 1
      If flxRptCat.TextMatrix(iRow, 5) = txtClientID.text And _
            flxRptCat.TextMatrix(iRow, 2) = txtCode.text Then
         IsCodeExists = True
         Exit For
      End If
   Next iRow
End Function

Private Sub flxRptCat_DblClick()
   RowIndex = flxRptCat.row
   txtCode.text = flxRptCat.TextMatrix(RowIndex, 2)
   txtName.text = flxRptCat.TextMatrix(RowIndex, 3)
   cmbClient.text = flxRptCat.TextMatrix(RowIndex, 4)
   txtClientID.text = flxRptCat.TextMatrix(RowIndex, 5)
   txtDescription.text = flxRptCat.TextMatrix(RowIndex, 6)
End Sub

Private Sub flxRptCat_RowColChange()
'   RowIndex = flxRptCat.row
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 7740
   Me.Height = 7200
   Me.BackColor = MODULEBACKCOLOR
   fraCodes.BackColor = MODULEBACKCOLOR
   RowIndex = -1

   ConfigFlxRptCat

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadCmbClient adoConn, cmbClient
   LoadFlxRptCat adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxRptCat(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim iRow As Integer
   Dim szSQL As String

   flxRptCat.Rows = 1
   flxRptCat.AddItem ""
   iRow = 1

   szSQL = "SELECT S.*, C.ClientName " & _
           "FROM ReportCategory AS S INNER JOIN Client AS C ON S.ClientID = C.ClientID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      flxRptCat.TextMatrix(iRow, 1) = adoRst.Fields.Item("RecordID").Value
      flxRptCat.TextMatrix(iRow, 2) = adoRst.Fields.Item("CategoryCode").Value
      flxRptCat.TextMatrix(iRow, 3) = adoRst.Fields.Item("CategoryName").Value
      flxRptCat.TextMatrix(iRow, 4) = adoRst.Fields.Item("ClientName").Value
      flxRptCat.TextMatrix(iRow, 5) = adoRst.Fields.Item("ClientID").Value
      flxRptCat.TextMatrix(iRow, 6) = IIf(IsNull(adoRst.Fields.Item("CatDesc").Value), "", adoRst.Fields.Item("CatDesc").Value)
      iRow = iRow + 1
      adoRst.MoveNext
      If Not adoRst.EOF Then flxRptCat.AddItem ""
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub ConfigFlxRptCat()
   Dim szHeader As String

   flxRptCat.Clear
   flxRptCat.Cols = 7
   flxRptCat.Rows = 2
   flxRptCat.RowHeight(0) = 0

   szHeader$ = "X|ID|<Code|<Name|<Client|ClientID|Description"

   flxRptCat.FormatString = szHeader$
   flxRptCat.ColWidth(0) = 0
   flxRptCat.ColWidth(1) = 0
   flxRptCat.ColWidth(2) = Label5(1).Left - Label5(0).Left
   flxRptCat.ColWidth(3) = Label5(2).Left - Label5(1).Left
   flxRptCat.ColWidth(4) = flxRptCat.Width + flxRptCat.Left - Label5(2).Left - 340
   flxRptCat.ColWidth(5) = 0
   flxRptCat.ColWidth(6) = 0
End Sub

Private Sub LoadCmbClient(adoConn As ADODB.Connection, cboC As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
        Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If IsLoadedAndVisible("frmNLAmendment") Then
      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString
      frmNLAmendment.LoadLstCategoryCodes adoConn
      frmNLAmendment.Enabled = True

      adoConn.Close
      Set adoConn = Nothing
   End If

   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtCode_Change()
   If RowIndex > -1 Then flxRptCat.Enabled = False
End Sub

Private Sub txtDescription_Change()
   If RowIndex > -1 Then flxRptCat.Enabled = False
End Sub

Private Sub txtName_Change()
   If RowIndex > -1 Then flxRptCat.Enabled = False
End Sub

Private Sub Label2_Click()
   txtName.SetFocus
End Sub

Private Sub Label3_Click()
   txtCode.SetFocus
End Sub

Private Sub Label4_Click()
   txtDescription.SetFocus
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
