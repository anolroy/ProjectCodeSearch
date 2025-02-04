VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lessee Account Statement Summary"
   ClientHeight    =   6285
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRAS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   5670
      ScaleHeight     =   4740
      ScaleWidth      =   5535
      TabIndex        =   15
      Top             =   135
      Visible         =   0   'False
      Width           =   5565
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
         Left            =   5190
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   17
         Top             =   675
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   7091
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   21
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
         Left            =   1620
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   375
         Width           =   3825
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6747;450"
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
         Height          =   240
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5085
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Properties:"
      Height          =   4110
      Left            =   90
      TabIndex        =   8
      Top             =   1260
      Width           =   5595
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPropertyLookup 
         Height          =   2805
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   4948
         _Version        =   393216
         ForeColor       =   0
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
         BackColorSel    =   15329508
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         HighLight       =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   14
         Top             =   240
         Width           =   180
         VariousPropertyBits=   276824083
         Caption         =   "ID"
         Size            =   "317;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   1110
         TabIndex        =   13
         Top             =   240
         Width           =   1245
         VariousPropertyBits=   276824083
         Caption         =   "Properties Name"
         Size            =   "2196;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   195
         Index           =   2
         Left            =   4350
         TabIndex        =   12
         Top             =   240
         Width           =   735
         VariousPropertyBits=   276824083
         Caption         =   "Post Code"
         Size            =   "1296;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkSelectAllRow 
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   195
         Width           =   240
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "423;450"
         Value           =   "0"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdGenerateReportButton 
         Height          =   495
         Left            =   1815
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
         Caption         =   "Print"
         Size            =   "2143;873"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.CommandButton cmdClientList 
      Caption         =   ".."
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
      Left            =   5370
      TabIndex        =   6
      Top             =   90
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Income Collection Period:"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5565
      Begin VB.TextBox txtFromDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtToDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   210
      End
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1140
      TabIndex        =   7
      Top             =   90
      Width           =   4230
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "7461;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmRAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Name : Lessee Summary Statement
Private szPropertyList As String
Dim sTextBox As String
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL Clients"
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cboClientList_Click()
   Dim i As Integer

   chkSelectAllRow.Value = False

   For i = 1 To flxPropertyLookup.Rows - 1
      flxPropertyLookup.RowHeight(i) = 240
   Next i

   If txtClientList.text = "ALL" Then Exit Sub

   For i = 1 To flxPropertyLookup.Rows - 1
      If flxPropertyLookup.TextMatrix(i, 4) <> txtClientList.text Then
         flxPropertyLookup.RowHeight(i) = 0
         flxPropertyLookup.TextMatrix(i, 0) = ""
      End If
   Next i
End Sub

Private Sub cboClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub chkSelectAllRow_Click()
   Dim i As Integer

   For i = 1 To flxPropertyLookup.Rows - 1
      If chkSelectAllRow.Value And flxPropertyLookup.RowHeight(i) = 240 Then
         If flxPropertyLookup.TextMatrix(i, 0) = "" Then
            SelectFlxGridRow 0, flxPropertyLookup, i
         End If
      Else
         If flxPropertyLookup.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxPropertyLookup, i
         End If
      End If
   Next i
End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    Frame1.Enabled = True
    Frame2.Enabled = True
    picClient.Visible = False
    cmdClientList.SetFocus
End Sub

Private Sub flxPropertyLookup_Click()
   SelectFlxGridRow 0, flxPropertyLookup, flxPropertyLookup.row

   If flxPropertyLookup.TextMatrix(flxPropertyLookup.row, 0) = "X" Then
      chkSelectAllRow.Value = False
   End If
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

'   PrepareList adoConn, cboClientList
   Frame2.BackColor = MODULEBACKCOLOR
   txtClientList.text = "ALL"
   ConfigFlxPropertyLookup
   LoadProperties adoConn

   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 5865
   Me.Height = 5910
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub
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
Private Sub flxClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp And flxClient.row = 1 Then
        txtSearchClientID.SetFocus
     End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
        End If
        picClient.Visible = False
    End If
    If KeyAscii = 27 Then
         picClient.Visible = False
          Frame1.Enabled = True
          Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
          
           End If
    End If
End Sub

Private Sub flxClient_Click()
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 1)
                cboClientList_Click
        End If
        picClient.Visible = False
        
End Sub
Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            flxClient.SetFocus
        End If
    End If
End Sub
Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
              flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxClient.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchClientName.SetFocus
    End If
End Sub
Public Function LoadProperties(adoConn As ADODB.Connection)
   Dim rstProperty_ As New ADODB.Recordset
   Dim iRow As Integer, szSQL As String

   szSQL = "SELECT PROPERTYID, PROPERTYNAME, ProPostCode, ClientID " & _
           "FROM   PROPERTY " & _
           "ORDER BY PROPERTYID;"

   rstProperty_.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   On Error Resume Next
   While Not rstProperty_.EOF
      flxPropertyLookup.TextMatrix(iRow, 1) = IIf(rstProperty_.Fields.Item(0) = Null, "", rstProperty_.Fields.Item(0))
      flxPropertyLookup.TextMatrix(iRow, 2) = IIf(rstProperty_.Fields.Item(1) = Null, "", rstProperty_.Fields.Item(1))
      flxPropertyLookup.TextMatrix(iRow, 3) = IIf(rstProperty_.Fields.Item(2) = Null, "", rstProperty_.Fields.Item(2))
      flxPropertyLookup.TextMatrix(iRow, 4) = IIf(rstProperty_.Fields.Item(3) = Null, "", rstProperty_.Fields.Item(3))

      rstProperty_.MoveNext
      If Not rstProperty_.EOF Then flxPropertyLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstProperty_.Close
   Set rstProperty_ = Nothing
End Function

Private Sub ConfigFlxPropertyLookup()
   flxPropertyLookup.Clear
   flxPropertyLookup.Rows = 2
   flxPropertyLookup.Cols = 5

   flxPropertyLookup.ColWidth(0) = Label2(0).Left - flxPropertyLookup.Left
   flxPropertyLookup.ColWidth(1) = Label2(1).Left - Label2(0).Left
   flxPropertyLookup.ColWidth(2) = Label2(2).Left - Label2(1).Left
   flxPropertyLookup.ColWidth(3) = flxPropertyLookup.Width - Label2(2).Left - 120
   flxPropertyLookup.ColWidth(4) = 0
   flxPropertyLookup.RowHeight(0) = 0
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub cmdGenerateReportButton_Click()
   If txtFromDate.text = "" Then
      ShowMsgInTaskBar "Please enter FROM date."
      txtFromDate.SetFocus
      Exit Sub
   End If
   If txtToDate.text = "" Then
      ShowMsgInTaskBar "Please enter TO date."
      txtToDate.SetFocus
      Exit Sub
   End If
   cmdGenerateReportButton.Enabled = False
   If CreateListOfProp = 0 Then
      ShowMsgInTaskBar "Please select Property from the grid."
      flxPropertyLookup.SetFocus
      Exit Sub
   End If

   Dim adoConn    As New ADODB.Connection
   Dim datetype   As Integer
   Dim reportApp  As New CRAXDRT.Application
   Dim Report     As CRAXDRT.Report

   adoConn.Open getConnectionString
   

   MarkPropOfSelection adoConn 'putting a flag into the database

   adoConn.Close
   Set adoConn = Nothing

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\RentAccountStatement.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   ' Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
   Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)

   Load frmReport
   frmReport.LoadReportViewer Report
   cmdGenerateReportButton.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Function CreateListOfProp() As Integer
   Dim i As Integer

   szPropertyList = ""
   
   For i = 0 To flxPropertyLookup.Rows - 1
      If flxPropertyLookup.TextMatrix(i, 0) = "X" And flxPropertyLookup.RowHeight(i) > 0 Then
         szPropertyList = "'" & flxPropertyLookup.TextMatrix(i, 1) & "'" & ", " & szPropertyList
      End If
   Next i
   If Len(szPropertyList) > 2 Then
      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
      CreateListOfProp = Len(szPropertyList)
      Exit Function
   End If
   CreateListOfProp = 0
End Function

Private Sub MarkPropOfSelection(adoConn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = '';"
   adoConn.Execute szSQL

   szSQL = "UPDATE PROPERTY " & _
           "SET    RAS = 'X' " & _
           "WHERE  PropertyID IN (" & szPropertyList & ");"
'Debug.Print szSQL
   adoConn.Execute szSQL
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtToDate.SetFocus
    End If
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxPropertyLookup.SetFocus
    End If
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate
End Sub


