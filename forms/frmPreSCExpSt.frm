VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreSCExpSt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Charge Expenditure Statement Report"
   ClientHeight    =   6030
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   8925
   ClipControls    =   0   'False
   DrawMode        =   2  'Blackness
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreSCExpSt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   45
      ScaleHeight     =   4335
      ScaleWidth      =   5130
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   5160
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
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   5415
         Left            =   45
         TabIndex        =   17
         Top             =   675
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   9551
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
         Width           =   3465
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6112;450"
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
         Width           =   4770
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
      Left            =   4935
      TabIndex        =   0
      Top             =   90
      Width           =   345
   End
   Begin VB.CheckBox chkProp 
      Appearance      =   0  'Flat
      Caption         =   "Select All"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Range:"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   3855
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   2685
      Left            =   120
      TabIndex        =   7
      Top             =   735
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4736
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12632256
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
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   90
      Width           =   4185
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "7382;503"
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID"
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property Name"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   480
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   450
      Width           =   5175
   End
End
Attribute VB_Name = "frmPreSCExpSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Dim sTextBox As String

Private Sub flxClient_Click()
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
'                txtPropertyName.text = "ALL"
'                txtPropertyName.Tag = "ALL"
                Dim adoConn As New ADODB.Connection
                adoConn.Open getConnectionString
'                LoadCmbFinancialYear adoConn
                ConfigFlxProperties
'                LoadProperty adoConn
                FilterProperties
                adoConn.Close
                flxProperties.SetFocus
        End If
        picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
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

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
          'If sTextBox = "1" Then
           cmdClientList.SetFocus
'           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
'           ElseIf sTextBox = "3" Then
'                cmdFundLookUp.SetFocus
           'End If
    End If
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
Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Call ExportData2NominalLedger(adoConn)

   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub

'Private Sub cboClientList_Click()
''   If cboClientList.ListCount > 0 Then FilterProperties
'End Sub

Private Sub FilterProperties()
   Dim i As Integer, j As Integer, r As Integer
   Dim szaTemp() As String

   For i = 1 To flxProperties.Rows - 1
      flxProperties.RowHeight(i) = 240
   Next i
   If txtClientList.Tag = "ALL" Then Exit Sub
   If txtClientList.Tag <> "" Then
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 3) <> txtClientList.Tag Then
            flxProperties.RowHeight(i) = 0
         End If
      Next i
   End If
End Sub

Private Sub chkProp_Click()
   Dim i As Integer

   If chkProp.Value = 1 Then
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) <> "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
   Else
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
   End If
End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
'    Frame2.Enabled = False
'    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
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
            flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = "ALL"
               flxClient.TextMatrix(rRow, 2) = "ALL Client"
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

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
'    Frame1.Enabled = True
'    Frame2.Enabled = True
    cmdClientList.SetFocus
End Sub


Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
   If txtSCYRRStDt.text = "" Then
      txtSCYRRStDt.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim szProperties As String

   szProperties = ListOfProperties

   If Len(szProperties) = 0 Then
      ShowMsgInTaskBar "Please select at least one property to generate the report", "Y", "N"
      Exit Sub
   End If

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SC_St_Preview.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy"))
   Report.ParameterFields(2).AddCurrentValue CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy"))
   Report.ParameterFields(3).AddCurrentValue ListOfProperties

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Function ListOfProperties() As String
   Dim i As Integer

   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" And flxProperties.RowHeight(i) > 0 Then
         ListOfProperties = ListOfProperties & flxProperties.TextMatrix(i, 1) & ", "
      End If
   Next i
   If Len(ListOfProperties) > 0 Then ListOfProperties = Left(ListOfProperties, Len(ListOfProperties) - 2)
'Debug.Print ListOfProperties
End Function

Private Sub flxProperties_RowColChange()
   SelectFlxGridRow 0, flxProperties, flxProperties.row
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 5145
   Me.Width = 5595
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   chkProp.BackColor = MODULEBACKCOLOR

   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
   ConfigFlxProperties

'   connect to database
   adoConn.Open getConnectionString

   PrepareList adoConn

   adoConn.Close
   Set adoConn = Nothing

   chkProp.Value = 1

   Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtSCYRREnDt_Change()
    TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
   If Len(txtSCYRREnDt.text) < 10 Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSCYRRPrint.SetFocus
    End If
    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
    If txtSCYRREnDt.text <> "" Then TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
   If Len(txtSCYRRStDt.text) < 10 Then txtSCYRRStDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSCYRREnDt.SetFocus
    End If
    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
    If txtSCYRRStDt.text <> "" Then TextBoxFormatDate txtSCYRRStDt
End Sub

Private Sub ConfigFlxProperties()
   Dim szHeader As String

   szHeader$ = "<|<|<|<"
   With flxProperties
      .FormatString = szHeader
      .Cols = 4
      .RowHeight(0) = 0
      .ColWidth(0) = Label2(0).Left - .Left '200                 '"X"
      .ColWidth(1) = Label2(1).Left - Label2(0).Left 'Property ID
      .ColWidth(2) = Label2(2).Left - Label2(1).Left 'Property Name
      .ColWidth(3) = .Width + .Left - Label2(2).Left - 300 'Client ID
   End With
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim r       As Integer

'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTID;"
''
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   Data(0, 0) = "ALL"
'   Data(1, 0) = "All Clients"
'   For i = 1 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'
'   cboClient.Column() = Data()
'   cboClient.ListIndex = 0
   
        txtClientList.text = "All Clients"
        txtClientList.Tag = "ALL"
  

'------------------------------------------------------------------------------------------
   szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY " & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 1

   While Not adoRst.EOF
      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      flxProperties.TextMatrix(r, 3) = adoRst.Fields.Item("ClientID").Value

      r = r + 1
      adoRst.MoveNext
      If Not adoRst.EOF Then flxProperties.AddItem ""
   Wend

   adoRst.Close
   flxProperties.row = 0
'------------------------------------------------------------------------------------------

   Set adoRst = Nothing
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
