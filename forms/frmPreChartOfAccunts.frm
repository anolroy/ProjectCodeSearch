VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreChartOfAccunts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chart of Accounts"
   ClientHeight    =   5100
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   6615
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
   Icon            =   "frmPreChartOfAccunts.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   420
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4590
      Width           =   1575
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   420
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4590
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   3825
      Left            =   120
      TabIndex        =   3
      Top             =   735
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6747
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
   Begin MSForms.TextBox TextBox1 
      Height          =   255
      Left            =   5085
      TabIndex        =   2
      Top             =   405
      Width           =   1395
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2461;450"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtSearchClientName 
      Height          =   255
      Left            =   1575
      TabIndex        =   1
      Top             =   405
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
   Begin MSForms.TextBox txtSearchClientID 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   405
      Width           =   1170
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2064;450"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
      Height          =   195
      Index           =   1
      Left            =   1545
      TabIndex        =   7
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   2
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID"
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   6
      Top             =   120
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   90
      Width           =   6375
   End
End
Attribute VB_Name = "frmPreChartOfAccunts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Call ExportData2NominalLedger(adoConn)

   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub
'
'Private Sub cboClientList_Click()
'   If cboClientList.ListCount > 0 Then FilterProperties
'End Sub
'
'Private Sub FilterProperties()
'   Dim i As Integer, j As Integer, r As Integer
'   Dim szaTemp() As String
'
'   For i = 1 To flxProperties.Rows - 1
'      flxProperties.RowHeight(i) = 240
'   Next i
'   If cboClientList.Value > 0 Then
'      For i = 1 To flxProperties.Rows - 1
'         If flxProperties.TextMatrix(i, 3) <> cboClientList.Column(0) Then
'            flxProperties.RowHeight(i) = 0
'         End If
'      Next i
'   End If
'End Sub
'
'Private Sub chkProp_Click()
'   Dim i As Integer
'
'   If chkProp.Value = 1 Then
'      For i = 1 To flxProperties.Rows - 1
'         If flxProperties.TextMatrix(i, 0) <> "X" Then
'            SelectFlxGridRow 0, flxProperties, i
'         End If
'      Next i
'   Else
'      For i = 1 To flxProperties.Rows - 1
'         If flxProperties.TextMatrix(i, 0) = "X" Then
'            SelectFlxGridRow 0, flxProperties, i
'         End If
'      Next i
'   End If
'End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If CountSelClients <> 1 Then
      ShowMsgInTaskBar "Please select one client to generate the report", "Y", "N"
      Exit Sub
   End If

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NL_List.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue flxProperties.TextMatrix(flxProperties.row, 1)

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Function CountSelClients() As String
   Dim i As Integer

   CountSelClients = 0
   For i = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(i, 0) = "X" And flxProperties.RowHeight(i) > 0 Then
         CountSelClients = CountSelClients + 1
      End If
   Next i
End Function

Private Sub flxProperties_Click()
     SelectFlxGridRow 0, flxProperties, flxProperties.row
End Sub

Private Sub flxProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
         SelectFlxGridRow 0, flxProperties, flxProperties.row
    End If
End Sub

Private Sub flxProperties_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSCYRRPrint.SetFocus
    End If
End Sub

Private Sub flxProperties_RowColChange()
'   SelectFlxGridRow 0, flxProperties, flxProperties.row
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 5580
   Me.Width = 6705
   Me.BackColor = MODULEBACKCOLOR
   ConfigFlxProperties

'   connect to database
   adoConn.Open getConnectionString

   PrepareList adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
'End Sub
'
'Private Sub txtSCYRREnDt_Change()
'    TextBoxChangeDate txtSCYRREnDt
'End Sub
'
'Private Sub txtSCYRREnDt_GotFocus()
'   If Len(txtSCYRREnDt.text) < 10 Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
'   SelTxtInCtrl txtSCYRREnDt
'End Sub
'
'Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
'    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
'End Sub
'
'Private Sub txtSCYRREnDt_LostFocus()
'    If txtSCYRREnDt.text <> "" Then TextBoxFormatDate txtSCYRREnDt
'End Sub
'
'Private Sub txtSCYRRStDt_Change()
'    TextBoxChangeDate txtSCYRRStDt
'End Sub
'
'Private Sub txtSCYRRStDt_GotFocus()
'   If Len(txtSCYRRStDt.text) < 10 Then txtSCYRRStDt.text = Format(Date, "dd/mm/yyyy")
'   SelTxtInCtrl txtSCYRRStDt
'End Sub
'
'Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
'    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
'End Sub
'
'Private Sub txtSCYRRStDt_LostFocus()
'    If txtSCYRRStDt.text <> "" Then TextBoxFormatDate txtSCYRRStDt
'End Sub

Private Sub PrepareList(adoConn As ADODB.Connection)
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim r       As Integer

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 1

   While Not adoRst.EOF
      flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("CLIENTID").Value
      flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("CLIENTNAME").Value
      flxProperties.TextMatrix(r, 3) = IIf(IsNull(adoRst.Fields.Item("CLIENTPOSTCODE").Value), "", adoRst.Fields.Item("CLIENTPOSTCODE").Value)

      r = r + 1
      adoRst.MoveNext
      If Not adoRst.EOF Then flxProperties.AddItem ""
   Wend

   adoRst.Close
   Set adoRst = Nothing

   flxProperties.row = 0
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

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
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

Private Sub txtSearchClientID_Change()
        'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxProperties.Rows - 1 To 1 Step -1
      flxProperties.RowHeight(i) = 240
      flxProperties.TextMatrix(i, 0) = ""
     ' If sTextBox = "1" Then
            If InStr(1, UCase(flxProperties.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
                flxProperties.RowHeight(i) = 0
            End If
       'End If
      If flxProperties.RowHeight(i) = 240 Then
            flxProperties.row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientName_Change()
       'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxProperties.Rows - 1 To 1 Step -1
            flxProperties.RowHeight(i) = 240
            flxProperties.TextMatrix(i, 0) = ""
            If InStr(1, UCase(flxProperties.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
                  flxProperties.RowHeight(i) = 0
            End If
      
      If flxProperties.RowHeight(i) = 240 Then
            flxProperties.row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxProperties.SetFocus
    End If
End Sub
