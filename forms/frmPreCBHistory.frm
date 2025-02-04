VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreCBHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cashbook History Transactions"
   ClientHeight    =   8790
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   16320
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreCBHistory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6067.016
   ScaleMode       =   0  'User
   ScaleWidth      =   15325.33
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   5535
      ScaleHeight     =   3795
      ScaleWidth      =   6255
      TabIndex        =   25
      Top             =   225
      Visible         =   0   'False
      Width           =   6285
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
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3075
         Left            =   45
         TabIndex        =   27
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   5424
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   33
         Top             =   375
         Width           =   4545
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   32
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   31
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   30
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
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
         Width           =   5850
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
      Left            =   4905
      TabIndex        =   0
      Top             =   225
      Width           =   300
   End
   Begin VB.CommandButton cmdBC 
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
      Left            =   4905
      TabIndex        =   1
      Top             =   675
      Width           =   300
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date options"
      Height          =   465
      Left            =   135
      TabIndex        =   21
      Top             =   1305
      Visible         =   0   'False
      Width           =   5190
      Begin VB.OptionButton Option1 
         Caption         =   "By Date"
         Height          =   195
         Left            =   3195
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Financial Year"
         Height          =   195
         Left            =   1530
         TabIndex        =   2
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select the Financial Period"
      Height          =   1410
      Left            =   5310
      TabIndex        =   17
      Top             =   1620
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period To:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   20
         Top             =   1065
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From:"
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   19
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year:"
         Height          =   255
         Index           =   66
         Left            =   600
         TabIndex        =   18
         Top             =   345
         Width           =   975
      End
      Begin MSForms.ComboBox cmbFinancialYear 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   330
         Width           =   2760
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4868;503"
         TextColumn      =   2
         ColumnCount     =   5
         ListRows        =   20
         cColumnInfo     =   5
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0;0"
      End
      Begin MSForms.ComboBox cmbPeriodTo 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1050
         Width           =   1920
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
      Begin MSForms.ComboBox cmbPeriodFrom 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   690
         Width           =   1920
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;1940;0;0"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Year"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   1275
      Width           =   5175
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRefreshData 
      Caption         =   "&Refresh Data"
      Height          =   360
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3915
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3285
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3330
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCashBook 
      Height          =   4020
      Left            =   135
      TabIndex        =   16
      Top             =   4590
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   7091
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      Enabled         =   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label5 
      Caption         =   "Client"
      Height          =   240
      Left            =   270
      TabIndex        =   24
      Top             =   225
      Width           =   1095
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1620
      TabIndex        =   23
      Top             =   225
      Width           =   3285
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "5794;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtBC 
      Height          =   285
      Left            =   1620
      TabIndex        =   22
      Top             =   675
      Width           =   3285
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "5794;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label84 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account:"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   720
      Width           =   1035
   End
End
Attribute VB_Name = "frmPreCBHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String
Dim sTextBox As String
Private Sub LoadCmbFinancialYear(adoconn As ADODB.Connection)
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer              'Open Flag index
   Dim adoRst     As New ADODB.Recordset
If Trim(txtClientList.text) = "" Then Exit Sub
   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status " & _
           "FROM   FinancialYear " & _
           "WHERE  ClientID = '" & txtClientList.Tag & "' " & _
           "ORDER BY FY_StDate Desc;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1
   ReDim Data(TotalCol, TotalRow) As String

   K = -1
   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
         If K = -1 And j = 4 Then
            If adoRst.Fields("Status").Value Then
               K = i
'               dtStartPnL = CDate(adoRst.Fields("FY_StDate").Value)
'               dtStartBS = CDate("01 January 2000")
            End If
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cmbFinancialYear.Column() = Data()
   cmbFinancialYear.ListIndex = K

   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

NoRes:
   Set adoRst = Nothing
   ShowMsgInTaskBar "Financial year has not been created for the client", "Y", "N"
   Exit Sub
End Sub
Private Sub cmbFinancialYear_Change()
        Dim adoconn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim szSQL      As String
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer                    'Open flag index

   If Not IsNull(cmbFinancialYear.Value) Then
      adoconn.Open getConnectionString
      
      szSQL = "SELECT PeriodID, Period_Descp, P_StDate, P_EndDate, Status " & _
              "FROM   Periods " & _
              "WHERE  FYrID = '" & cmbFinancialYear.Value & "' " & _
              "ORDER BY P_StDate;"

'      Debug.Print szSQL
      
      adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

      If adoRst.EOF Then GoTo NoRes

      TotalRow = adoRst.RecordCount - 1
      TotalCol = adoRst.Fields.count - 1
      ReDim Data(TotalCol, TotalRow) As String

      K = -1
      For i = 0 To TotalRow
         For j = 0 To TotalCol
            Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
            If K = -1 And j = 4 Then
               If adoRst.Fields("Status").Value Then
                  K = i
'                  dtEnd = CDate(adoRst.Fields("P_EndDate").Value)
               End If
            End If
         Next j
         adoRst.MoveNext
         If adoRst.EOF Then Exit For
      Next i
      
      cmbPeriodFrom.Column() = Data()
      cmbPeriodTo.Column() = Data()
      
      cmbPeriodFrom.ListIndex = 0
      If (cmbPeriodTo.ListCount > 0) Then
         cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If

      adoconn.Close
      Set adoconn = Nothing
   End If
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Periods are not found. Please contact with system support", "Y", "N"
   Set adoconn = Nothing
End Sub



Private Sub cmbFinancialYear_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmbPeriodFrom.SetFocus
    End If
End Sub

Private Sub cmbPeriodFrom_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmbPeriodTo.SetFocus
    End If
End Sub

Private Sub cmbPeriodTo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdSCYRRPrint.SetFocus
    End If
End Sub

Private Sub cmdBC_Click()
     picClient.Left = 269.029
    picClient.Top = 355.299
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    sTextBox = "2"
    ConfigureFlxBank
    Dim szAllBankBalance As String
    szAllBankBalance = BankAndBalance(adoconn)
    adoconn.Close
    Set adoconn = Nothing
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub loadflxclient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
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

   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
                 rRow = 1
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub ConfigureFlxBank()
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment = vbLeftJustify
   lblClientID.Caption = "Bank Code"
   lblClientName.Caption = "Bank Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
End Sub


Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    cmdClientList.SetFocus
End Sub

Private Sub flxClient_Click()
            Dim adoconn As New ADODB.Connection
            adoconn.Open getConnectionString
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    txtBC.Tag = ""
                    txtBC.text = ""
                    LoadCmbFinancialYear adoconn
                    cmdBC.SetFocus
            ElseIf sTextBox = "2" Then
                    txtBC.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtBC.text = flxClient.TextMatrix(flxClient.row, 2)
                    LoadFlxCashBook adoconn
                    'Option2.SetFocus
                    txtSCYRRStDt.SetFocus
            End If
            adoconn.Close
            picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame1.Visible = False
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbFinancialYear.SetFocus
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
Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    loadflxclient
    
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
'Private Sub cmdFundLookUp_Click()
'   If cboClientID.text = "" Then
'      ShowMsgInTaskBar "Please select a client to continue.", , "N"
'      Exit Sub
'   End If
'
'   fmePropertyLookup.Top = Frame1.Top + Frame1.Height + 5
'   fmePropertyLookup.Left = Frame1.Left - (fmePropertyLookup.Width - Frame1.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   gridPropertyLookup.Visible = True
'   txtSearchProperty.text = ""
'   txtSearchProperty.Enabled = True
'   txtSearchProperty.SetFocus
'
'   LOOKUPCommand = "Fund"
'
'   PopulatePropertyLookup IIf(cboClientID.Value = "ALL", "", " WHERE CLIENTID = '" & cboClientID.Value & "'")
'End Sub
'
'Private Sub cmdGridPropertyLookup_Click()
'   fmePropertyLookup.Visible = False
'End Sub
'
'Private Sub cmdPropertyLookup_Click()
'   If cboClientID.text = "" Then
'      ShowMsgInTaskBar "Please select a client to continue.", , "N"
'      Exit Sub
'   End If
'
'   fmePropertyLookup.Top = txtPropertyID.Top + txtPropertyID.Height + 5
'   fmePropertyLookup.Left = txtPropertyID.Left - (fmePropertyLookup.Width - txtPropertyID.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   gridPropertyLookup.Visible = True
'   txtSearchProperty.text = ""
'   txtSearchProperty.Enabled = True
'   txtSearchProperty.SetFocus
'
'   LOOKUPCommand = "Property"
'
'   PopulatePropertyLookup IIf(cboClientID.Value = "ALL", "", " WHERE CLIENTID = '" & cboClientID.Value & "'")
'End Sub

Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass
   
   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString

   Call ExportData2NominalLedger(adoconn)

   adoconn.Close
   Set adoconn = Nothing
   'Me.MousePointer = vbArrow
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub
Private Sub CreateTable(adoconn As ADODB.Connection)
    
     Dim adoRst As New ADODB.Recordset
     On Error GoTo CreateReportCashbookHistory
       
       adoRst.Open "SELECT * FROM ReportCashbookHistory;", adoconn, adOpenStatic, adLockReadOnly
       adoRst.Close
    
       GoTo alreadycreated
    
CreateReportCashbookHistory:
           adoconn.Execute _
              "CREATE TABLE ReportCashbookHistory " & _
                 "(" & _
                    "ReportingDate DateTime  NOT NULL, " & _
                    "SessionID     TEXT(100) NOT NULL, " & _
                    "ClientID      TEXT(10), " & _
                    "iRow      Number, " & _
                    "TDate DateTime, " & _
                    "No   TEXT(50) NOT NULL, " & _
                    "tTYpe      TEXT(100), " & _
                    "Account      TEXT(100), " & _
                    "Reference      TEXT(200), " & _
                    "Detail      TEXT(250), " & _
                    "Debit       CURRENCY, " & _
                    "Credit         CURRENCY, " & _
                    "Reconciled        TEXT(10), " & _
                    "StDate        TEXT(20), " & _
                    "PRIMARY KEY (ReportingDate, SessionID, iRow)" & _
                 ");"
        
alreadycreated:
End Sub
Private Sub cmdSCYRRPrint_Click()
    Dim i As Integer
    If txtBC.text = "" Then
      MsgBox "Please select a Bank.", vbCritical + vbOKOnly, "Cashbook History"
      txtBC.SetFocus
      Exit Sub
   End If
   'added by anol 08  Aug 2016
   If Option1.Value = True Then
        If txtSCYRRStDt.text = "" Then
            MsgBox "Please enter the starting date", vbInformation, "Warning"
            txtSCYRRStDt.SetFocus
            Exit Sub
        End If
        If txtSCYRREnDt.text = "" Then
            MsgBox "Please enter the end date", vbInformation, "Warning"
            txtSCYRREnDt.SetFocus
            Exit Sub
        End If
        
        'height change flexgrid  based on date
     
      For i = 1 To flxCashBook.Rows - 1
         flxCashBook.RowHeight(i) = 240
      Next i

      For i = 1 To flxCashBook.Rows - 1
         If CDate(flxCashBook.TextMatrix(i, 1)) < CDate(txtSCYRRStDt.text) Or _
               CDate(flxCashBook.TextMatrix(i, 1)) > CDate(txtSCYRREnDt.text) Then
            flxCashBook.RowHeight(i) = 0
         End If
      Next i
      
   Else ' For financial Radio button
     
      For i = 1 To flxCashBook.Rows - 1
         flxCashBook.RowHeight(i) = 240
      Next i

      For i = 1 To flxCashBook.Rows - 1
         If CDate(flxCashBook.TextMatrix(i, 1)) < CDate(CStr(cmbPeriodFrom.Column(2))) Or _
               CDate(flxCashBook.TextMatrix(i, 1)) > CDate(cmbPeriodTo.Column(3)) Then
            flxCashBook.RowHeight(i) = 0
         End If
      Next i
   
   End If
    
  
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Dim rsAdd As New ADODB.Recordset
    Call CreateTable(adoconn)
    Dim sessionID As String
    Dim reportingDate As String
    
    reportingDate = Format(Date, "dd mmmm yyyy")
    sessionID = GetTimeStamp
    adoconn.Execute _
    "DELETE FROM ReportCashbookHistory WHERE SessionID = '" & sessionID & "';"
    adoconn.Execute "DELETE FROM ReportCashbookHistory WHERE ReportingDate < #" & reportingDate & "# ;"
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
    rsAdd.Open " Select * from ReportCashbookHistory where 1=2", adoconn, adOpenKeyset, adLockBatchOptimistic
    With rsAdd
    For i = 1 To flxCashBook.Rows - 1
             If flxCashBook.RowHeight(i) <> 0 Then
                .AddNew
                !reportingDate = reportingDate
                !sessionID = sessionID
                !clientID = frmCashbook.txtClientList.Tag
                !iRow = i
                !TDate = flxCashBook.TextMatrix(i, 1)
                !No = flxCashBook.TextMatrix(i, 2)
                !TType = flxCashBook.TextMatrix(i, 10)
                !Account = flxCashBook.TextMatrix(i, 3)
                !Reference = flxCashBook.TextMatrix(i, 4)
                !Detail = flxCashBook.TextMatrix(i, 5)
                !Debit = Val(flxCashBook.TextMatrix(i, 6))
                !Credit = Val(flxCashBook.TextMatrix(i, 7))
                !Reconciled = flxCashBook.TextMatrix(i, 8)
                !StDate = flxCashBook.TextMatrix(i, 9)
             End If
             .UpdateBatch
    Next
    End With
    adoconn.Close
    'Call Sleep(100)
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookHistoryRpt.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue txtBC.Tag
   '#
        If Option2.Value = True Then
            'By financial year
            Report.ParameterFields(2).AddCurrentValue CDate(CStr(cmbPeriodFrom.Column(2)))
            Report.ParameterFields(3).AddCurrentValue CDate(cmbPeriodTo.Column(3))
            Report.ParameterFields(4).AddCurrentValue frmCashbook.CalDrCrAcBalance(CDate(CStr(cmbPeriodFrom.Column(2))), CDate(cmbPeriodTo.Column(3)))
            
        Else
             Report.ParameterFields(2).AddCurrentValue CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy"))
             Report.ParameterFields(3).AddCurrentValue CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy"))
             Report.ParameterFields(4).AddCurrentValue CalDrCrAcBalance(CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy")), CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy")))
        End If
        '#
   
   
   'resolved by BOSL
   'issue 523 cashbook report was not working properly
   'added by anol 19 Jan 2015
   Report.ParameterFields(5).AddCurrentValue txtClientList.Tag
   Report.ParameterFields(6).AddCurrentValue sessionID
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Function CalDrCrAcBalance(dtStart As Date, dtEnd As Date) As Double
   Dim iRow As Integer, cDr As Currency, cCr As Currency

   On Error Resume Next

   For iRow = 1 To flxCashBook.Rows - 1
      If CDate(flxCashBook.TextMatrix(iRow, 1)) >= dtStart And _
            CDate(flxCashBook.TextMatrix(iRow, 1)) <= dtEnd Then
         cDr = cDr + CCur(flxCashBook.TextMatrix(iRow, 6))
         cCr = cCr + CCur(flxCashBook.TextMatrix(iRow, 7))
      End If
   Next iRow

   CalDrCrAcBalance = cDr - cCr
End Function
'
'Private Sub UpdateDrCr4NC(adoConn As ADODB.Connection)
'   Dim szSQL As String, szPropSrc As String
'   Dim szFundPA As String, szFundDEPT_ID As String, szFundSageDepartment As String
'   Dim szFundSA As String, szFundSR As String, szFundBank As String
'   Dim adoRst As New ADODB.Recordset, adoNL As New ADODB.Recordset
'
'   If txtPropertyID.text = "ALL" Then
'      szPropSrc = ""
'   Else
'      szPropSrc = "P.PropertyID = '" & txtPropertyID.text & "' AND "
'   End If
'
'   If txtFundNo.text = "ALL" Then
'      szFundDEPT_ID = ""
'      szFundSageDepartment = ""
'      szFundPA = ""
'      szFundSA = ""
'      szFundSR = ""
'      szFundBank = ""
'   Else
'      szFundDEPT_ID = "S.DEPT_ID = '" & txtFundNo.text & "' AND "
'      szFundPA = "P.FundID = " & txtFundNo.text & " AND "
'      szFundSageDepartment = "S.SageDepartment = " & txtFundNo.text & " AND "
'      szFundSA = "R.FundID = " & txtFundNo.text & " AND "
'      szFundSR = "S.SageDepartment = " & txtFundNo.text & " AND "
'      szFundBank = "B.DEPT_ID = '" & txtFundNo.text & "' AND "
'   End If
'--------------------------------------------------------------------------------------------------
'##########                            Purchase Invoices & Credit - PI, PC
'--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT S.NOMINAL_CODE AS NC, SUM(S.TOTAL_AMOUNT) AS T, P.TransactionType AS TT " & _
'           "FROM tblPurInvSRec AS S INNER JOIN tblPurInv AS P ON S.ParentID = P.MY_ID " & _
'           "WHERE " & szFundDEPT_ID & _
'              szPropSrc & _
'              "P.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "P.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY S.NOMINAL_CODE, P.TransactionType;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         If adoRst.Fields.Item(2).Value = 6 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
'         If adoRst.Fields.Item(2).Value = 7 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
'      Else
'         If adoRst.Fields.Item(2).Value = 6 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
'         If adoRst.Fields.Item(2).Value = 7 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
'      End If
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(S.TOTAL_AMOUNT) AS T, P.TransactionType AS TT " & _
'           "FROM tblPurInvSRec AS S INNER JOIN tblPurInv AS P ON S.ParentID = P.MY_ID " & _
'           "WHERE " & szFundDEPT_ID & _
'              szPropSrc & _
'              "P.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "P.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY P.TransactionType;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         If adoRst.Fields.Item(1).Value = 6 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
'         If adoRst.Fields.Item(1).Value = 7 Then adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
'      Else
'         If adoRst.Fields.Item(1).Value = 6 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
'         If adoRst.Fields.Item(1).Value = 7 Then adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
'      End If
'      adoRst.MoveNext
'   Wend
'
'   adoNL.Update
'
'   adoNL.Close
'   adoRst.Close
'
''--------------------------------------------------------------------------------------------------
''##########                             Payment on AC - PA
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT P.BankCode AS NC, SUM(Amount) AS T " & _
'           "FROM   tlbPayment AS P " & _
'           "WHERE  P.Type = 9 AND " & _
'              szFundPA & _
'              "P.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "P.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY P.BankCode;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item("NC").Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
'               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'      Else
'         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + _
'               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(Amount) AS T " & _
'           "FROM tlbPayment AS P " & _
'           "WHERE  P.Type = 9 AND " & _
'              szFundPA & _
'              "P.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "P.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "#;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'            IIf(IsNull(adoRst.Fields.Item("T").Value), 0, adoRst.Fields.Item("T").Value)
'   Else
''If IsNull(adoRST.Fields.Item("T").Value) Then
'      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'            IIf(IsNull(adoRst.Fields.Item("T").Value), 0, adoRst.Fields.Item("T").Value)
'   End If
'
'   adoNL.Update
'
'   adoNL.Close
'   adoRst.Close
'
''--------------------------------------------------------------------------------------------------
''##########                             Purchase Payment - PP
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SQ.NC, SUM(SQ.A) AS T " & _
'           "FROM (" & _
'               "SELECT P1.BankCode AS NC, P1.Amount AS A, P1.TransactionID " & _
'               "FROM   tlbPayment AS P1, tlbPayment AS P2, PayTransactions AS P, tblPurInvSRec AS S " & _
'               "WHERE  P1.Type = 8 AND P1.TransactionID = P.FromTran AND " & _
'                  "P2.TransactionID = P.ToTran AND P2.PI = S.ParentID AND " & _
'                  szFundDEPT_ID & _
'                  "P1.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'                  "P1.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'               "GROUP BY P1.TransactionID, P1.BankCode, P1.Amount" & _
'           ") AS SQ " & _
'           "GROUP BY SQ.NC;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item("NC").Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'      Else
'         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'               IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '2100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(SQ.A) AS T " & _
'           "FROM (" & _
'               "SELECT P1.Amount AS A, P1.TransactionID " & _
'               "FROM   tlbPayment AS P1, tlbPayment AS P2, PayTransactions AS P, tblPurInvSRec AS S " & _
'               "WHERE  P1.Type = 8 AND P1.TransactionID = P.FromTran AND " & _
'                  "P2.TransactionID = P.ToTran AND P2.PI = S.ParentID AND " & _
'                  szFundDEPT_ID & _
'                  "P1.PDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'                  "P1.PDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'               "GROUP BY P1.TransactionID, P1.Amount" & _
'           ") AS SQ;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'            IIf(IsNull(adoRst.Fields.Item(0).Value), 0, (adoRst.Fields.Item(0).Value))
'   Else
'      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'            IIf(IsNull(adoRst.Fields.Item(0).Value), 0, (adoRst.Fields.Item(0).Value))
'   End If
'
'   adoNL.Update
'
'   adoNL.Close
'   adoRst.Close
'
''--------------------------------------------------------------------------------------------------
''##########                             Sales Invoice and Credit - SI & SC
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT S.NominalCodeforAmount AS NC, SUM(S.TotalAmount) AS T, D.TransactionType AS TT " & _
'           "FROM DemandSplitRecords AS S, DemandRecords AS D, Property AS P, Units AS U " & _
'           "WHERE " & szFundSageDepartment & " S.DemandID = D.DemandID AND " & _
'              szPropSrc & " P.PropertyID = U.PropertyID AND " & _
'              "D.IssueDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "D.IssueDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# AND " & _
'              "U.UnitNumber = D.UnitNumber " & _
'           "GROUP BY S.NominalCodeforAmount, D.TransactionType;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         If adoRst.Fields.Item(2).Value = 1 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
'                  IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'         If adoRst.Fields.Item(2).Value = 2 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'                  IIf(IsNull(adoRst.Fields.Item("T").Value), 0, (adoRst.Fields.Item("T").Value))
'      Else
'         If adoRst.Fields.Item(2).Value = 1 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Debit").Value) + (adoRst.Fields.Item("T").Value)
'         If adoRst.Fields.Item(2).Value = 2 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Debit").Value) - (adoRst.Fields.Item("T").Value)
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(S.TotalAmount) AS T, D.TransactionType AS TT " & _
'           "FROM DemandSplitRecords AS S, DemandRecords AS D, Property AS P, Units AS U " & _
'           "WHERE " & szFundSageDepartment & " S.DemandID = D.DemandID AND " & _
'              szPropSrc & _
'              "D.IssueDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "D.IssueDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# AND " & _
'              "U.UnitNumber = D.UnitNumber AND P.PropertyID = U.PropertyID " & _
'           "GROUP BY D.TransactionType;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         If adoRst.Fields.Item("TT").Value = 1 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + _
'                                                Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                                adoRst.Fields.Item("T").Value))
'         If adoRst.Fields.Item("TT").Value = 2 Then _
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'                                                Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                                adoRst.Fields.Item("T").Value))
'      Else
'         If adoRst.Fields.Item("TT").Value = 1 Then _
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + _
'                                               Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                               adoRst.Fields.Item("T").Value))
'         If adoRst.Fields.Item("TT").Value = 2 Then _
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'                                               Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                               adoRst.Fields.Item("T").Value))
'      End If
'
'      adoNL.Update
'   End If
'
'   adoNL.Close
'   adoRst.Close
'
''--------------------------------------------------------------------------------------------------
''##########                             Sales Receipt on Account - SA
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT R.BankCode, SUM(Amount) AS T " & _
'           "FROM   tlbReceipt AS R " & _
'           "WHERE  R.Type = 4 AND " & _
'              szFundSA & _
'              "R.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "R.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY R.BankCode;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item("BankCode").Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'                                             Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                             adoRst.Fields.Item("T").Value))
'      Else
'         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'                                            Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                            adoRst.Fields.Item("T").Value))
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(Amount) AS T " & _
'           "FROM tlbReceipt AS R " & _
'           "WHERE R.Type = 4 AND " & _
'              szFundSA & _
'              "R.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "R.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "#;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'                                          Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                          adoRst.Fields.Item("T").Value))
'   Else
'      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'                                         Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                         adoRst.Fields.Item("T").Value))
'   End If
'
'   adoNL.Update
'
'   adoNL.Close
'   adoRst.Close
''--------------------------------------------------------------------------------------------------
''##########                             Sales Receipt - SR
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SQ.NC, SUM(SQ.A) AS T " & _
'           "FROM (" & _
'               "SELECT R1.BankCode AS NC, R1.Amount AS A, R1.TransactionID " & _
'               "FROM   tlbReceipt AS R1, tlbReceipt AS R2, RptTransactions AS R, DemandSplitRecords AS S " & _
'               "WHERE  R1.Type = 3 AND R1.TransactionID = R.FromTran AND " & _
'                  "R2.TransactionID = R.ToTran AND R2.DemandRef = S.DemandID AND " & _
'                  szFundSR & _
'                  "R1.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'                  "R1.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'               "GROUP BY R1.TransactionID, R1.BankCode, R1.Amount" & _
'           ") AS SQ " & _
'           "GROUP BY SQ.NC;"
'
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
'      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - _
'                                             Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                             adoRst.Fields.Item("T").Value))
'      Else
'         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - _
'                                            Val(IIf(IsNull(adoRst.Fields.Item("T").Value), 0, _
'                                            adoRst.Fields.Item("T").Value))
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'   adoNL.Close
'   adoRst.Close
'
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT SUM(SQ.A) AS T " & _
'           "FROM (" & _
'               "SELECT R1.Amount AS A, R1.TransactionID " & _
'               "FROM   tlbReceipt AS R1, tlbReceipt AS R2, RptTransactions AS R, DemandSplitRecords AS S " & _
'               "WHERE  R1.Type = 3 AND R1.TransactionID = R.FromTran AND " & _
'                  "R2.TransactionID = R.ToTran AND R2.DemandRef = S.DemandID AND " & _
'                  szFundSR & _
'                  "R1.RDate >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'                  "R1.RDate <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'               "GROUP BY R1.TransactionID, R1.Amount" & _
'           ") AS SQ"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'      adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item(0).Value)
'   Else
'      adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item(0).Value)
'   End If
'
'   adoNL.Update
'
'   adoNL.Close
'   adoRst.Close
'
''--------------------------------------------------------------------------------------------------
''##########                             Bank Payment and Receipt - BP & BR
''--------------------------------------------------------------------------------------------------
'   adoNL.Open "SELECT * FROM NominalLedger;", adoConn, adOpenDynamic, adLockOptimistic
'
'   szSQL = "SELECT B.BANK_AC, (SUM(NET_AMOUNT) + SUM(VAT)) AS T, B.TRANS " & _
'           "FROM   tlbBankPayment AS B " & _
'           "WHERE " & _
'              szFundBank & _
'              "B.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "B.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY B.BANK_AC, B.TRANS;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
'      If adoRst.Fields.Item(2).Value = "BR" Then
'         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
'         Else
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
'         End If
'      End If
'      If adoRst.Fields.Item(2).Value = "BP" Then
'         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
'         Else
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
'         End If
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
''   adoNL.Close
'   adoRst.Close
'
'   szSQL = "SELECT B.NOMINAL_CODE, (SUM(NET_AMOUNT) + SUM(VAT)) AS T, B.TRANS " & _
'           "FROM   tlbBankPayment AS B " & _
'           "WHERE " & _
'              szFundBank & _
'              "B.TRAN_DATE >= #" & Format(txtSCYRRStDt.text, "DD MMMM YYYY") & "# AND " & _
'              "B.TRAN_DATE <= #" & Format(txtSCYRREnDt.text, "DD MMMM YYYY") & "# " & _
'           "GROUP BY B.NOMINAL_CODE, B.TRANS;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      adoNL.Find "Code = '" & adoRst.Fields.Item(0).Value & "'", , , 1
''      If adoNL.Fields.Item("DrCr").Value = "Cr" Then
''         adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRST.Fields.Item("T").Value)
''      Else
''         adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRST.Fields.Item("T").Value)
''      End If
'      If adoRst.Fields.Item(2).Value = "BR" Then
'         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) + Val(adoRst.Fields.Item("T").Value)
'         Else
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) + Val(adoRst.Fields.Item("T").Value)
'         End If
'      End If
'      If adoRst.Fields.Item(2).Value = "BP" Then
'         If adoNL.Fields.Item("DrCr").Value = "Cr" Then
'            adoNL.Fields.Item("Credit").Value = Val(adoNL.Fields.Item("Credit").Value) - Val(adoRst.Fields.Item("T").Value)
'         Else
'            adoNL.Fields.Item("Debit").Value = Val(adoNL.Fields.Item("Debit").Value) - Val(adoRst.Fields.Item("T").Value)
'         End If
'      End If
'
'      adoRst.MoveNext
'      adoNL.Update
'   Wend
'
'   adoNL.Close
'   adoRst.Close
'
'   Set adoNL = Nothing
'   Set adoRst = Nothing
'End Sub

Private Sub Form_Load()
   Me.Height = 5025
   Me.Width = 6720
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Frame2.Left = 126.772
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   Frame3.BackColor = MODULEBACKCOLOR
   Option2.BackColor = MODULEBACKCOLOR
   Option1.BackColor = MODULEBACKCOLOR
   Label5.BackColor = MODULEBACKCOLOR
   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")

   Dim szSQL As String
   Dim adoconn As New ADODB.Connection
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   adoconn.Open getConnectionString
'#

    Dim adoRst As New ADODB.Recordset
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
   End If
   adoRst.Close
    szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
     adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        txtBC.Tag = adoRst.Fields("BNC").Value
        txtBC.text = adoRst.Fields("BNN").Value
   End If
   adoRst.Close
   LoadCmbFinancialYear adoconn
   LoadFlxCashBook adoconn
   '#
   Call BankAndBalance(adoconn)

   adoconn.Close
   Set adoconn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Option1_Click()
    Frame1.Visible = True
    Frame2.Visible = False
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
'
'Private Sub txtSearchProperty_Change()
'   Dim sFilter_ As String
'
'   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
'                 "ORDER BY PropertyID;"
'   PopulatePropertyLookup sFilter_
'End Sub

Private Function BankAndBalance(adoconn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim rRow As Integer
   
  
         szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
 
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
   'Modified by anol 22 Feb 2015
            If txtClientList.text <> "Consolidated" Then
                    MsgBox "Please setup your Client Bank Accounts." & Chr(13) & _
                           "Please also check the nominal chart of account for the client."
             End If
             
   Else

                rRow = 1
                While Not adoRst.EOF
                    flxClient.row = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("BNC").Value
                    flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("BNN").Value
                    flxClient.TextMatrix(rRow, 3) = adoRst.Fields.Item("ID").Value
                    flxClient.TextMatrix(rRow, 4) = adoRst.Fields.Item("CLIENT_ID").Value
                    flxClient.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   'LoadAdoBank

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function


Public Sub LoadFlxCashBook(adoconn As ADODB.Connection)
   Dim szSQL As String, i As Integer, r As Integer, szaTemp() As String
   Dim adoRst As New ADODB.Recordset

'  Column Heading: Trans ID, Trans Type, Date, Ref, Details, Debit, Credit, Reconciled, Statement Date
'                    ^           ^         ^           ^       ^      ^          ^           ^
'   szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
'                  "Type as Type1, R.ExtRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
'           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT " & _
'           "WHERE (Type = 3 OR Type = 4 OR Type = 23) AND TT.TYPE_ID = type AND " & _
'                  "R.BankCode = '" & cboBC.Column(0) & "' " & _
'           "UNION " & _
'           "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
'                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
'                  "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
'                  "BP.BANK_AC AS ACC " & _
'           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT " & _
'           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
'                  "BANK_AC = '" & cboBC.Column(0) & "' AND BP.TransactionType = TT.TYPE_ID " & _
'           "UNION " & _
'           "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
'                  "P.Amount, P.Type AS Type1, P.ExtRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
'           "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
'           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
'                  "BankCode = '" & cboBC.Column(0) & "' AND P.Type = TT.TYPE_ID " & _
'           "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"
 szSQL = "SELECT SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, RDate, Details, Amount, " & _
                  "Type as Type1, R.EXTRef AS Rfn, R.ReconNow AS SDate, R.Reconciled, R.SageAccountNumber AS ACC " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND " & _
                  "R.BankCode = '" & txtBC.Tag & "' AND " & _
                  "U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT BP.TRAN_ID AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, BP.TRAN_DATE AS RDate, " & _
                  "BP.DESCRIPTION AS Details, (BP.NET_AMOUNT + BP.VAT) AS Amount, " & _
                  "BP.TransactionType AS Type1, BP.PROJ_REF AS Rfn, BP.ReconNow AS SDate, BP.Reconciled, " & _
                  "BP.BANK_AC AS ACC " & _
           "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
           "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                  "BP.BANK_AC = '" & txtBC.Tag & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                  "BP.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID " & _
           "UNION "
   szSQL = szSQL & _
           "SELECT P.SlNumber AS T_ID, MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS Type2, P.PDate AS RDate, Details, " & _
                  "P.Amount, P.Type AS Type1, P.EXTRef AS Rfn, P.ReconNow AS SDate, P.Reconciled, P.SageAccountNumber AS ACC " & _
           "FROM tlbPayment AS P, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
           "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND " & _
                  "P.BankCode = '" & txtBC.Tag & "' AND P.Type = TT.TYPE_ID AND " & _
                  "P.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "B.NominalCode = P.BankCode AND B.CLIENT_ID = P.ClientID " & _
           "ORDER BY RDate ASC, Type2 ASC, T_ID ASC;"

'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 1
   flxCashBook.Clear
   flxCashBook.Rows = 2
   flxCashBook.Cols = 11
   While Not adoRst.EOF
      flxCashBook.TextMatrix(i, 1) = adoRst.Fields.Item("RDate").Value                            'Date

      flxCashBook.TextMatrix(i, 2) = adoRst.Fields.Item("Type2").Value & _
                                     adoRst.Fields.Item("T_ID").Value                             'Type
      flxCashBook.TextMatrix(i, 3) = adoRst.Fields.Item("ACC").Value                              'Account
      flxCashBook.TextMatrix(i, 4) = adoRst.Fields.Item("Rfn").Value                              'Name
      flxCashBook.TextMatrix(i, 5) = adoRst.Fields.Item("Details").Value                          'Details
      If adoRst.Fields.Item("Type1").Value = "3" Or adoRst.Fields.Item("Type1").Value = "4" Or _
         adoRst.Fields.Item("Type1").Value = "12" Or adoRst.Fields.Item("Type1").Value = "24" Then
         flxCashBook.TextMatrix(i, 6) = Format(adoRst.Fields.Item("Amount").Value, "0.00")         'Debit
      Else
         flxCashBook.TextMatrix(i, 7) = Format(adoRst.Fields.Item("Amount").Value, "0.00")         'Credit
      End If
      flxCashBook.TextMatrix(i, 8) = IIf(Val(adoRst.Fields.Item("Amount").Value) - _
                                         IIf(adoRst.Fields.Item("Reconciled").Value < 0, _
                                         adoRst.Fields.Item("Reconciled").Value * (-1), _
                                         adoRst.Fields.Item("Reconciled").Value) = 0, "YES", _
                                         IIf(IsNull(adoRst.Fields.Item("Reconciled").Value), _
                                         "NO", "PART"))                                            'Reconcialid
      If Not IsNull(adoRst.Fields.Item("SDate").Value) Then
         szaTemp() = Split(adoRst.Fields.Item("SDate").Value, "#")
         flxCashBook.TextMatrix(i, 9) = IIf(szaTemp(1) = "Saved", "", szaTemp(0))                     'Statement Date
         flxCashBook.TextMatrix(i, 8) = IIf(szaTemp(1) = "Saved", "NO", flxCashBook.TextMatrix(i, 8))
      End If
       flxCashBook.TextMatrix(i, 10) = adoRst.Fields.Item("Type2").Value
      adoRst.MoveNext
      If Not adoRst.EOF Then flxCashBook.AddItem ""
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing

'   For i = 1 To flxCashBook.Rows - 1
'      flxCashBook.row = i
'      If flxCashBook.TextMatrix(i, 8) = "YES" Then
'         For r = 1 To flxCashBook.Cols - 1
'            flxCashBook.col = r
'            flxCashBook.CellBackColor = RGB(162, 185, 224)
'         Next r
'      End If
'   Next i
'   flxCashBook.row = 0
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
