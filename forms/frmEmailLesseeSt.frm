VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmailLesseeSt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lessee Statement"
   ClientHeight    =   10665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailLesseeSt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9540
      TabIndex        =   32
      Top             =   1080
      Width           =   795
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   2655
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   21
      Top             =   6345
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
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3525
         Left            =   45
         TabIndex        =   23
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6218
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
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   28
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6975
      TabIndex        =   20
      Top             =   1080
      Width           =   2550
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6030
      TabIndex        =   19
      Top             =   1080
      Width           =   930
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Print"
      Height          =   400
      Left            =   2040
      TabIndex        =   18
      Top             =   5640
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   680
      Left            =   7380
      TabIndex        =   16
      Top             =   135
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   3855
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while system processing email..."
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   15
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   400
      Left            =   6255
      TabIndex        =   13
      Top             =   5640
      Width           =   1200
   End
   Begin VB.CommandButton cmdSendStByEmail 
      Caption         =   "&Email Statement"
      Height          =   400
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   1800
   End
   Begin VB.TextBox txtTenantSearchUnitName 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   1065
      Width           =   1695
   End
   Begin VB.TextBox txtTenantSearchName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   1065
      Width           =   2595
   End
   Begin VB.TextBox txtTenantSearchID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   255
      TabIndex        =   2
      Top             =   1065
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLeaseList 
      Height          =   4140
      Left            =   120
      TabIndex        =   5
      Top             =   1410
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7303
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdClientList 
      Caption         =   ".."
      Height          =   300
      Left            =   6750
      TabIndex        =   0
      Top             =   135
      Width           =   300
   End
   Begin VB.CommandButton cmdProperty 
      Caption         =   ".."
      Height          =   300
      Left            =   6750
      TabIndex        =   1
      Top             =   495
      Width           =   300
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   945
      TabIndex        =   31
      Top             =   450
      Width           =   5805
      VariousPropertyBits=   746604571
      Size            =   "10239;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   945
      TabIndex        =   30
      Top             =   135
      Width           =   5805
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "10239;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Balance"
      Height          =   195
      Index           =   3
      Left            =   5820
      TabIndex        =   17
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   495
      Width           =   960
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee Id"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   9
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee Name"
      Height          =   195
      Index           =   1
      Left            =   1620
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Name"
      Height          =   195
      Index           =   2
      Left            =   4005
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Index           =   10
      Left            =   6945
      TabIndex        =   6
      Top             =   840
      Width           =   405
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   120
      Top             =   840
      Width           =   10155
   End
End
Attribute VB_Name = "frmEmailLesseeSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szaTenantBalance()     As String
Dim sTextBox  As String
Private Type SendDemandByEmail
   szLesseeID    As String
   szLesseeEmail As String
   colAtt        As Collection
   szClient      As String
   szLesseeName  As String
End Type
Private uLessee()   As SendDemandByEmail
Private iLes        As Integer
Private szLesseeList As String
Dim szSub As String, szBody As String
Private Sub LoadflxClient()
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
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL"
           flxClient.RowHeight(rRow) = 240
           flxClient.AddItem ""
           rRow = 2
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 240
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   
   adoconn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = "ALL"
           flxClient.TextMatrix(rRow, 2) = "ALL"
           flxClient.RowHeight(rRow) = 240
           flxClient.AddItem ""
           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadflxClient
    
   ' fraGrid.Enabled = False
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
End Sub

Private Sub cmdproperty_Click()
    sTextBox = "2"
     picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadPropertyList
    'fraGrid.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub cmdClose_Click()
   Unload Me
End Sub

'implementation of grid control on client and properties
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
Private Sub flxClient_Click()
            If sTextBox = "1" Then
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    txtPropertyName.Tag = "ALL"
                    txtPropertyName.text = "ALL"
                    txtPropertyName.SetFocus
                    
            ElseIf sTextBox = "2" Then
                    txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                    txtTenantSearchID.SetFocus
            End If
            picClient.Visible = False
            Dim adoconn As New ADODB.Connection
            adoconn.Open getConnectionString
            TenantAccountBalance adoconn
            LoadFlxLeaseList adoconn
            adoconn.Close
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
'End of implementation of grod control on client and properties by anol 22 Aug 2016
'  Create selected Lessees' List for email
Private Function CreateLesList4Email() As Boolean
   Dim i As Integer

   For i = 1 To flxLeaseList.Rows - 1
      If flxLeaseList.TextMatrix(i, 0) = "X" And _
            flxLeaseList.RowHeight(i) > 0 And _
            flxLeaseList.TextMatrix(i, 6) = "Yes" Then
         szLesseeList = szLesseeList & "'" & flxLeaseList.TextMatrix(i, 1) & "'" & ", "
      End If
   Next i
   If Len(szLesseeList) > 0 Then szLesseeList = Left(szLesseeList, Len(szLesseeList) - 2)

   For i = 1 To flxLeaseList.Rows - 1
      If flxLeaseList.TextMatrix(i, 0) = "X" And _
            flxLeaseList.RowHeight(i) > 0 And _
            flxLeaseList.TextMatrix(i, 6) = "No" Then
         CreateLesList4Email = True
         Exit Function
      End If
   Next i

   CreateLesList4Email = False
End Function

'  Create selected Lessees' List for printing
Private Sub CreateLesList4Print()
   Dim i As Integer

   For i = 1 To flxLeaseList.Rows - 1
      If flxLeaseList.TextMatrix(i, 0) = "X" And flxLeaseList.RowHeight(i) > 0 Then
         szLesseeList = szLesseeList & "'" & flxLeaseList.TextMatrix(i, 1) & "'" & ", "
      End If
   Next i
   If Len(szLesseeList) > 0 Then szLesseeList = Left(szLesseeList, Len(szLesseeList) - 2)
End Sub

Private Sub cmdGenerate_Click()
   Call ChangeReportODBC
   szLesseeList = ""
   CreateLesList4Print

   If Len(szLesseeList) = 0 Then Exit Sub

   Dim adoconn As New ADODB.Connection
   Dim szSQL   As String
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   adoconn.Open getConnectionString

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = '' " & _
           "WHERE  spare2 = 'Y';"
   adoconn.Execute szSQL

   szSQL = "UPDATE Tenants " & _
           "SET    spare2 = 'Y' " & _
           "WHERE  SageAccountNumber IN (" & szLesseeList & ");"
   adoconn.Execute szSQL

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False  ''Only outstanding statement
   Report.ParameterFields(2).AddCurrentValue False ' 'specific lessee
   Report.ParameterFields(3).AddCurrentValue "1"    'Lessee ID

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub cmdRefresh_Click()
   Dim adoconn As New ADODB.Connection

   iLes = 0

   adoconn.Open getConnectionString

'  Load all tenants with balance in gridTenantLookup
   TenantAccountBalance adoconn
   LoadFlxLeaseList adoconn

   iLes = 0
   ReDim uLessee(0) As SendDemandByEmail

   adoconn.Close
   Set adoconn = Nothing
End Sub
Private Sub totalSeleactionInGrid()
    Dim i As Integer
    Dim K As Integer
    For i = 0 To flxLeaseList.Rows - 1
        If flxLeaseList.TextMatrix(i, 0) = "X" Then
            K = K + 1
        End If
    Next
End Sub
Private Function IntClearSearchFilter() As Boolean
    'if it clear some value then this function shall return true
    Dim i As Integer
   
    
    For i = flxLeaseList.Rows - 1 To 1 Step -1
         If flxLeaseList.RowHeight(i) = 0 Then
               IntClearSearchFilter = True
               flxLeaseList.RowHeight(i) = 240
               flxLeaseList.row = i
         End If
    Next i

End Function
Private Sub cmdSendStByEmail_Click()
   Dim bLesWOSetup As Boolean
   Call ChangeReportODBC
   If iLes = 0 Then
      ShowMsgInTaskBar "No Lessee has been selected.", "Y", "N"
      Exit Sub
   End If
   'issue 496 B
   'Also when user clears search entry all items selected in a selection are being cleared. This should not happen. If the user does not clear the search value in the search grid and clicks Email Statement, the system will clear the search value and display the full list of selected items, The user will now be able to click Email statement again to send the emails.
   If IntClearSearchFilter Then
        txtSearchClientName.text = ""
        txtTenantSearchID.text = ""
        txtTenantSearchUnitName.text = ""
        MsgBox "Search Filter has been cleared down. Now you are ready to Email statements"
        Exit Sub
   End If
   szLesseeList = ""
   bLesWOSetup = CreateLesList4Email
   If iLes > 0 And Len(szLesseeList) = 0 Then
      ShowMsgInTaskBar "There is no valid Lessee selected to which to send a statement", "Y", "N"
      Exit Sub
   End If

   If bLesWOSetup Then
      If MsgBox("WARNING: You have selected some lessees who are not setup to receive statements by email." & Chr(13) & _
                "Do you wish to continue?", vbQuestion + vbYesNo, "Warning: Send Statement") = vbNo Then Exit Sub
   End If
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

   Dim szPath        As String
   Dim i             As Integer
   Dim bEmailResult  As Boolean

   Dim adoconn As New ADODB.Connection

   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoconn.Open getConnectionString

   adoconn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = '' " & _
                   "WHERE  spare2 = 'Y';"

   adoconn.Execute "UPDATE Tenants " & _
                   "SET    spare2 = 'Y' " & _
                   "WHERE  SageAccountNumber IN (" & szLesseeList & ");"
                   
 Dim adoRstEmailDetails  As New ADODB.Recordset
 Dim szSQL As String
'-----------------------------------------------------------------------------------------------
'  Get the subject and body of the email from template
   szSQL = "SELECT * FROM Template WHERE TemplateName = 'Statement Email Template';"
   adoRstEmailDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRstEmailDetails.EOF Then
      szSub = adoRstEmailDetails.Fields.Item("Description").Value
      szBody = adoRstEmailDetails.Fields.Item("Body").Value
   End If
   adoRstEmailDetails.Close

'-----------------------------------------------------------------------------------------------------------


   adoconn.Close
   Set adoconn = Nothing

   For i = 0 To UBound(uLessee)
      If uLessee(i).szLesseeID <> "" And InStr(szLesseeList, uLessee(i).szLesseeID) > 0 Then
         '  Create the pdf file name of the statement
         szPath = uLessee(i).szLesseeID & "_" & UniqueID() & ".pdf"

         CreatePDF_Statement uLessee(i).szLesseeID, DB_PATH & "\AllStuff\Temp\" & szPath

         '  Attaching the statement to the email
         SaveAttachment DB_PATH & "\AllStuff\Temp\" & szPath, i

         EmailDelay 20
      End If
   Next i

   bEmailResult = SendDemandByE_Mail

   fmeLoading.Visible = False
   If iLes > 0 And bEmailResult Then
      MsgBox "Email sent.", vbInformation + vbOKOnly, "Lessee Statement"
   Else
      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Lessee Statement"
   End If
End Sub

Private Function SendDemandByE_Mail() As Boolean
   Dim i As Integer
   For i = 0 To iLes - 1
      If uLessee(i).szLesseeEmail <> "" And InStr(szLesseeList, uLessee(i).szLesseeID) > 0 Then
            If szBody = "" Then 'This condition has been added by anol 2020-08-04 issue 863
                szSub = "Account Statement"
                szBody = "Please find attachment your account statement."
            Else
                szSub = Replace(szSub, "<CLIENT NAME>", uLessee(i).szClient)
                szBody = Replace(szBody, "<CLIENT NAME>", uLessee(i).szClient)
                
                szSub = Replace(szSub, "<LESSEE NAME>", uLessee(i).szLesseeName)
                szBody = Replace(szBody, "<LESSEE NAME>", uLessee(i).szLesseeName)
            End If
            SendDemandByE_Mail = SendEmail(szFromEmail, Trim(uLessee(i).szLesseeEmail), _
                                            szSub, _
                                            szBody, , , _
                                            uLessee(i).colAtt, uLessee(i).szLesseeID, "Account Statement")
      End If
   Next i
End Function

Private Sub flxLeaseList_Click()
   Dim iSel As Integer
   Dim i    As Integer
   Dim K As Integer
   iSel = SelectFlxGridRow(0, flxLeaseList, flxLeaseList.RowSel)
Rem by anol 20161103
'   If iSel = 1 Then
'      ReceiverEmailList flxLeaseList.TextMatrix(flxLeaseList.row, 1), flxLeaseList.TextMatrix(flxLeaseList.row, 5), "A"  'A --> Add +
'   Else
'      ReceiverEmailList flxLeaseList.TextMatrix(flxLeaseList.row, 1), flxLeaseList.TextMatrix(flxLeaseList.row, 5), "S" 'S --> Subtract -
'   End If
    For i = 0 To flxLeaseList.Rows - 1
        If flxLeaseList.TextMatrix(i, 0) = "X" Then
            K = K + 1
        End If
    Next
    iLes = K
    Call ReceiverEmailList2(flxLeaseList, K)
End Sub
Private Function ReceiverEmailList2(flxLeaseList As MSHFlexGrid, K As Integer) As Integer
    '   ReceiverEmailList function was giving error out of bounds error so I have written this function which will work fine anol20161103
        Dim i As Integer
        Dim j As Integer
        If K = 0 Then
            ReDim uLessee(0) As SendDemandByEmail
        Else
            ReDim uLessee(K) As SendDemandByEmail
            For i = 0 To flxLeaseList.Rows - 1
                If flxLeaseList.TextMatrix(i, 0) = "X" Then
                    uLessee(j).szLesseeID = flxLeaseList.TextMatrix(i, 1)
                    uLessee(j).szLesseeName = flxLeaseList.TextMatrix(i, 2)
                    uLessee(j).szLesseeEmail = flxLeaseList.TextMatrix(i, 5)
                    uLessee(j).szClient = flxLeaseList.TextMatrix(i, 7)
                    j = j + 1
                End If
            Next i
        End If
     
End Function
Private Function ReceiverEmailList(szLessee As String, szEmail As String, szOption As String) As Integer
   Dim i As Integer

   ReceiverEmailList = 0
   If iLes = 0 Then
      iLes = 1
      ReDim uLessee(0) As SendDemandByEmail
      uLessee(0).szLesseeID = szLessee
      uLessee(0).szLesseeEmail = szEmail
      Exit Function
   End If
   
   If szOption = "S" Then GoTo Subtract

   ReDim Preserve uLessee(iLes) As SendDemandByEmail

   uLessee(iLes).szLesseeID = szLessee
   uLessee(iLes).szLesseeEmail = szEmail
   ReceiverEmailList = iLes
   iLes = iLes + 1
   Exit Function

Subtract:
   For i = 0 To UBound(uLessee)
      If uLessee(i).szLesseeID = szLessee Then Exit For
   Next i
   uLessee(i).szLesseeID = ""
   uLessee(i).szLesseeEmail = ""
   iLes = iLes - 1
End Function

Private Sub Form_Load()
   Dim adoconn As New ADODB.Connection

   iLes = 0

'   Me.Width = 9525 ' 7650
   Me.Height = 6595
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR

   adoconn.Open getConnectionString

   'PrepareList adoConn, cboRptClientList, cboRptPropertyList
    
    '#
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
        txtPropertyName.Tag = "ALL"
        txtPropertyName.text = "ALL"
   End If
   adoRst.Close
   'From refresh button
   iLes = 0
'  Load all tenants with balance in gridTenantLookup
   TenantAccountBalance adoconn
   LoadFlxLeaseList adoconn
   iLes = 0
   ReDim uLessee(0) As SendDemandByEmail
   
   
'  Load all tenants with balance in gridTenantLookup
'   TenantAccountBalance adoConn
'   LoadFlxLeaseList adoConn
  'End From refresh button
 
 
   adoconn.Close
   Set adoconn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareList(adoconn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRst.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

'   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub ConfigureFlxLeaseList()
   Dim szHeader As String

'   flxLeaseList.Width = 6735
'   Label20(0).Caption = "Lessee Id"
'   Label20(1).Caption = "Lessee Name"
'   Label20(0).Width = 690
   Label20(0).Left = 120
'   Label20(1).Width = 930
   Label20(1).Left = 1560
  ' txtTenantSearchName.Left = Label20(1).Left
'   Label20(2).Visible = True
'   txtTenantSearchUnitName.Visible = True
'   Shape4(6).Width = 6480

   flxLeaseList.Clear
   flxLeaseList.Cols = 8
   flxLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Lessee ID|<Lessee Name|<Lessee Name|>Ac Balance|Email|EmailSetup"
   flxLeaseList.FormatString = szHeader$
   flxLeaseList.ColWidth(0) = 200 'Label20(0).Left - flxLeaseList.Left       '0        Solid column
   flxLeaseList.ColWidth(1) = Label20(1).Left - Label20(0).Left - 20    '1400    'Tenant ID
   flxLeaseList.ColWidth(2) = Label20(2).Left - Label20(1).Left - 20             'Tenant Name
   flxLeaseList.ColWidth(3) = Label20(3).Left - Label20(2).Left - 20             'Unit Name
   flxLeaseList.ColWidth(4) = Label20(10).Left - Label20(3).Left - 20             'Ac Balance
   flxLeaseList.ColWidth(5) = 2500                                                  'Email
   flxLeaseList.ColAlignment(5) = vbLeftJustify
'   flxLeaseList.ColWidth(6) = 0                                                  'Email Setup
   flxLeaseList.ColWidth(6) = 700 'flxLeaseList.Left + flxLeaseList.Width - Label20(10).Left - 300 'Email Setup
   flxLeaseList.ColWidth(7) = 0 'client ID

   flxLeaseList.Rows = 2
End Sub

Public Function PopulateTenantLookup(adoconn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery_, adoconn, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   While Not adoRst.EOF
     ' If Not IsNull(adoRst!EMail) Then
         flxLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
         flxLeaseList.TextMatrix(iRow, 2) = adoRst!Name
         flxLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber
         'Modified by anol 19 Oct 2015
        ' flxLeaseList.TextMatrix(iRow, 5) = adoRst!EMail
         flxLeaseList.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!EMail), "No Email address", adoRst!EMail)
         flxLeaseList.TextMatrix(iRow, 6) = IIf(adoRst!EmailSt, "Yes", "No")
         flxLeaseList.TextMatrix(iRow, 7) = adoRst!clientID

         iRow = iRow + 1
         adoRst.MoveNext

         If Not adoRst.EOF Then flxLeaseList.AddItem ""
'      Else
'         adoRst.MoveNext
'      End If
   Wend
   adoRst.Close
   Set adoRst = Nothing
'MsgBox flxLeaseList.TextMatrix(iRow - 1, 6)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub LoadFlxLeaseList(adoconn As ADODB.Connection)
  ' Me.MousePointer = vbHourglass

   Dim szSQL As String

   ConfigureFlxLeaseList
   If txtClientList.Tag = "ALL" And txtPropertyName.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt,Property.ClientID " & _
              "From   Tenants, LeaseDetails, Units, Property  " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber " & _
                  "ORDER BY Tenants.SageAccountNumber;"
   End If
'Debug.Print szSQL
   If txtClientList.Tag <> "ALL" And txtPropertyName.Tag = "ALL" Then
   'Fixed by anol 20 Oct 2015
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt,Property.ClientID " & _
              "From   Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "Property.ClientID = '" & txtClientList.Tag & "' order by Tenants.SageAccountNumber"
'     szSQL = "SELECT T.SageAccountNumber, T.NAME, IQ.UnitNumber, " & _
'                  "IIF(T.InvoiceTo = 'H', T.Email1, T.Email2) AS Email, " & _
'                  "T.EmailSt " & _
'              "FROM Tenants AS T LEFT JOIN " & _
'                   "[" & _
'                   "SELECT U.UnitName, L.SageAccountNumber, " & _
'                          "P.PropertyID, P.ClientID, U.UnitNumber " & _
'                   "FROM Units AS U, LeaseDetails AS L, " & _
'                        "Property AS P " & _
'                   "WHERE U.UnitNumber = L.UnitNumber AND " & _
'                      "L.Status = TRUE AND U.PropertyID = P.PropertyID "
'
'      szSQL = szSQL + "]. AS IQ ON T.SageAccountNumber = IQ.SageAccountNumber " & _
'                 "WHERE ((T.Comments) IS NULL OR T.Comments = '') AND OCCUPIDE_ = FALSE "
'            If txtClientList.tag <> "ALL" Then
'                If cboRptPropertyList.ListIndex > -1 Then 'if and else condition added by anol 17 Sep 2015 issue 571 note 1174
'                        If txtPropertyName.tag <> "ALL" Then
'                           szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyName.tag & "' AND "
'                        Else
'                           szSQL = szSQL + "AND "
'                        End If
'                Else
'                       cboRptPropertyList.ListIndex = 0
'                       szSQL = szSQL + "AND "
'                End If
'                 szSQL = szSQL + "IQ.ClientID = '" & txtClientList.tag & "' "
'              Else
'                 If txtPropertyName.tag <> "ALL" Then
'                    szSQL = szSQL + "AND IQ.PropertyID = '" & txtPropertyName.tag & "'"
'                 End If
'              End If
'              szSQL = szSQL + " order by T.SageAccountNumber"
   End If

   If txtClientList.Tag = "ALL" And txtPropertyName.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt,Property.ClientID " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                   "Units.PropertyID = Property.PropertyID AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = '" & txtPropertyName.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag <> "ALL" And txtPropertyName.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt,Property.ClientID " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "Property.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "Units.PropertyID = '" & txtPropertyName.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If
'Debug.Print szSQL
   PopulateTenantLookup adoconn, szSQL

   UpdateBalance

   txtTenantSearchID.text = ""
   txtTenantSearchName.text = ""
   txtTenantSearchUnitName.text = ""

   'Me.MousePointer = vbArrow
End Sub

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxLeaseList.Rows - 1
      For j = 0 To UBound(szaTenantBalance, 2) - 1
         If flxLeaseList.TextMatrix(i, 1) = szaTenantBalance(0, j) Then
            flxLeaseList.TextMatrix(i, 4) = Format(szaTenantBalance(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaTenantBalance, 2) Then flxLeaseList.TextMatrix(i, 4) = "0.00"
   Next i
End Sub

'  Get all lessees' Account History in an array
Private Sub TenantAccountBalance(adoconn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT SageAccountNumber  " & _
             "From tlbReceipt " & _
             "GROUP BY SageAccountNumber" & _
            ");"
'Debug.Print szSQL
   adoRptDr.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRptDr.EOF Then
      adoRptDr.Close
      Set adoRptDr = Nothing
      Exit Sub
   End If

   ReDim szaTenantBalance(1, adoRptDr.Fields.Item(0).Value) As String
   adoRptDr.Close

   szSQL = "SELECT Rpt.SageAccountNumber, SUM(Rpt.Amount) AS Dr " & _
           "FROM tlbReceipt AS Rpt " & _
           "WHERE Rpt.Type = 1 OR Rpt.Type = 23 " & _
           "GROUP BY Rpt.SageAccountNumber;"

   adoRptDr.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoRptDr.EOF
      szaTenantBalance(0, iIndex) = adoRptDr.Fields.Item("SageAccountNumber").Value
      szaTenantBalance(1, iIndex) = adoRptDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoRptDr.MoveNext
   Wend

   adoRptDr.Close

   szSQL = "SELECT tlbReceipt.SageAccountNumber, SUM(tlbReceipt.Amount) AS Cr " & _
           "FROM tlbReceipt " & _
           "WHERE tlbReceipt.Type <> 1 AND tlbReceipt.Type <> 23 " & _
           "GROUP BY tlbReceipt.SageAccountNumber;"

   adoRptCr.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   While Not adoRptCr.EOF
      For i = 0 To iIndex - 1
         If szaTenantBalance(0, i) = adoRptCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaTenantBalance(1, i) = szaTenantBalance(1, i) - Val(adoRptCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaTenantBalance(0, iIndex) = adoRptCr.Fields.Item("SageAccountNumber").Value
         szaTenantBalance(1, iIndex) = adoRptCr.Fields.Item("Cr").Value
      End If
      adoRptCr.MoveNext
   Wend

   adoRptCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmDemands3.Enabled = True
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub CreatePDF_Statement(szLessee As String, szFileName As String)
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeStatement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue False          'Only outstanding statement
   'Report.ParameterFields(2).AddCurrentValue False          'specific lessee
   'Modified by anol 11 April 2016
   Report.ParameterFields(2).AddCurrentValue True          'specific lessee
   Report.ParameterFields(3).AddCurrentValue szLessee       'Lessee id

'   Transfer report into PDF file
   Report.ExportOptions.DiskFileName = szFileName
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing

'  ##################################################################################################
'
'  It is blocked on 19/12/2011.
'
'  ##################################################################################################

'   Dim szSQL      As String
'   Dim dLesBal    As Double
'   Dim adoRst     As New ADODB.Recordset
'   Dim reportApp  As New CRAXDRT.Application
'   Dim Report     As CRAXDRT.Report
'
'   szSQL = "SELECT T.Name, U.UnitName, P.PropertyName, C.ClientName " & _
'           "FROM Tenants AS T,  LeaseDetails AS L, Units AS U, Property AS P, Client AS C " & _
'           "WHERE Status = TRUE AND T.SageAccountNumber = '" & szLessee & "' AND " & _
'               "T.SageAccountNumber = L.SageAccountNumber AND " & _
'               "L.UnitNumber = U.UnitNumber AND " & _
'               "U.PropertyID = P.PropertyID AND " & _
'               "P.ClientID = C.ClientID;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      Set adoRst = Nothing
'      Exit Sub
'   End If
'
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   If Report.HasSavedData Then Report.DiscardSavedData
'
'   Report.ParameterFields(1).AddCurrentValue szLessee
'   Report.ParameterFields(2).AddCurrentValue adoRst.Fields.Item("Name").Value
'   Report.ParameterFields(3).AddCurrentValue adoRst.Fields.Item("ClientName").Value
'   Report.ParameterFields(4).AddCurrentValue adoRst.Fields.Item("PropertyName").Value
'   Report.ParameterFields(5).AddCurrentValue adoRst.Fields.Item("UnitName").Value
'
'   dLesBal = CDbl(LesseeAccountBalance(adoConn, szLessee))
'   Report.ParameterFields(6).AddCurrentValue dLesBal
'
'   adoRst.Close
'   Set adoRst = Nothing
'
'   'Transfer report into PDF file
'   Report.ExportOptions.DiskFileName = szFileName
'   Report.ExportOptions.DestinationType = crEDTDiskFile
'   Report.ExportOptions.FormatType = crEFTPortableDocFormat
'   Report.ExportOptions.PDFExportAllPages = True
'   Report.Export False
'   Set Report = Nothing
End Sub

Private Sub SaveAttachment(szFile As String, i As Integer)
   Set uLessee(i).colAtt = New Collection
   uLessee(i).colAtt.Add szFile
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

Private Sub txtTenantSearchID_Change()
'issue 496
'(b) Also when user clears search entry all items selected in a selection are being cleared. This should not happen
'  Call filterlist
'  UpdateBalance
   Dim i As Integer
   If Len(txtTenantSearchID.text) > 0 Then
        txtSearchClientName.text = ""
   End If
   For i = flxLeaseList.Rows - 1 To 1 Step -1
        flxLeaseList.RowHeight(i) = 240
        If InStr(1, UCase(flxLeaseList.TextMatrix(i, 1)), UCase(txtTenantSearchID.text), vbTextCompare) = 0 Then
              flxLeaseList.RowHeight(i) = 0
        End If
        If flxLeaseList.RowHeight(i) = 240 Then
              flxLeaseList.row = i
        End If
   Next i
End Sub
Private Sub filterlist()
   Dim Filter As String
   'Wild card search has been implemented by anol
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   If Len(txtTenantSearchID.text) > 0 Then
      txtTenantSearchName.text = ""
      txtTenantSearchUnitName.text = ""
      Filter = " SageAccountNumber LIKE '%" + UCase(txtTenantSearchID.text) + "*'"
      
   End If
   
   If Len(txtTenantSearchName.text) > 0 Then
      txtTenantSearchID.text = ""
      txtTenantSearchUnitName.text = ""
      Filter = " CompanyName LIKE '%" + UCase(txtTenantSearchName.text) + "*'"
   End If

   If Len(txtTenantSearchUnitName.text) > 0 Then
      txtTenantSearchID.text = ""
      txtTenantSearchName.text = ""
      Filter = " UnitName LIKE '%" + UCase(txtTenantSearchUnitName.text) + "*'"
   End If
    
   Dim szSQL As String

   ConfigureFlxLeaseList
   If txtClientList.Tag = "ALL" And txtPropertyName.Tag = "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt " & _
              "From   Tenants, LeaseDetails " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.Status = True " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag <> "ALL" And txtPropertyName.Tag = "ALL" Then
   
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt " & _
              "From   Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "Property.ClientID = '" & txtClientList.Tag & "' order by Tenants.SageAccountNumber"

   End If

   If txtClientList.Tag = "ALL" And txtPropertyName.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt " & _
              "From Tenants, LeaseDetails, Units " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = '" & txtPropertyName.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If

   If txtClientList.Tag <> "ALL" And txtPropertyName.Tag <> "ALL" Then
      szSQL = "SELECT Tenants.SageAccountNumber, Name, LeaseDetails.UnitNumber, " & _
                  "IIF(Tenants.InvoiceTo = 'H', Tenants.Email1, Tenants.Email2) AS Email, " & _
                  "Tenants.EmailSt " & _
              "From Tenants, LeaseDetails, Units, Property " & _
              "WHERE ((Tenants.Comments) IS NULL OR Tenants.Comments='') AND " & _
                  "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber AND " & _
                  "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True AND " & _
                  "Units.PropertyID = Property.PropertyID AND " & _
                  "Property.ClientID = '" & txtClientList.Tag & "' AND " & _
                  "Units.PropertyID = '" & txtPropertyName.Tag & "' " & _
             "ORDER BY Tenants.SageAccountNumber;"
   End If
   
   Dim adoRst As New ADODB.Recordset

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   adoRst.Filter = Filter
   Dim iRow As Integer
   iRow = 1

   While Not adoRst.EOF
     ' If Not IsNull(adoRst!EMail) Then
         flxLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
         flxLeaseList.TextMatrix(iRow, 2) = adoRst!Name
         flxLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber
         'Modified by anol 19 Oct 2015
        ' flxLeaseList.TextMatrix(iRow, 5) = adoRst!EMail
         flxLeaseList.TextMatrix(iRow, 5) = IIf(IsNull(adoRst!EMail), "No Email address", adoRst!EMail)
         flxLeaseList.TextMatrix(iRow, 6) = IIf(adoRst!EmailSt, "Yes", "No")

         iRow = iRow + 1
         adoRst.MoveNext

         If Not adoRst.EOF Then flxLeaseList.AddItem ""
'      Else
'         adoRst.MoveNext
'      End If
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub txtTenantSearchName_Change()
           'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtTenantSearchName.text) > 0 Then
        txtTenantSearchID.text = ""
   End If

   For i = flxLeaseList.Rows - 1 To 1 Step -1
        flxLeaseList.RowHeight(i) = 240
        If InStr(1, UCase(flxLeaseList.TextMatrix(i, 2)), UCase(txtTenantSearchName.text), vbTextCompare) = 0 Then
            flxLeaseList.RowHeight(i) = 0
        End If
        If flxLeaseList.RowHeight(i) = 240 Then
            flxLeaseList.row = i
        End If
   Next i
   
End Sub

Private Sub txtTenantSearchUnitName_Change()
    Dim i As Integer

   If Len(txtTenantSearchName.text) > 0 Then
        txtTenantSearchID.text = ""
   End If

   For i = flxLeaseList.Rows - 1 To 1 Step -1
        flxLeaseList.RowHeight(i) = 240
        If InStr(1, UCase(flxLeaseList.TextMatrix(i, 3)), UCase(txtTenantSearchName.text), vbTextCompare) = 0 Then
            flxLeaseList.RowHeight(i) = 0
        End If
        If flxLeaseList.RowHeight(i) = 240 Then
            flxLeaseList.row = i
        End If
   Next i
   
End Sub
