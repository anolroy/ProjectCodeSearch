VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFinancialYear 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Years"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "frmFinancialYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdSetAsCurrent 
      Caption         =   "Set As Current"
      Height          =   375
      Left            =   3195
      TabIndex        =   30
      Top             =   6570
      Width           =   1485
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   225
      ScaleHeight     =   6135
      ScaleWidth      =   6255
      TabIndex        =   20
      Top             =   7605
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
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   5415
         Left            =   45
         TabIndex        =   22
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6510
      TabIndex        =   5
      Top             =   6570
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1575
      TabIndex        =   3
      Top             =   6570
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5175
      TabIndex        =   4
      Top             =   6570
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add"
      Height          =   375
      Left            =   255
      TabIndex        =   2
      Top             =   6570
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9F9F9&
      Height          =   990
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9060
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
         Left            =   5925
         TabIndex        =   0
         Top             =   270
         Width           =   345
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   585
         TabIndex        =   29
         Top             =   270
         Width           =   5670
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "10001;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblPeriods 
         BackStyle       =   0  'Transparent
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
         Left            =   6630
         TabIndex        =   12
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Periods:"
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
         Index           =   4
         Left            =   5265
         TabIndex        =   11
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
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
         Left            =   1080
         TabIndex        =   10
         Top             =   630
         Width           =   2745
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description: "
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
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lblYear 
         BackStyle       =   0  'Transparent
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
         Left            =   6855
         TabIndex        =   8
         Top             =   270
         Width           =   1710
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Left            =   6435
         TabIndex        =   7
         Top             =   270
         Width           =   345
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFinancialYear 
      Height          =   5085
      Left            =   90
      TabIndex        =   1
      Top             =   1245
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8969
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
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Index           =   5
      Left            =   8100
      TabIndex        =   19
      Top             =   1005
      Width           =   450
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
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
      Left            =   240
      TabIndex        =   18
      Top             =   1005
      Width           =   435
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Index           =   3
      Left            =   5400
      TabIndex        =   16
      Top             =   1005
      Width           =   840
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   3360
      TabIndex        =   15
      Top             =   1005
      Width           =   840
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
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
      Index           =   4
      Left            =   6495
      TabIndex        =   13
      Top             =   1005
      Width           =   765
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
      Top             =   1005
      Width           =   9015
   End
End
Attribute VB_Name = "frmFinancialYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClientList_Click()
     picClient.Left = 269.029
    picClient.Top = 155.299
    'sTextBox = "1"
    LoadflxClient
'    Frame2.Enabled = False
'    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub



Private Sub cmdSetAsCurrent_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If flxFinancialYear.row > 0 Then
            adoConn.Execute "Update financialYear set Setascurrent=false where clientID='" & txtClientList.Tag & "'"
            adoConn.Execute "Update financialYear set Setascurrent=true where clientID='" & txtClientList.Tag & "' and FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "'"
            LoadFlxFinancialYear adoConn
    End If
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub flxClient_Click()
     txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
     txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
     RefreshGrid
     picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
            RefreshGrid
            picClient.Visible = False
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub



'Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 13 Then
'        cmdProperty.SetFocus
'    End If
'End Sub

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
          
          
                 cmdClientList.SetFocus

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
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
   
    cmdClientList.SetFocus
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

Public Sub RefreshGrid()
   Dim iRow As Integer
   Dim conConn As New ADODB.Connection

   ConfigureFlxFinancialYear

   conConn.Open getConnectionString

   LoadFlxFinancialYear conConn

   conConn.Close
   Set conConn = Nothing

   For iRow = 1 To flxFinancialYear.Rows - 1
      If flxFinancialYear.TextMatrix(iRow, 1) = txtClientList.Tag Then
         flxFinancialYear.RowHeight(iRow) = 240
      Else
         flxFinancialYear.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub



Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
        If flxFinancialYear.TextMatrix(flxFinancialYear.row, 7) = "Close" Then
            MsgBox "You cannot delete a closed financial year", vbInformation, "Closed financial year!!"
    '        If MsgBox("Closing the financial year will also close all the open periods in the financial year." & Chr(13) & _
    '                "Do you wish to continue?", vbYesNo + vbQuestion, "Close opened Financial Year") = vbNo Then Exit Sub
    '
    '      Dim conConn As New ADODB.Connection
    '      conConn.Open getConnectionString
    '
    '      'Resolved By BOSL. Issue: 0000484
    '      'Modified by Asif. 26-10-2014
    '      CloseFinancialYear flxFinancialYear.TextMatrix(flxFinancialYear.row, 0), conConn
    '
    '      ShowMsgInTaskBar "Financial Year is closed", "Y", "P"
    '      ConfigureFlxFinancialYear
    '      LoadFlxFinancialYear conConn
    '
    '      conConn.Close
    '      Set conConn = Nothing
            Exit Sub
     End If
 'check if there is next FY exists'
      Dim conConn As New ADODB.Connection
      conConn.Open getConnectionString
      Dim rsFY As New ADODB.Recordset
      Dim szSQL As String
      szSQL = "SELECT FYrID, ClientID, FinancialYear, FY_Description, " & _
              "FY_StDate, FY_EndDate, PeriodsCount, Status " & _
              "FROM   FinancialYear where   clientID='" & txtClientList.Tag & "' AND FY_StDate > (select FY_StDate from FinancialYear " & _
              " where FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "' and clientID='" & txtClientList.Tag & "')"
       rsFY.Open szSQL, conConn, adOpenKeyset, adLockReadOnly
       If Not rsFY.EOF Then
            MsgBox "There is a subsequent financial year.You need to delete that.", vbInformation, "Subsequent financial year!"
            Exit Sub
       End If
       rsFY.Close
         '  It should be possible to delete a financial year even if it has periods created against it.
        szSQL = "SELECT FYrID,  status " & _
              "FROM   Periods where  FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "'"
       rsFY.Open szSQL, conConn, adOpenKeyset, adLockReadOnly
       If Not rsFY.EOF Then
            MsgBox "It has periods created against it.You need to delete that.", vbInformation, "It has periods created against it!"
            Exit Sub
       End If
       rsFY.Close
       '1)  It should NOT be possible to delete a financial year if it has transactions created against it.
       szSQL = "SELECT TRANSACTION_DATE,  clientID " & _
              "FROM   NLPOSTING where   clientID='" & txtClientList.Tag & "'  AND TRANSACTION_DATE >=#" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 4) & "#" & _
              "AND TRANSACTION_DATE <=#" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 5) & "#"
       rsFY.Open szSQL, conConn, adOpenKeyset, adLockReadOnly
       If Not rsFY.EOF Then
            MsgBox "Not possible to delete a financial year, it has transactions against it.", vbInformation, "Not possible to delete a financial year!"
            Exit Sub
       End If
       rsFY.Close
        If MsgBox("Are you sure to delete this financial year?", vbYesNo, "Sure to delete ?") = vbYes Then
            conConn.Execute "DELETE from FinancialYear where FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "'"
            MsgBox "Finanacial year has been deleted successfully", vbInformation, "DELETED"
            ConfigureFlxFinancialYear
            LoadFlxFinancialYear conConn
        End If
        conConn.Close
End Sub

Private Sub cmdEdit_Click()
   If txtClientList.text = "" Then
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      cmdClientList.SetFocus
      Exit Sub
   End If

   If flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) = "" Or flxFinancialYear.row = 0 Then
      ShowMsgInTaskBar "Please select a financial year", "Y", "N"
      flxFinancialYear.SetFocus
      Exit Sub
   End If

   If flxFinancialYear.TextMatrix(flxFinancialYear.row, 7) = "False" Then
      ShowMsgInTaskBar "Financial year is closed. Edit not allowed", "Y", "N"
      Exit Sub
   End If

   Load frmFinancialYearCreate
   frmFinancialYearCreate.FinancialYearID = flxFinancialYear.TextMatrix(flxFinancialYear.row, 0)
   'Me.Hide

   frmFinancialYearCreate.txtStDate.Locked = IIf(flxFinancialYear.row = 1, False, True)

   frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Edit"
  ' frmFinancialYearCreate.cmdSave.Visible = False
   frmFinancialYearCreate.optPreDefined.Enabled = False
   frmFinancialYearCreate.LoadFlxPeriods
   frmFinancialYearCreate.optCustome.Value = True
   frmFinancialYearCreate.Show
End Sub

Private Sub flxFinancialYear_Click()
    If flxFinancialYear.TextMatrix(flxFinancialYear.row, 1) = "" Then
       cmdClientList.SetFocus
       Exit Sub
    End If
    lblDescription.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 3)
    lblYear.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 2)
    'Financial Years displaying incorrect months
    'Fixed by anol 20 07 2016
    Dim con As New ADODB.Connection
    con.Open getConnectionString
    Dim rsCheck As New ADODB.Recordset
    Dim rRow As Integer
    rsCheck.Open "Select PeriodID from Periods where FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "'", con, adOpenKeyset, adLockReadOnly
    rRow = 0
    If Not rsCheck.EOF Then
        rRow = rsCheck.RecordCount
    End If
    con.Execute "Update FinancialYear set PeriodsCount=" & rRow & " where FYrID='" & flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) & "'"
    lblPeriods.Caption = rRow 'flxFinancialYear.TextMatrix(flxFinancialYear.row, 6)
    rsCheck.Close
    con.Close
    txtClientList.Tag = flxFinancialYear.TextMatrix(flxFinancialYear.row, 1)
End Sub

Private Sub flxFinancialYear_DblClick()
   cmdEdit_Click
End Sub

Private Sub flxFinancialYear_RowColChange()
   If flxFinancialYear.TextMatrix(flxFinancialYear.row, 1) = "" Then
      cmdClientList.SetFocus
      Exit Sub
   End If

   lblDescription.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 3)
   lblYear.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 2)
   'lblPeriods.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 6)
   txtClientList.Tag = flxFinancialYear.TextMatrix(flxFinancialYear.row, 1)
End Sub

Private Sub Form_Load()
   Me.Height = 7560
   Me.Width = 9345
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR

   ConfigureFlxFinancialYear

   Dim conConn As New ADODB.Connection

   conConn.Open getConnectionString
   LoadClient conConn
      
   'Resolved By BOSL. Issue: 0000484
   'Modified by Asif. 26-10-2014
'   cmbClient.ListIndex = 0
   
'MsgBox cmbClient.Value
   LoadFlxFinancialYear conConn
   conConn.Close
   Set conConn = Nothing

   flxFinancialYear.row = 0

   Call WheelHook(Me.hWnd)
End Sub

Private Sub cmdAddNew_Click()
   If IsNull(txtClientList.text) Or txtClientList.Tag = "" Then
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      cmdClientList.SetFocus
      Exit Sub
   End If

   Load frmFinancialYearCreate
   'Me.Hide
   If flxFinancialYear.Rows = 1 Then
         flxFinancialYear.Rows = 2
   End If
   If flxFinancialYear.TextMatrix(1, 0) = "" Then
      frmFinancialYearCreate.txtStDate.Locked = False
   Else
      frmFinancialYearCreate.txtStDate.Locked = True

      'Resolved By BOSL. Issue: 0000484
      'Modified by Asif. 26-10-2014
      'frmFinancialYearCreate.txtStDate.text = DateAdd("d", 1, flxFinancialYear.TextMatrix(flxFinancialYear.Rows - 1, 5))
      frmFinancialYearCreate.txtStDate.text = Format$(DateAdd("d", 1, flxFinancialYear.TextMatrix(flxFinancialYear.Rows - 1, 5)), "dd/mm/yyyy")
   End If
   frmFinancialYearCreate.lblClientName.Caption = Me.txtClientList.text
   frmFinancialYearCreate.lblClientName.Tag = Me.txtClientList.Tag
   frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Add New"
   frmFinancialYearCreate.FinancialYearID = UniqueID()
   frmFinancialYearCreate.Show
End Sub
Private Sub ConfigflxFinancialYear()
    flxFinancialYear.Clear
    flxFinancialYear.Rows = 1
    flxFinancialYear.Cols = 10
    flxFinancialYear.RowHeight(0) = 0
    flxFinancialYear.ColWidth(0) = 0 'FYrID
    flxFinancialYear.ColWidth(1) = Label0(1).Left - Label0(0).Left '1500 'ClientID
    flxFinancialYear.ColWidth(2) = Label0(2).Left - Label0(1).Left '1500 'FinancialYear
    flxFinancialYear.ColWidth(3) = Label0(3).Left - Label0(2).Left '1500 'FY_Description
    flxFinancialYear.ColWidth(4) = Label0(4).Left - Label0(3).Left '1500 'FY_StDate
    flxFinancialYear.ColWidth(5) = Label0(5).Left - Label0(4).Left '1500 'FY_EndDate
    flxFinancialYear.ColWidth(6) = 0 'Label0(6).Left - Label0(5).Left '1500 'PeriodsCount
    flxFinancialYear.ColWidth(7) = Label0(5).Left - Label0(4).Left - 600 'Label0(7).Left - Label0(6).Left '1500 'Status
    flxFinancialYear.ColWidth(8) = 0 '1500 'setascurrent
    flxFinancialYear.ColWidth(9) = 0
             
    
End Sub
Private Sub LoadFlxFinancialYear(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim i As Integer
   Call ConfigflxFinancialYear
   szSQL = "SELECT FYrID, ClientID, FinancialYear, FY_Description, " & _
                  "FY_StDate, FY_EndDate, PeriodsCount, Status,setascurrent " & _
           "FROM   FinancialYear  "

   If txtClientList.text <> "" Then szSQL = szSQL & "WHERE  ClientID = '" & txtClientList.Tag & "';"

   'populateGridDefinedHeader adoConn, szSQL, flxFinancialYear
   'I am not using this predefined function you cannot customize easily
   'Now I am writing my code in detail so that I can modify easily 2021-01--23
   Dim adoRst As New ADODB.Recordset
   Dim iRow As Integer
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
   iRow = 1
  ' Exit Sub
   While Not adoRst.EOF
        flxFinancialYear.AddItem ""
        flxFinancialYear.TextMatrix(iRow, 0) = adoRst("FYrID").Value
        flxFinancialYear.TextMatrix(iRow, 1) = adoRst("ClientID").Value
        flxFinancialYear.TextMatrix(iRow, 2) = adoRst("FinancialYear").Value
        flxFinancialYear.TextMatrix(iRow, 3) = adoRst("FY_Description").Value
        flxFinancialYear.TextMatrix(iRow, 4) = adoRst("FY_StDate").Value
        flxFinancialYear.TextMatrix(iRow, 5) = adoRst("FY_EndDate").Value
        flxFinancialYear.TextMatrix(iRow, 6) = adoRst("PeriodsCount").Value
        flxFinancialYear.TextMatrix(iRow, 7) = adoRst("Status").Value
        flxFinancialYear.TextMatrix(iRow, 8) = adoRst("setascurrent").Value
'        flxFinancialYear.TextMatrix(iRow, 9) = adoRst("FYrID").Value
'        flxFinancialYear.TextMatrix(iRow, 10) = adoRst("FYrID").Value
        If flxFinancialYear.TextMatrix(iRow, 7) = "True" Then
           flxFinancialYear.TextMatrix(iRow, 7) = "Open"
        ElseIf flxFinancialYear.TextMatrix(iRow, 7) = "False" Then
           flxFinancialYear.TextMatrix(iRow, 7) = "Close"
        End If
        If flxFinancialYear.TextMatrix(iRow, 8) = True Then
            flxFinancialYear.row = iRow
                  For i = 1 To 8
                        flxFinancialYear.col = i
                        flxFinancialYear.CellFontBold = True
                 Next
        Else
                flxFinancialYear.row = iRow
                For i = 1 To 8
                        flxFinancialYear.col = i
                        flxFinancialYear.CellFontBold = False
                 Next
        End If
        adoRst.MoveNext
        iRow = iRow + 1
   Wend

   
'   For iRow = 1 To flxFinancialYear.Rows - 1
'      If flxFinancialYear.TextMatrix(iRow, 7) = "True" Then
'         flxFinancialYear.TextMatrix(iRow, 7) = "Open"
'      ElseIf flxFinancialYear.TextMatrix(iRow, 7) = "False" Then
'         flxFinancialYear.TextMatrix(iRow, 7) = "Close"
'      End If
'
'      'setascurrent
'
'
'        If flxFinancialYear.TextMatrix(iRow, 8) = True Then
'            flxFinancialYear.row = iRow
'                  For i = 1 To 8
'                        flxFinancialYear.col = i
'                        flxFinancialYear.CellFontBold = True
'                 Next
'        Else
'                flxFinancialYear.row = iRow
'                For i = 1 To 8
'                        flxFinancialYear.col = i
'                        flxFinancialYear.CellFontBold = False
'                 Next
'        End If
'
'   Next iRow
End Sub
Public Function populateGridDefinedHeader(ByVal adoConn As ADODB.Connection, ByVal sSQLQuery As String, ByVal gridMain As MSHFlexGrid, Optional RowHeight As Integer = 240) As Integer
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery, adoConn, adOpenStatic, adLockOptimistic

   populateGridDefinedHeader = adoRst.RecordCount

   gridMain.Rows = 2
   gridMain.RowHeight(1) = RowHeight
   If adoRst.EOF Then
       adoRst.Close
       Set adoRst = Nothing
       Exit Function
   End If

   Dim i As Integer, j As Integer

'   gridMain.AddItem ""

   For i = 0 To adoRst.RecordCount - 1
      For j = 0 To adoRst.Fields.Count - 1
         gridMain.TextMatrix(i + 1, j) = IIf(IsNull(adoRst.Fields(j)), "", adoRst.Fields(j))
      Next j
      gridMain.RowHeight(i + 1) = RowHeight
      adoRst.MoveNext
      If Not adoRst.EOF Then gridMain.AddItem ""
   Next i

   adoRst.Close
   Set adoRst = Nothing

   Exit Function

Error_Handler:
   MsgBox "An Error occurred while populating the grid"
End Function
Private Sub LoadClient(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   cmbClient.ColumnCount = TotalCol + 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'      For j = 0 To TotalCol
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'
'   cmbClient.Column() = Data()
'   cmbClient.ListIndex = 0
    txtClientList.Tag = adoRst.Fields(0).Value
    txtClientList.text = adoRst.Fields(1).Value
    RefreshGrid
   adoRst.Close

NoRes:
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub ConfigureFlxFinancialYear()
   Dim szHeader As String, iCol As Integer

   flxFinancialYear.Clear
   flxFinancialYear.Cols = 9
   flxFinancialYear.Rows = 2
   flxFinancialYear.RowHeight(0) = 0
   szHeader$ = "FYrID|<ClientID|<FinancialYear|<FY_Description|<FY_StDate|<FY_EndDate|PeriodsCount|Status"
   flxFinancialYear.FormatString = szHeader$

   flxFinancialYear.ColWidth(0) = 0                                    'FinancialYrID
   flxFinancialYear.ColWidth(1) = Label0(1).Left - Label0(0).Left      'ClientID
   flxFinancialYear.ColWidth(2) = Label0(2).Left - Label0(1).Left      'Year
   flxFinancialYear.ColWidth(3) = Label0(3).Left - Label0(2).Left      'Description
   flxFinancialYear.ColWidth(4) = Label0(4).Left - Label0(3).Left      'StartDate
   flxFinancialYear.ColWidth(5) = 1400 'flxFinancialYear.Width + flxFinancialYear.Left - Label0(4).Left - 300        'EndDate
   flxFinancialYear.ColWidth(6) = 0                                    'PeriodsCount
   flxFinancialYear.ColWidth(7) = 1200                                   'Status
   flxFinancialYear.ColWidth(8) = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
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
