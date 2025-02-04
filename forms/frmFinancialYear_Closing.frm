VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFinancialYear_Closing 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Closing Financial Years"
   ClientHeight    =   11355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmFinancialYear_Closing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11355
   ScaleWidth      =   11160
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   2520
      ScaleHeight     =   5055
      ScaleWidth      =   6255
      TabIndex        =   21
      Top             =   585
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
         Height          =   4335
         Left            =   45
         TabIndex        =   23
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7646
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   24
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
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Open Financial Year"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Close Financial Year"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6975
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9F9F9&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8085
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
         Left            =   5295
         TabIndex        =   18
         Top             =   180
         Width           =   345
      End
      Begin MSForms.TextBox txtClientID 
         Height          =   285
         Left            =   675
         TabIndex        =   20
         Top             =   180
         Width           =   1305
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "2302;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1980
         TabIndex        =   19
         Top             =   180
         Width           =   3375
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5953;503"
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
         TabIndex        =   13
         Top             =   180
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
         Left            =   5640
         TabIndex        =   8
         Top             =   540
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
         Left            =   4920
         TabIndex        =   7
         Top             =   540
         Width           =   1125
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
         TabIndex        =   6
         Top             =   540
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
         TabIndex        =   5
         Top             =   540
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
         Left            =   6270
         TabIndex        =   4
         Top             =   180
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
         Left            =   5730
         TabIndex        =   3
         Top             =   180
         Width           =   345
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFinancialYear 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8070
      _ExtentX        =   14235
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
      Left            =   7320
      TabIndex        =   17
      Top             =   960
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
      TabIndex        =   14
      Top             =   960
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
      TabIndex        =   12
      Top             =   960
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
      TabIndex        =   11
      Top             =   960
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
      TabIndex        =   10
      Top             =   960
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
      Left            =   6360
      TabIndex        =   9
      Top             =   960
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
      Top             =   960
      Width           =   8085
   End
End
Attribute VB_Name = "frmFinancialYear_Closing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 155.299
    LoadflxClient
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub



Private Sub flxClient_Click()
     txtClientID.text = flxClient.TextMatrix(flxClient.row, 1)
     txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
     RefreshGrid
     picClient.Visible = False
     FocusControl cmdClientList
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            txtClientID.text = flxClient.TextMatrix(flxClient.row, 1)
            txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
            RefreshGrid
            picClient.Visible = False
            FocusControl cmdClientList
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

'   For iRow = 1 To flxFinancialYear.Rows - 1
'      If flxFinancialYear.TextMatrix(iRow, 1) = txtClientID.text Then
'         flxFinancialYear.RowHeight(iRow) = 240
'      Else
'         flxFinancialYear.RowHeight(iRow) = 0
'      End If
'   Next iRow
End Sub

Private Sub cmbClient_Click()
   RefreshGrid
'   flxFinancialYear.row = 0
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdEdit_Click()
'   If cmbClient.text = "" Then
'      ShowMsgInTaskBar "Please select a client", "Y", "N"
'      cmbClient.SetFocus
'      Exit Sub
'   End If
'
'   If flxFinancialYear.TextMatrix(flxFinancialYear.row, 0) = "" Or flxFinancialYear.row = 0 Then
'      ShowMsgInTaskBar "Please select a financial year", "Y", "N"
'      flxFinancialYear.SetFocus
'      Exit Sub
'   End If
'
'   If flxFinancialYear.TextMatrix(flxFinancialYear.row, 7) = "False" Then
'      ShowMsgInTaskBar "Financial year is closed. Edit not allowed", "Y", "N"
'      Exit Sub
'   End If
'
'   Load frmFinancialYearCreate
'   frmFinancialYearCreate.FinancialYearID = flxFinancialYear.TextMatrix(flxFinancialYear.row, 0)
'   Me.Hide
'
'   frmFinancialYearCreate.txtStDate.Locked = IIf(flxFinancialYear.row = 1, False, True)
'
'   frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Edit"
'   frmFinancialYearCreate.cmdSave.Visible = False
'   frmFinancialYearCreate.optPreDefined.Enabled = False
'   frmFinancialYearCreate.LoadFlxPeriods
'   frmFinancialYearCreate.optCustome.Value = True
'   frmFinancialYearCreate.Show

 If flxFinancialYear.TextMatrix(flxFinancialYear.row, 7) = "Open" Then
      If MsgBox("Closing the financial year will also close all the open periods in the financial year." & Chr(13) & _
                "Do you wish to continue?", vbYesNo + vbQuestion, "Close opened Financial Year") = vbNo Then Exit Sub

      Dim conConn As New ADODB.Connection
      conConn.Open getConnectionString

      'Resolved By BOSL. Issue: 0000484
      'Modified by Asif. 26-10-2014
      CloseFinancialYear flxFinancialYear.TextMatrix(flxFinancialYear.row, 0), conConn
      
      ShowMsgInTaskBar "Financial Year is closed", "Y", "P"
      ConfigureFlxFinancialYear
      LoadFlxFinancialYear conConn
      
      conConn.Close
      Set conConn = Nothing
 End If
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
   lblPeriods.Caption = flxFinancialYear.TextMatrix(flxFinancialYear.row, 6)
   txtClientID.text = flxFinancialYear.TextMatrix(flxFinancialYear.row, 1)
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 6325
   Me.Width = 8430
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR

   ConfigureFlxFinancialYear

   Dim conConn As New ADODB.Connection

   conConn.Open getConnectionString
   LoadFirstClient conConn
   
   'Resolved By BOSL. Issue: 0000484
   'Modified by Asif. 26-10-2014
'   cmbClient.ListIndex = 0

   LoadFlxFinancialYear conConn
   conConn.Close
   Set conConn = Nothing

   flxFinancialYear.row = 0

   Call WheelHook(Me.hWnd)
End Sub

Private Sub cmdAddNew_Click()
'   If IsNull(cmbClient.Value) Or cmbClient.Value = "" Then
'      ShowMsgInTaskBar "Please select a client", "Y", "N"
'      cmbClient.SetFocus
'      Exit Sub
'   End If
'
'   Load frmFinancialYearCreate
'   Me.Hide
'
'   If flxFinancialYear.TextMatrix(1, 0) = "" Then
'      frmFinancialYearCreate.txtStDate.Locked = False
'   Else
'      frmFinancialYearCreate.txtStDate.Locked = True
'
'      frmFinancialYearCreate.txtStDate.text = DateAdd("d", 1, flxFinancialYear.TextMatrix(flxFinancialYear.Rows - 1, 5))
'   End If
'
'   frmFinancialYearCreate.Caption = frmFinancialYearCreate.Caption & " - Add New"
'   frmFinancialYearCreate.FinancialYearID = UniqueID()
'   frmFinancialYearCreate.Show

 If flxFinancialYear.TextMatrix(flxFinancialYear.row, 7) = "Close" Then
      If MsgBox("Opening a closed financial year will also open all the closed periods in the financial year." & Chr(13) & _
                "Do you wish to continue?", vbYesNo + vbQuestion, "Open Financial Year") = vbNo Then Exit Sub

      Dim conConn As New ADODB.Connection
      conConn.Open getConnectionString

      'Resolved By BOSL. Issue: 0000484
      'Modified by Asif. 26-10-2014
      OpenFinancialYear flxFinancialYear.TextMatrix(flxFinancialYear.row, 0), conConn
      
      ShowMsgInTaskBar "Financial Year is opened", "Y", "P"
      ConfigureFlxFinancialYear
      LoadFlxFinancialYear conConn
      
      conConn.Close
      Set conConn = Nothing
 End If
End Sub

Private Sub LoadFlxFinancialYear(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim iRow  As Integer

   szSQL = "SELECT FYrID, ClientID, FinancialYear, FY_Description, " & _
                  "FY_StDate, FY_EndDate, PeriodsCount, Status " & _
           "FROM   FinancialYear WHERE  ClientID = '" & txtClientID.Value & "';"

  

   populateGridDefinedHeader adoConn, szSQL, flxFinancialYear

   For iRow = 1 To flxFinancialYear.Rows - 1
      If flxFinancialYear.TextMatrix(iRow, 7) = "True" Then
         flxFinancialYear.TextMatrix(iRow, 7) = "Open"
      Else
         flxFinancialYear.TextMatrix(iRow, 7) = "Close"
      End If
   Next iRow
End Sub
Private Sub LoadFirstClient(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes
   txtClientID.text = adoRst("CLIENTID").Value
   txtClientList.text = adoRst("CLIENTNAME").Value
   adoRst.Close

NoRes:
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub
'Private Sub LoadClient(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
''   cmbClient.ColumnCount = TotalCol + 1
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
'   adoRst.Close
'
'NoRes:
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox Err.description & "::" & Err.Number
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub ConfigureFlxFinancialYear()
   Dim szHeader As String, iCol As Integer

   flxFinancialYear.Clear
   flxFinancialYear.Cols = 8
   flxFinancialYear.Rows = 2
   flxFinancialYear.RowHeight(0) = 0
   szHeader$ = "FYrID|<ClientID|<FinancialYear|<FY_Description|<FY_StDate|<FY_EndDate|PeriodsCount|Status"
   flxFinancialYear.FormatString = szHeader$

   flxFinancialYear.ColWidth(0) = 0                                    'FinancialYrID
   flxFinancialYear.ColWidth(1) = Label0(1).Left - Label0(0).Left      'ClientID
   flxFinancialYear.ColWidth(2) = Label0(2).Left - Label0(1).Left      'Year
   flxFinancialYear.ColWidth(3) = Label0(3).Left - Label0(2).Left      'Description
   flxFinancialYear.ColWidth(4) = Label0(4).Left - Label0(3).Left      'StartDate
   flxFinancialYear.ColWidth(5) = Label0(5).Left - Label0(4).Left      'EndDate
   flxFinancialYear.ColWidth(6) = 0                                    'PeriodsCount
   flxFinancialYear.ColWidth(7) = flxFinancialYear.Width + flxFinancialYear.Left - Label0(5).Left - 300 'Status
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
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
