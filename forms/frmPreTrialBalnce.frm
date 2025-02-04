VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreTrialBalnce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Balance Report"
   ClientHeight    =   7725
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   7140
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
   Icon            =   "frmPreTrialBalnce.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5331.932
   ScaleMode       =   0  'User
   ScaleWidth      =   6704.832
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   315
      ScaleHeight     =   4740
      ScaleWidth      =   6165
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   6195
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
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   10
         Top             =   675
         Width           =   6075
         _ExtentX        =   10716
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   9
         Top             =   375
         Width           =   4500
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "7937;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   8
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   22
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
         Width           =   5715
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2445
      Left            =   43
      TabIndex        =   26
      Top             =   0
      Width           =   6585
      Begin VB.CheckBox chkProperty 
         Caption         =   "Excl."
         Height          =   195
         Left            =   5715
         TabIndex        =   33
         Top             =   945
         Width           =   780
      End
      Begin VB.CommandButton cmdProperty 
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
         Left            =   5310
         TabIndex        =   1
         Top             =   900
         Width           =   300
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
         Left            =   5325
         TabIndex        =   0
         Top             =   495
         Width           =   300
      End
      Begin MSForms.CommandButton cmdFundLookUp 
         Height          =   345
         Left            =   5310
         TabIndex        =   2
         Top             =   1350
         Width           =   300
         Caption         =   ".."
         Size            =   "529;609"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1725
         TabIndex        =   32
         Top             =   495
         Width           =   3780
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6667;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   1710
         TabIndex        =   31
         Top             =   900
         Width           =   3915
         VariousPropertyBits=   746604571
         Size            =   "6906;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   255
         Left            =   585
         TabIndex        =   30
         Top             =   525
         Width           =   555
      End
      Begin MSForms.TextBox txtFundName 
         Height          =   315
         Left            =   1695
         TabIndex        =   29
         Tag             =   "ALL"
         Top             =   1365
         Width           =   3735
         VariousPropertyBits=   746604571
         Size            =   "6588;556"
         Value           =   "All Funds"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Name:"
         Height          =   255
         Index           =   1
         Left            =   585
         TabIndex        =   28
         Top             =   1410
         Width           =   915
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   255
         Index           =   0
         Left            =   585
         TabIndex        =   27
         Top             =   975
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select the Financial Period"
      Height          =   1635
      Left            =   43
      TabIndex        =   17
      Top             =   2610
      Width           =   6615
      Begin MSForms.ComboBox cmbPeriodFrom 
         Height          =   285
         Left            =   2250
         TabIndex        =   4
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
      Begin MSForms.ComboBox cmbPeriodTo 
         Height          =   285
         Left            =   2250
         TabIndex        =   5
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
      Begin MSForms.ComboBox cmbFinancialYear 
         Height          =   285
         Left            =   2250
         TabIndex        =   3
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Year:"
         Height          =   255
         Index           =   66
         Left            =   1050
         TabIndex        =   20
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From:"
         Height          =   255
         Index           =   0
         Left            =   1050
         TabIndex        =   19
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period To:"
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   18
         Top             =   1065
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRefreshData 
      Caption         =   "&Refresh Data"
      Height          =   360
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4620
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4590
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4590
      Width           =   1335
   End
   Begin VB.PictureBox fmePropertyLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   1080
      ScaleHeight     =   2025
      ScaleWidth      =   4200
      TabIndex        =   14
      Top             =   5475
      Visible         =   0   'False
      Width           =   4196
      Begin VB.CommandButton cmdGridPropertyLookup 
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
         Left            =   3795
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPropertyLookup 
         Height          =   1605
         Left            =   80
         TabIndex        =   13
         Top             =   330
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   2831
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   16761024
         ForeColorFixed  =   16777215
         BackColorSel    =   13884353
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorUnpopulated=   15728607
         GridColor       =   12632256
         GridColorFixed  =   4210752
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
      Begin MSForms.TextBox txtSearchProperty 
         Height          =   285
         Left            =   80
         TabIndex        =   12
         Top             =   0
         Width           =   1425
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2514;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmPreTrialBalnce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String

Dim reportingDate As String
Dim sessionID As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim sTextBox As String
Dim szaFundCode() As String
'Private Sub cboClientID_Click()
'    Dim adoConn    As New ADODB.Connection
'   adoConn.Open getConnectionString
'   LoadCmbFinancialYear adoConn
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub chkProperty_Click()
    If chkProperty.Value = 0 Then
        txtPropertyName.text = "ALL"
        cmdProperty.Enabled = True
    Else
        txtPropertyName.text = ""
        txtPropertyName.Tag = ""
        cmdProperty.Enabled = False
    End If
End Sub

Private Sub cmbFinancialYear_Change()
Dim adoConn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim szSQL      As String
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer                    'Open flag index

   If Not IsNull(cmbFinancialYear.Value) Then
      adoConn.Open getConnectionString
      
      szSQL = "SELECT PeriodID, Period_Descp, P_StDate, P_EndDate, Status " & _
              "FROM   Periods " & _
              "WHERE  FYrID = '" & cmbFinancialYear.Value & "' " & _
              "ORDER BY P_StDate ;"

'      Debug.Print szSQL
      
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      If adoRst.EOF Then GoTo NoRes

      TotalRow = adoRst.RecordCount - 1
      TotalCol = adoRst.Fields.Count - 1
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

      adoConn.Close
      Set adoConn = Nothing
   End If
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Periods are not found. Please contact with system support", "Y", "N"
   Set adoConn = Nothing
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
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 0
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
   
   
   adoConn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
            flxClient.TextMatrix(rRow, 0) = ""
            flxClient.TextMatrix(rRow, 1) = "ALL"
            flxClient.TextMatrix(rRow, 2) = "ALL Properties"
            flxClient.RowHeight(rRow) = 280
            flxClient.AddItem ""
             rRow = 2
'            flxClient.TextMatrix(rRow, 0) = ""
'            flxClient.TextMatrix(rRow, 1) = "ZZZZ"
'            flxClient.TextMatrix(rRow, 2) = "Common Properties"
'            flxClient.RowHeight(rRow) = 280
'            flxClient.AddItem ""
'
'
'           rRow = 3
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

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdClientList.SetFocus
End Sub

Private Sub cmdproperty_Click()
        picClient.Left = 269.029
        picClient.Top = 155.299
        sTextBox = "2"
        LoadPropertyList
        Frame1.Enabled = False
        Frame2.Enabled = False
        picClient.Visible = True
        txtSearchClientID.SetFocus
End Sub

Private Sub flxClient_Click()
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
                Dim adoConn As New ADODB.Connection
                adoConn.Open getConnectionString
                LoadCmbFinancialYear adoConn
                adoConn.Close
                chkProperty.Value = 0
                FocusControl cmdProperty
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                cmdFundLookUp.SetFocus
        End If
        If sTextBox = "3" Then
                txtFundName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFundName.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                cmbFinancialYear.SetFocus
        End If
       
       
        picClient.Visible = False
        
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
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
                Dim adoConn As New ADODB.Connection
                adoConn.Open getConnectionString
                LoadCmbFinancialYear adoConn
                adoConn.Close
                cmdProperty.SetFocus
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                cmdFundLookUp.SetFocus
        End If
        If sTextBox = "3" Then
                txtFundName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFundName.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                cmbFinancialYear.SetFocus
        End If
       
       
        picClient.Visible = False
    End If
    If KeyAscii = 27 Then
         picClient.Visible = False
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
           End If
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub

Private Sub txtFundName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdFundLookUp.SetFocus
    End If
End Sub

Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdProperty.SetFocus
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
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
           End If
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
Private Sub cmdFundLookUp_Click()
   If txtClientList.text = "" Then
      ShowMsgInTaskBar "Please select a client to continue.", , "N"
      Exit Sub
   End If

'   fmePropertyLookup.Top = txtFundNo.Top + txtFundNo.Height + 5
'   fmePropertyLookup.Left = txtFundNo.Left - (fmePropertyLookup.Width - txtFundNo.Width) + 200
'   fmePropertyLookup.Visible = True
'   fmePropertyLookup.ZOrder 0
'   gridPropertyLookup.Visible = True
'   txtSearchProperty.text = ""
'   txtSearchProperty.Enabled = True
'   txtSearchProperty.SetFocus
'
'   LOOKUPCommand = "Fund"
'
'   PopulatePropertyLookup IIf(txtClientList.Tag = "ALL", "", " WHERE CLIENTID = '" & txtClientList.Tag & "'")
    picClient.Left = 269.029
    picClient.Top = 225.299
     sTextBox = "3"
     picClient.Visible = True
    
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFunds(adoConn)
    adoConn.Close
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadFunds(conConnection As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim SQLStr1 As String
   SQLStr1 = "SELECT FundID, FundCode, FundName FROM Fund;"
   adoRst.Open SQLStr1, conConnection, adOpenKeyset, adLockReadOnly

   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 2700
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   
'   flxClient.ColWidth(0) = 200
'   flxClient.ColWidth(1) = 0
'   flxClient.ColWidth(2) = 3000
'   picClient.Width = 3500
'   cmdPicCLose.Left = 3200
'   txtSearchClientID.Left = 45

   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   txtSearchClientID.Width = 1500
   txtSearchClientName.Visible = True
'   picClient.Width = 5295
'   flxClient.Width = 5175
   
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   txtSearchClientName.Left = 1580
   ' txtSearchClientName.Width = 3600
'   picClient.Height = 4095
'   flxClient.Height = 3345
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 250
'   lblClientName.Width = 3600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   If adoRst.RecordCount > 0 Then
        ReDim szaFundCode(adoRst.RecordCount, 2) As String
   End If
   
   Dim rRow As Integer
   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
       If sTextBox = "3" Then
            rRow = 1
            While Not adoRst.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
'               flxClient.TextMatrix(rRow, 0) = adoRst.Fields.Item("FundCode").Value
'               flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("FundName").Value
'               flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("FundID").Value
               
               flxClient.TextMatrix(rRow, 0) = "  " & adoRst.Fields.Item("FundID").Value
               flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("FundCode").Value
               flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("FundName").Value
               
               szaFundCode(rRow - 1, 0) = adoRst.Fields.Item("FundCode").Value
               szaFundCode(rRow - 1, 1) = adoRst.Fields.Item("FundName").Value
               szaFundCode(rRow - 1, 2) = adoRst.Fields.Item("FundID").Value
                flxClient.RowHeight(rRow) = 280
               adoRst.MoveNext
               If Not adoRst.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
       End If
   End If
End Sub
Private Sub cmdGridPropertyLookup_Click()
   fmePropertyLookup.Visible = False
End Sub

'Private Sub cmdPropertyLookup_Click()
'   If txtClientList.text = "" Then
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
'   PopulatePropertyLookup IIf(txtClientList.text = "ALL", "", " WHERE CLIENTID = '" & txtClientList.text & "'")
'End Sub

Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Call ExportData2NominalLedger(adoConn)

   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
cmdSCYRRPrint.Enabled = False

Dim adoConn As New ADODB.Connection
Dim adoRst As New ADODB.Recordset

adoConn.Open getConnectionString

On Error GoTo CreateReportTrialBalance

   adoRst.Open "SELECT * FROM ReportTrialBalance;", adoConn, adOpenStatic, adLockReadOnly
   adoRst.Close

   GoTo LoadProfitAndLossReport

CreateReportTrialBalance:
   adoConn.Execute _
      "CREATE TABLE ReportTrialBalance " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "ClientID      TEXT(10), " & _
            "SubHeader     TEXT(200), " & _
            "NominalCode   TEXT(15) NOT NULL, " & _
            "Name          TEXT(200), " & _
            "NominalType   TEXT(50), " & _
            "Balance       CURRENCY, " & _
            "Debit         CURRENCY, " & _
            "Credit        CURRENCY, " & _
            "PRIMARY KEY (ReportingDate, SessionID, NominalCode)" & _
         ");"
         
LoadProfitAndLossReport:
    
    Dim szSQL As String
    Dim periodFrom As String
    Dim periodTo As String
    
On Error GoTo ErrorHandler:
    
    periodFrom = Format(cmbPeriodFrom.Column(2), "dd mmmm yyyy")
    periodTo = Format(cmbPeriodTo.Column(3), "dd mmmm yyyy")
    'Geting the balance sheet Items
    szSQL = GetTrialBalanceQuery(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
    
    adoConn.Execute "DELETE FROM ReportTrialBalance WHERE SessionID = '" & sessionID & "';"
    'added by anol 20161025
    adoConn.Execute "DELETE FROM ReportTrialBalance WHERE ReportingDate < #" & reportingDate & "# ;"
    
    adoConn.Execute "INSERT INTO ReportTrialBalance " & _
    "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
    szSQL
     'added by anol 20170201, issue 296 trial balance is not correct
    'Getting RETAINED EARNINGS BEFORE THE FINANCIAL YEAR
    szSQL = GetTrialBalanceQuery3(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
    adoConn.Execute "INSERT INTO ReportTrialBalance " & _
    "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
    szSQL
    'added by anol 20161025
    'Geting the Profit and loss  Items
    szSQL = GetTrialBalanceQuery2(txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom, periodTo, reportingDate, sessionID)
    adoConn.Execute "INSERT INTO ReportTrialBalance " & _
    "(ReportingDate, SessionID, CLIENTID, SUBHEADER, NOMINALCODE, NAME, NOMINALTYPE, BALANCE, DEBIT, CREDIT) " & _
    szSQL
    
   
'
    Dim reportApp As New CRAXDRT.Application
    Dim Report As CRAXDRT.Report
  
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TrialBalanceNew.rpt")
    
    Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
    
    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    '
    Report.ParameterFields(1).AddCurrentValue sessionID
    Report.ParameterFields(2).AddCurrentValue txtClientList.text
    Report.ParameterFields(3).AddCurrentValue txtPropertyName.text
    Report.ParameterFields(4).AddCurrentValue txtFundName.text
    
    Report.ParameterFields(5).AddCurrentValue Format(cmbPeriodFrom.Column(2), "dd/mm/yyyy")
    Report.ParameterFields(6).AddCurrentValue Format(cmbPeriodTo.Column(3), "dd/mm/yyyy")
    
    Load frmReport
    frmReport.LoadReportViewer Report
    cmdSCYRRPrint.Enabled = True
    Exit Sub
    
ErrorHandler:

    MsgBox ERR.Number & " " & ERR.description, vbExclamation + vbOKOnly, "Could not load trial balance report"
    Set adoRst = Nothing
    cmdSCYRRPrint.Enabled = True
    'old coding........................
    '   If cboClientID.text = "" Then
'      cboClientID.SetFocus
'      ShowMsgInTaskBar "Please select a client", "Y", "N"
'      Exit Sub
'   End If
'   If txtSCYRRStDt.text = "" Then
'      txtSCYRRStDt.SetFocus
'      Exit Sub
'   End If
'   If txtSCYRREnDt.text = "" Then
'      txtSCYRREnDt.SetFocus
'      Exit Sub
'   End If
'
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   'On Error Resume Next
'   adoConn.Open getConnectionString
'
'   If Not AreCA_Setup(adoConn) Then
'      ShowMsgInTaskBar "Please setup control accounts for the client(s)", "Y", "N"
'      adoConn.Close
'      Set adoConn = Nothing
'      Exit Sub
'   End If
'
'   adoConn.Execute "UPDATE NominalLedger " & _
'                   "SET Debit = 0, Credit = 0 " & _
'                   "WHERE Type > 0 AND ClientID = '" & cboClientID.Value & "';"
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   Dim reportApp As New CRAXDRT.Application
'   Dim Report As CRAXDRT.Report
'
''  All option selected
'   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\TrialBalance.rpt")
'
'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
'
'   Report.EnableParameterPrompting = False
'   Report.DiscardSavedData
''
''   Report.ParameterFields(1).AddCurrentValue cboClientID.Value
'''   Report.ParameterFields(2).AddCurrentValue txtPropertyID.text
'''   Report.ParameterFields(3).AddCurrentValue IIf(txtFundNo.text = "ALL", 0, Val(txtFundNo.text))
''   Report.ParameterFields(2).AddCurrentValue CDate(txtSCYRRStDt.text)
''   Report.ParameterFields(3).AddCurrentValue CDate(txtSCYRREnDt.text)
''   Report.ParameterFields(4).AddCurrentValue cboClientID.Column(1)
''   Report.ParameterFields(5).AddCurrentValue txtPropertyName.text
''   Report.ParameterFields(6).AddCurrentValue txtFundName.text
'''   Report.ParameterFields(9).AddCurrentValue gCurrentShopCentreName
'
'   Report.ParameterFields(1).AddCurrentValue cboClientID.Value
'   Report.ParameterFields(2).AddCurrentValue txtPropertyID.text
'   Report.ParameterFields(3).AddCurrentValue IIf(txtFundNo.text = "ALL", 0, Val(txtFundNo.text))
'   Report.ParameterFields(4).AddCurrentValue ""
'   Report.ParameterFields(5).AddCurrentValue ""
'   Report.ParameterFields(6).AddCurrentValue CDate(txtSCYRRStDt.text)
'   Report.ParameterFields(7).AddCurrentValue CDate(txtSCYRREnDt.text)
'   Report.ParameterFields(8).AddCurrentValue cboClientID.Column(1)
'   Report.ParameterFields(9).AddCurrentValue txtPropertyName.text
'   Report.ParameterFields(10).AddCurrentValue txtFundName.text
'   Report.ParameterFields(11).AddCurrentValue gCurrentShopCentreName
'   Load frmReport
'   frmReport.LoadReportViewer Report
End Sub

'Private Sub UpdateDrCr4NC(adoConn As ADODB.Connection)
'   Dim szSQL As String, szPropSrc As String
'   Dim szFundPA As String, szFundDEPT_ID As String, szFundSageDepartment As String
'   Dim szFundSA As String, szFundSR As String, szFundBank As String
'   Dim adoRst As New ADODB.Recordset, adoNL As New ADODB.Recordset
'
'   'Resolved by BOSL
'   'Issue No: 0000476
'   'Retrieving the Sales Ledger Control Account from the Tools > Configuration
'   'Modified By: Asif. 22 Sep 2014
'
'   Dim SalesLedgerControl As String
'   SalesLedgerControl = ""
'
'   SalesLedgerControl = GetNominalCodeForControlAccount(adoConn, "Sales Ledger Control", cboClientID.Value)
'
'   If txtPropertyID.text = "ALL" Then
'      szPropSrc = ""
'   Else
'      szPropSrc = "P.PropertyID = '" & txtPropertyID.text & "' AND "
'   End If
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
''--------------------------------------------------------------------------------------------------
''##########                            Purchase Invoices & Credit - PI, PC
''--------------------------------------------------------------------------------------------------
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
'   'Resolved by BOSL
'   'Issue No: 0000476
'   'Retrieving the Sales Ledger Control Account from the Tools > Configuration
'   'Modified By: Asif. 22 Sep 2014
'
'   'adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '" & SalesLedgerControl & "';", adoConn, adOpenDynamic, adLockOptimistic
'
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
'   'Resolved by BOSL
'   'Issue No: 0000476
'   'Retrieving the Sales Ledger Control Account from the Tools > Configuration
'   'Modified By: Asif. 22 Sep 2014
'
'   SalesLedgerControl = GetNominalCodeForControlAccount(adoConn, "Sales Ledger Control", cboClientID.Value)
'
'   'adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '" & SalesLedgerControl & "';", adoConn, adOpenDynamic, adLockOptimistic
'
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
'   'Resolved by BOSL
'   'Issue No: 0000476
'   'Retrieving the Sales Ledger Control Account from the Tools > Configuration
'   'Modified By: Asif. 22 Sep 2014
'
'   SalesLedgerControl = GetNominalCodeForControlAccount(adoConn, "Sales Ledger Control", cboClientID.Value)
'
'   'adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '1100';", adoConn, adOpenDynamic, adLockOptimistic
'   adoNL.Open "SELECT * FROM NominalLedger WHERE Code = '" & SalesLedgerControl & "';", adoConn, adOpenDynamic, adLockOptimistic
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

Private Sub LoadCmbFinancialYear(adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim Data()     As String
   Dim i          As Integer
   Dim j          As Integer
   Dim K          As Integer              'Open Flag index
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status " & _
           "FROM   FinancialYear " & _
           "WHERE  ClientID = '" & txtClientList.Tag & "' " & _
           "ORDER BY FY_StDate DESC;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1
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

Private Function LoadCmbFunds(adoConn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbFunds = False
   On Error GoTo ErrorHandler

   szSQL = "SELECT FundID, FundCode " & _
           "FROM Fund ORDER BY FundCode;"


   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Funds"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cboC.Column() = Data()
   cboC.ListIndex = 0
   
   LoadCmbFunds = True
   Exit Function

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   'ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub Form_Load()
   Me.Height = 5790
   Me.Width = 6795
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   chkProperty.BackColor = MODULEBACKCOLOR
'   txtSCYRRStDt.text = "01/01/2000"
'   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
'
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   adoConn.Open getConnectionString

   ' Clients
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT  order by CLIENTID"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol - 1, TotalRow - 1) As String
'
'   For i = 0 To TotalRow - 1
'      For j = 0 To TotalCol - 1
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboClientID.Column() = Data()
'   cboClientID.ListIndex = 0
    If Not adoRst.EOF Then
                txtClientList.Tag = adoRst.Fields("CLIENTID").Value
                txtClientList.text = adoRst.Fields("CLIENTNAME").Value
                 adoRst.Close
'                        szSQL = "SELECT PropertyID, PropertyName, " & _
'                             "ProAddressLine1, ProPostCode " & _
'                            "FROM Property " & _
'                           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                           "ORDER BY PropertyID;"
'                         adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                    If Not adoRST.EOF Then
'                        txtPropertyName.text = IIf(IsNull(adoRST.Fields("PropertyName").Value), "", adoRST.Fields("PropertyName").Value)
'                        txtPropertyName.Tag = IIf(IsNull(adoRST.Fields("PropertyID").Value), "", adoRST.Fields("PropertyID").Value)
'                    Else
'                        txtPropertyName.text = ""
'                        txtPropertyName.Tag = ""
'                    End If
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
   End If
   LoadCmbFinancialYear adoConn
  ' LoadCmbFunds adoConn, cmbFund
  ' Call WheelHook(Me.hWnd)
    Call WheelHook(Me.hWnd)
   'Added by BOSL. Issue: 0000476.
   'Added by Asif. 20 Dec 2014
   sessionID = GetTimeStamp
'   MsgBox sessionID
   reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
   If adoRst.State = 1 Then
        adoRst.Close
   End If
   Set adoRst = Nothing
   
   On Error GoTo ErrorHandler
   adoConn.Execute "DELETE FROM ReportProfitAndLoss WHERE ReportingDate < #" & reportingDate & "# ;"
   
ErrorHandler:

  

NoRes:
   adoConn.Close
   Set adoConn = Nothing
End Sub

'Public Function PopulatePropertyLookup(strFilter_ As String)
'   Dim conProperty_ As New ADODB.Connection
'   Dim rstProperty_ As New ADODB.Recordset
'   Dim szSQL As String
'   Dim iRow As Integer
'
'   'On Error Resume Next
'   conProperty_.Open getConnectionString
'
'   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'   If LOOKUPCommand = "Property" Then
'      szSQL = "SELECT PropertyID, PropertyNAME " _
'            & "From Property " & strFilter_
'   End If
'   If LOOKUPCommand = "Fund" Then szSQL = "SELECT FundID, FundName FROM FUND;"
'
'   rstProperty_.Open szSQL, conProperty_, adOpenStatic, adLockReadOnly
'
'   gridPropertyLookup.Clear
'   gridPropertyLookup.Rows = 2
'   gridPropertyLookup.Cols = 2
'   ConfigurFlexGrid
'
'   iRow = 1
'   On Error Resume Next
'   While Not rstProperty_.EOF
'      gridPropertyLookup.TextMatrix(iRow, 0) = IIf(rstProperty_.Fields.Item(0) = Null, "", rstProperty_.Fields.Item(0))
'      gridPropertyLookup.TextMatrix(iRow, 1) = IIf(rstProperty_.Fields.Item(1) = Null, "", rstProperty_.Fields.Item(1))
'
'      rstProperty_.MoveNext
'      If Not rstProperty_.EOF Then gridPropertyLookup.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   rstProperty_.Close
'   conProperty_.Close
'   Set rstProperty_ = Nothing
'   Set conProperty_ = Nothing
'End Function

'Private Sub ConfigurFlexGrid()
'   fmePropertyLookup.Visible = True
'   gridPropertyLookup.Visible = True
'
'   gridPropertyLookup.RowHeight(0) = 255
'   gridPropertyLookup.row = 0
'   Dim i As Integer
'   For i = 0 To gridPropertyLookup.Cols - 1
'        gridPropertyLookup.col = i
'        gridPropertyLookup.CellFontBold = True
'   Next i
'
'   gridPropertyLookup.ColWidth(0) = 800
'
'   If LOOKUPCommand = "Property" Then _
'      gridPropertyLookup.TextMatrix(0, 0) = "ID"
'   If LOOKUPCommand = "Fund" Then _
'      gridPropertyLookup.TextMatrix(0, 0) = "No"
'
'   gridPropertyLookup.ColWidth(1) = 2860
'   gridPropertyLookup.TextMatrix(0, 1) = "Name"
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearReportData "ReportProfitAndLoss", sessionID
   UnLoadForm Me
   'Call WheelUnHook(Me.hWnd)
End Sub

'Private Sub gridPropertyLookup_Click()
'   SEARCHPropertyMODE_ = False
'
''   crash after this line
'   If LOOKUPCommand = "Property" Then
'      txtPropertyName.Tag = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 0)
'      txtPropertyName.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 1)
'   End If
'   If LOOKUPCommand = "Fund" Then
'      txtFundNo.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 0)
'      txtFundName.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 1)
'   End If
'
''    SET OTHERS
'   fmePropertyLookup.Visible = False
'   SEARCHPropertyMODE_ = True
'End Sub

'Private Sub txtSearchProperty_Change()
'   Dim sFilter_ As String
'
'   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
'                 "ORDER BY PropertyID;"
'   PopulatePropertyLookup sFilter_
'End Sub

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

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
            picClient.Visible = False
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
           End If
    End If
End Sub
