VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreVAT_Details 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VAT Details Report"
   ClientHeight    =   6690
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   6945
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
   Icon            =   "frmPreVAT_Details.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4617.557
   ScaleMode       =   0  'User
   ScaleWidth      =   6521.717
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   315
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   17
      Top             =   5715
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
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   11
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   18
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
   Begin VB.Frame Frame5 
      Height          =   1185
      Left            =   45
      TabIndex        =   35
      Top             =   3735
      Width           =   6585
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Vat Transactions Type:"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   37
         Top             =   225
         Width           =   1635
      End
      Begin MSForms.ComboBox cboVAT_Trans_Type 
         Height          =   315
         Left            =   2745
         TabIndex        =   4
         Top             =   225
         Width           =   2610
         VariousPropertyBits=   1820346395
         DisplayStyle    =   3
         Size            =   "4604;556"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "881"
      End
      Begin MSForms.ComboBox cboVatCode 
         Height          =   315
         Left            =   2745
         TabIndex        =   5
         Top             =   705
         Width           =   2610
         VariousPropertyBits=   1820346395
         DisplayStyle    =   3
         Size            =   "4604;556"
         TextColumn      =   1
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058"
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Vat Code:"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   36
         Top             =   705
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date options"
      Height          =   510
      Left            =   45
      TabIndex        =   29
      Top             =   1620
      Width           =   6540
      Begin VB.OptionButton Option1 
         Caption         =   "By Date"
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tax Period/Year"
         Height          =   195
         Left            =   1530
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Left            =   43
      TabIndex        =   22
      Top             =   0
      Width           =   6585
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
         TabIndex        =   8
         Top             =   1350
         Visible         =   0   'False
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         Index           =   0
         Left            =   585
         TabIndex        =   26
         Top             =   525
         Width           =   555
      End
      Begin MSForms.TextBox txtFundName 
         Height          =   315
         Left            =   1695
         TabIndex        =   25
         Tag             =   "ALL"
         Top             =   1365
         Visible         =   0   'False
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
         TabIndex        =   24
         Top             =   1410
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   255
         Index           =   0
         Left            =   585
         TabIndex        =   23
         Top             =   975
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tax Period/year"
      Height          =   1635
      Left            =   810
      TabIndex        =   14
      Top             =   2295
      Width           =   6615
      Begin VB.CheckBox chkCloseTaxPeriod 
         Height          =   195
         Left            =   2970
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   4005
         TabIndex        =   41
         Top             =   270
         Width           =   180
      End
      Begin MSForms.TextBox txtPeriodEnd 
         Height          =   315
         Left            =   4410
         TabIndex        =   40
         Top             =   225
         Width           =   1350
         VariousPropertyBits=   746604571
         MaxLength       =   10
         Size            =   "2381;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtVatPeriod 
         Height          =   315
         Left            =   2475
         TabIndex        =   38
         Top             =   225
         Width           =   1350
         VariousPropertyBits=   746604571
         MaxLength       =   10
         Size            =   "2381;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Period/year:"
         Height          =   195
         Index           =   66
         Left            =   735
         TabIndex        =   16
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Close Tax Period After Report"
         Height          =   255
         Index           =   0
         Left            =   735
         TabIndex        =   15
         Top             =   705
         Visible         =   0   'False
         Width           =   2190
      End
   End
   Begin VB.CommandButton cmdRefreshData 
      Caption         =   "&Refresh Data"
      Height          =   360
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5025
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4995
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   750
      Left            =   45
      TabIndex        =   30
      Top             =   2115
      Visible         =   0   'False
      Width           =   6585
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4305
         TabIndex        =   32
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1995
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   255
         Index           =   3
         Left            =   1380
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   255
         Index           =   2
         Left            =   3945
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmPreVAT_Details"
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

Private Sub cboVAT_Trans_Type_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboVatCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdSCYRRPrint.SetFocus
    End If
End Sub


'Private Sub cmbFinancialYear_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 13 Then
'        cmbPeriodFrom.SetFocus
'    End If
'End Sub

'Private Sub cmbPeriodFrom_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 13 Then
'            cmbPeriodTo.SetFocus
'    End If
'End Sub

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
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                cmdProperty.SetFocus
        End If
        If sTextBox = "3" Then
                txtFundName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFundName.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
''                cmbFinancialYear.SetFocus
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
'                cmbFinancialYear.SetFocus
        End If
       
       
        picClient.Visible = False
    End If
    If KeyAscii = 27 Then
         picClient.Visible = False
          Frame1.Enabled = True
          Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
           End If
    End If
End Sub

Private Sub Option1_Click()
        Frame4.Visible = True
        Frame1.Visible = False
        SelTxtInCtrl txtSCYRRStDt
        txtSCYRRStDt.SetFocus
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

Private Sub txtPeriodEnd_Change()
     TextBoxChangeDate txtPeriodEnd
End Sub

Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdProperty.SetFocus
    End If
End Sub

Private Sub txtSCYRREnDt_Change()
        TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
    SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboVAT_Trans_Type.SetFocus
    End If
    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
    'TextBoxChangeDate txtSCYRREnDt
    TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
    SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSCYRREnDt.SetFocus
    End If
    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
    'TextBoxChangeDate txtSCYRRStDt
    TextBoxFormatDate txtSCYRRStDt
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
          Frame1.Enabled = True
          Frame2.Enabled = True
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
   'txtSearchClientName.Width = 3600
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
'Private Sub cmdGridPropertyLookup_Click()
'   fmePropertyLookup.Visible = False
'End Sub

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
    If Option2.Value = False Then
       If txtSCYRRStDt.text = "" Then
          txtSCYRRStDt.SetFocus
          Exit Sub
       End If
       If txtSCYRREnDt.text = "" Then
          txtSCYRREnDt.SetFocus
          Exit Sub
       End If
    End If
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   'On Error Resume Next
   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE NominalLedger " & _
                   "SET Debit = 0, Credit = 0 " & _
                   "WHERE Type > 0;"

'   UpdateDrCr4NC adoconn

   adoConn.Close
   Set adoConn = Nothing
'rem by anol 2020-08-25
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\Vat_Details.rpt")

'   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

    Report.EnableParameterPrompting = False
    Report.DiscardSavedData
    
    Report.ParameterFields(1).AddCurrentValue txtClientList.Tag
    Report.ParameterFields(2).AddCurrentValue txtPropertyName.Tag
   
    Report.ParameterFields(3).AddCurrentValue cboVAT_Trans_Type.text
    Report.ParameterFields(4).AddCurrentValue CDate(txtVatPeriod.text)
    Report.ParameterFields(5).AddCurrentValue CDate(txtPeriodEnd.text)
    
    Report.ParameterFields(6).AddCurrentValue txtClientList.text
    Report.ParameterFields(7).AddCurrentValue txtPropertyName.text
    Report.ParameterFields(8).AddCurrentValue cboVatCode.text
    Report.ParameterFields(9).AddCurrentValue gCurrentShopCentreName
    Report.ParameterFields(10).AddCurrentValue CInt(cboVAT_Trans_Type.Value)
     Report.ParameterFields(11).AddCurrentValue True
  
   Load frmReport
   frmReport.LoadReportViewer Report
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
'   Dim szSQL      As String
'   Dim TotalRow   As Integer
'   Dim TotalCol   As Integer
'   Dim Data()     As String
'   Dim i          As Integer
'   Dim j          As Integer
'   Dim K          As Integer              'Open Flag index
'   Dim adoRst     As New ADODB.Recordset
'
'   szSQL = "SELECT FYrID, FinancialYear, ClientID, FY_StDate, Status " & _
'           "FROM   FinancialYear " & _
'           "WHERE  ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY FY_StDate DESC;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'   ReDim Data(TotalCol, TotalRow) As String
'
'   K = -1
'   For i = 0 To TotalRow
'      For j = 0 To TotalCol
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'         If K = -1 And j = 4 Then
'            If adoRst.Fields("Status").Value Then
'               K = i
''               dtStartPnL = CDate(adoRst.Fields("FY_StDate").Value)
''               dtStartBS = CDate("01 January 2000")
'            End If
'         End If
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cmbFinancialYear.Column() = Data()
'   cmbFinancialYear.ListIndex = K
'
'   adoRst.Close
'   Set adoRst = Nothing
'   Exit Sub
'
'NoRes:
'   Set adoRst = Nothing
'   ShowMsgInTaskBar "Financial year has not been created for the client", "Y", "N"
'   Exit Sub
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
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub Option2_Click()
    Frame4.Visible = False
    Frame1.Visible = True
'    cmbFinancialYear.SetFocus
End Sub
Private Sub Form_Load()
   Me.Height = 6030
   Me.Width = 6795
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Frame1.Left = 40
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   Frame3.BackColor = MODULEBACKCOLOR
   Option2.BackColor = MODULEBACKCOLOR
   Option1.BackColor = MODULEBACKCOLOR
   Frame5.BackColor = MODULEBACKCOLOR
   Frame4.BackColor = MODULEBACKCOLOR
   chkCloseTaxPeriod.BackColor = MODULEBACKCOLOR
   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
   
   ReDim Data(1, 4) As String
   
   
   Data(0, 0) = "5"
   Data(1, 0) = "All"
   Data(0, 1) = "1"
   Data(1, 1) = "Sales"
   Data(0, 2) = "2"
   Data(1, 2) = "Purchase"
   Data(0, 3) = "3"
   Data(1, 3) = "Cash"
   Data(0, 4) = "4"
   Data(1, 4) = "Nominal"
  
   cboVAT_Trans_Type.Column() = Data()
   cboVAT_Trans_Type.ListIndex = 0
   
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   adoConn.Open getConnectionString

   ' Clients
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT "
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
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
                
   End If
   LoadCmbFinancialYear adoConn
   If adoRst.State = 1 Then
        adoRst.Close
   End If
   szSQL = "SELECT VAT_CODE, VAT_RATE FROM tlbVatCode WHERE IN_USE;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count
   
   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All VAT"
   
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboVatCode.Column() = Data()
   cboVatCode.ListIndex = 0
   adoRst.Close
   
  ' LoadCmbFunds adoConn, cmbFund
   
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

   Call WheelHook(Me.hWnd)

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
   'ClearReportData "ReportProfitAndLoss", sessionID
   'Call WheelUnHook(Me.hWnd)
   UnLoadForm Me
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

       '        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
                        If Not ctl Is picClient Then
                            PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
                        Else
                            bHandled = False
                        End If

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
          Frame1.Enabled = True
          Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
                cmdProperty.SetFocus
           ElseIf sTextBox = "3" Then
                cmdFundLookUp.SetFocus
           End If
    End If
End Sub
Public Sub TextBoxChangeDatePeriod(conTextBox As Control)
   'If bPutSlash Then
      If Len(conTextBox.text) = 1 Then
         'conTextBox = "0" & conTextBox.text
         conTextBox.SelStart = Len(conTextBox.text)
         Exit Sub
      End If
'      If Len(conTextBox.text) = 4 Then
'         conTextBox = Left(conTextBox.text, 3) & "0" & Right(conTextBox.text, 1)
'         conTextBox.SelStart = Len(conTextBox.text)
'         Exit Sub
'      End If
    '  bPutSlash = False
   'End If
   If Len(conTextBox.text) = 2 Then
      conTextBox.text = conTextBox.text + "/"
      conTextBox.SelStart = Len(conTextBox.text)
      Exit Sub
   End If
'   If Len(conTextBox.text) = 5 Then
'      conTextBox.text = conTextBox.text + "/"
'      conTextBox.SelStart = Len(conTextBox.text)
'   End If
'   conTextBox.SelStart = Len(conTextBox.text)
End Sub
Private Sub txtVatPeriod_Change()
    TextBoxChangeDate txtVatPeriod
End Sub

Private Sub txtVatPeriod_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    Dim ascd As Integer
'    ascd = KeyCode
'    If ascd = 47 Then
'        txtVatPeriod.text = Replace(txtVatPeriod, "/", "")
'            Exit Sub
'    End If
End Sub

Private Sub txtVatPeriod_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    Dim ascd As Integer
'    ascd = KeyAscii
'    If ascd = 47 Then
'        txtVatPeriod.text = Replace(txtVatPeriod, "/", "")
'            Exit Sub
'    End If
'    TextBoxKeyPrsDate txtVatPeriod, ascd
End Sub
