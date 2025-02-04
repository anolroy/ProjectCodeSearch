VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBudgetView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Budget"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBudgetView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Cancel          =   -1  'True
      Caption         =   "&Print"
      Height          =   420
      Left            =   3285
      TabIndex        =   39
      Top             =   6300
      Width           =   1200
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   4095
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   30
      Top             =   1440
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
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   32
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   33
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
      Left            =   5580
      TabIndex        =   0
      Top             =   135
      Width           =   300
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
      Height          =   345
      Left            =   9555
      TabIndex        =   1
      Top             =   135
      Width           =   300
   End
   Begin VB.TextBox txtTotalBudget 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   6300
      Width           =   1695
   End
   Begin VB.TextBox txtTotalActual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6320
      Width           =   1095
   End
   Begin VB.TextBox txtTotalVariance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12195
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6320
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   12690
      TabIndex        =   9
      Top             =   6720
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Display"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Generate Payment later"
      Top             =   675
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Service Charge Period:"
      Height          =   675
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   12555
      Begin VB.CheckBox chkExcludeInc 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Exclude Income"
         Height          =   255
         Left            =   10935
         TabIndex        =   7
         Top             =   225
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkYtD 
         BackColor       =   &H00F7F7F7&
         Caption         =   "Y&TD"
         Height          =   255
         Left            =   10005
         TabIndex        =   6
         Top             =   225
         Value           =   1  'Checked
         Width           =   735
      End
      Begin MSForms.ComboBox cmbPeriodFrom 
         Height          =   285
         Left            =   5100
         TabIndex        =   4
         Top             =   225
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Period From:"
         Height          =   195
         Index           =   12
         Left            =   4200
         TabIndex        =   17
         Top             =   270
         Width           =   885
      End
      Begin MSForms.ComboBox cmbPeriodTo 
         Height          =   285
         Left            =   7920
         TabIndex        =   5
         Top             =   225
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Period To:"
         Height          =   195
         Index           =   14
         Left            =   7155
         TabIndex        =   16
         Top             =   270
         Width           =   705
      End
      Begin MSForms.ComboBox cmbFinancialYear 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   225
         Width           =   2880
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5080;503"
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
         Height          =   195
         Index           =   66
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   1005
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBudget 
      Height          =   4575
      Left            =   75
      TabIndex        =   10
      Top             =   1680
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
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
   Begin MSForms.TextBox txtFundName 
      Height          =   315
      Left            =   10725
      TabIndex        =   29
      Tag             =   "ALL"
      Top             =   135
      Width           =   2385
      VariousPropertyBits=   746604571
      Size            =   "4207;556"
      Value           =   "All Funds"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   6690
      TabIndex        =   28
      Tag             =   "ALL"
      Top             =   135
      Width           =   2880
      VariousPropertyBits=   746604571
      Size            =   "5080;556"
      Value           =   "ALL Properties"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   810
      TabIndex        =   27
      Top             =   135
      Width           =   4770
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "8414;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdFundLookUp 
      Height          =   345
      Left            =   13110
      TabIndex        =   2
      Top             =   135
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Variance"
      Height          =   255
      Index           =   5
      Left            =   11760
      TabIndex        =   26
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal Name"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   23
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nominal Code"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Index           =   6
      Left            =   6720
      TabIndex        =   20
      Top             =   6315
      Width           =   390
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "SC Fund"
      Height          =   195
      Index           =   9
      Left            =   10035
      TabIndex        =   14
      Top             =   165
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   16000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   6
      Left            =   6000
      TabIndex        =   13
      Top             =   165
      Width           =   645
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   5
      Left            =   195
      TabIndex        =   12
      Top             =   165
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   16000
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frmBudgetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dtStartPnL    As Date
Public dtStartBS     As Date
Public dtEnd         As Date
Dim sTextBox As String
Dim szSelectedFundCode As String

Private Sub cmdPicCLose_Click()
        picClient.Visible = False
        Frame1.Enabled = True
End Sub

Private Sub cmdPrint_Click()
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report
      Dim rep As frmReport
      Dim strPeriodFrom As String
      Dim strPeriodTo As String
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\BudgetView.rpt")

      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData
      strPeriodFrom = cmbPeriodFrom.text
      strPeriodTo = cmbPeriodTo.text
      Report.ParameterFields(1).AddCurrentValue txtClientList.Tag
      Report.ParameterFields(2).AddCurrentValue txtPropertyName.Tag
      Report.ParameterFields(3).AddCurrentValue txtPropertyName.text
      Report.ParameterFields(4).AddCurrentValue txtFundName.text
      Report.ParameterFields(5).AddCurrentValue szSelectedFundCode
      Report.ParameterFields(6).AddCurrentValue strPeriodFrom
      Report.ParameterFields(7).AddCurrentValue strPeriodTo
      Report.ParameterFields(8).AddCurrentValue txtClientList.text
      
      Set rep = New frmReport
      Load rep
      rep.LoadReportViewer Report
End Sub

Private Sub cmdProperty_Click()
        picClient.Left = 4747.029
        picClient.Top = 155.299
        sTextBox = "2"
        LoadPropertyList
        'SSTab1.Enabled = False
        Frame1.Enabled = False
        picClient.Visible = True
        txtSearchClientID.SetFocus
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

Private Sub flxClient_Click()
'        SSTab1.Enabled = True
          Frame1.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = "ALL Properties"
                txtPropertyName.Tag = "ALL"
                flxBudget.Clear
                Dim adoconn As New ADODB.Connection
                adoconn.Open getConnectionString
                LoadCmbFinancialYear adoconn
                adoconn.Close
                ConfigFlxBudget
                txtTotalActual.text = "0.00"
                txtTotalBudget.text = "0.00"
                txtTotalVariance.text = "0.00"
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                ConfigFlxBudget
                txtTotalActual.text = "0.00"
                txtTotalBudget.text = "0.00"
                txtTotalVariance.text = "0.00"
                cmdFundLookUp.SetFocus
        End If
        If sTextBox = "3" Then
                txtFundName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtFundName.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                szSelectedFundCode = flxClient.TextMatrix(flxClient.row, 1)
                ConfigFlxBudget
                txtTotalActual.text = "0.00"
                txtTotalBudget.text = "0.00"
                txtTotalVariance.text = "0.00"
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
         ' SSTab1.Enabled = True
          Frame1.Enabled = True
          flxClient_Click
    End If
    If KeyAscii = 27 Then
            picClient.Visible = False
'            SSTab1.Enabled = True
            Frame1.Enabled = True
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
'          SSTab1.Enabled = True
            Frame1.Enabled = True
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
    picClient.Left = 5500.029
    picClient.Top = 225.299
     sTextBox = "3"
     picClient.Visible = True
    
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call LoadFunds(adoconn)
    adoconn.Close
'    SSTab1.Enabled = False
      Frame1.Enabled = False
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
             flxClient.TextMatrix(rRow, 0) = "ALL"
               flxClient.TextMatrix(rRow, 1) = "ALL"
               flxClient.TextMatrix(rRow, 2) = "ALL Funds"
               szaFundCode(rRow - 1, 0) = "ALL"
               szaFundCode(rRow - 1, 1) = "ALL"
               szaFundCode(rRow - 1, 2) = "ALL Funds"
               flxClient.RowHeight(rRow) = 280
               flxClient.AddItem ""
            rRow = 2
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


Private Sub cboClient_Change()
'   If MiLoading Then Exit Sub
   If txtClientList.text = "" Then Exit Sub

   On Error GoTo ERR_HANDER

   Dim K          As Integer
   Dim iRow       As Integer
   Dim adoconn    As New ADODB.Connection
   Dim szSQL      As String
   Dim adoRstSrc  As New ADODB.Recordset
   Dim adoRstDst  As New ADODB.Recordset
   
   adoconn.Open getConnectionString
   
   LoadCmbFinancialYear adoconn
   
   'LoadCmbProperties adoConn, cboProperty
   
'   GenerateNominalAccounts adoConn

   adoconn.Close
   Set adoconn = Nothing

   Exit Sub

ERR_HANDER:
   ShowMsgInTaskBar "Nominal Ledger could not be loaded for the selected client", "Y", "N"
   MsgBox Err.description
'   adoConn.RollbackTrans
   adoconn.Close
   Set adoconn = Nothing
End Sub
Private Sub createTable(adoconn As ADODB.Connection)
    Dim adoRst As New ADODB.Recordset
    On Error GoTo CreateReportBudVsAE

   adoRst.Open "SELECT * FROM ReportBudVsAE;", adoconn, adOpenStatic, adLockReadOnly
   adoRst.Close

   GoTo LoadBudVsAE

CreateReportBudVsAE:
   adoconn.Execute _
      "CREATE TABLE ReportBudVsAE " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "ClientID      TEXT(10), " & _
            "NC   TEXT(15) NOT NULL, " & _
            "tBalance       CURRENCY, " & _
            "Budget         CURRENCY, " & _
            "Variance        CURRENCY, " & _
            "PRIMARY KEY (ReportingDate, SessionID, NC)" & _
         ");"

LoadBudVsAE:


End Sub
Private Sub cboProperty_Change()

Dim adoconn As New ADODB.Connection
adoconn.Open getConnectionString

If Not IsNull(txtPropertyName.Tag) And txtPropertyName.Tag <> "ALL" Then
   cmbFinancialYear.Value = GetCurrentFYFromProperty(adoconn, txtPropertyName.Tag)
End If

If cmbFinancialYear.ListIndex < 0 Then
   cmbFinancialYear.ListIndex = 0
End If

adoconn.Close
Set adoconn = Nothing
End Sub

Private Sub chkYtD_Click()
   If chkYtD.Value = 1 Then
      cmbPeriodFrom.ListIndex = -1
      cmbPeriodFrom.Enabled = False
      
      If cmbPeriodTo.ListCount > 0 Then
        cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If
   Else
      cmbPeriodFrom.Enabled = True
      If cmbPeriodFrom.ListCount > 0 Then
        cmbPeriodFrom.ListIndex = 0
      End If
      
      If cmbPeriodTo.ListCount > 0 Then
        cmbPeriodTo.ListIndex = cmbPeriodTo.ListCount - 1
      End If
   End If
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

      adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

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
                  dtEnd = CDate(adoRst.Fields("P_EndDate").Value)
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
      
      chkYtD_Click

      adoconn.Close
      Set adoconn = Nothing
   End If
   Exit Sub

NoRes:
   ShowMsgInTaskBar "Periods are not found. Please contact with system support", "Y", "N"
   Set adoconn = Nothing
End Sub

Private Sub cmdClientList_Click()
     picClient.Left = 269.029
    picClient.Top = 155.299
    sTextBox = "1"
    LoadflxClient
'    SSTab1.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
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

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
        cmdOK.Enabled = False
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        
        Call ExportData2NominalLedger(adoconn)
        GenerateNominalAccounts adoconn
        '
        'UpdateBudgetBalance adoConn
        'UpdateVariance
        
        adoconn.Close
        Set adoconn = Nothing
        cmdOK.Enabled = True
        FocusControl cmdClientList
End Sub

Private Function GenerateNominalAccounts(adoconn As ADODB.Connection)
   
   Dim szSQL As String
   Dim clientID As String
   
   
    If IsNull(txtClientList.Tag) Then
       ShowMsgInTaskBar "Please select a client", "Y", "N"
       cmdClientList.SetFocus
       Exit Function
    End If
    If IsNull(cmbFinancialYear.Value) Then
       ShowMsgInTaskBar "Please select the financial year", "Y", "N"
       cmbFinancialYear.SetFocus
       Exit Function
    End If
    If chkYtD.Value = 0 And IsNull(cmbPeriodFrom.Value) Then
       ShowMsgInTaskBar "Please select the period from date", "Y", "N"
       cmbPeriodFrom.SetFocus
       Exit Function
    End If
    If IsNull(cmbPeriodTo.Value) Then
       ShowMsgInTaskBar "Please select the period to date", "Y", "N"
       cmbPeriodTo.SetFocus
       Exit Function
    End If
    
    If chkYtD.Value = 0 Then
       If CDate(cmbPeriodFrom.Column(2)) > CDate(cmbPeriodTo.Column(2)) Then
          ShowMsgInTaskBar "From date cannot be after To date", "Y", "N"
          cmbPeriodFrom.SetFocus
          Exit Function
       End If
    End If
    
   
   Dim periodFrom As String
   Dim periodTo As String
   
    If chkYtD.Value = 1 Then
       If IsNull(cmbFinancialYear.Column(3)) Then
            MsgBox "Financial Year start date cannot be null", vbInformation, "Warning"
            Exit Function
       End If
       dtStartPnL = cmbFinancialYear.Column(3)            'Beginning of the Financial year
       dtStartBS = CDate("01 January 2000")               'Beginning of the System
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
       
       
    Else
       dtStartPnL = cmbPeriodFrom.Column(2)
       dtStartBS = cmbPeriodFrom.Column(2)
       dtEnd = cmbPeriodTo.Column(3)
       
       periodFrom = Format(dtStartPnL, "dd mmmm yyyy")
       periodTo = Format(dtEnd, "dd mmmm yyyy")
    End If
   
    ' Check if Nominal Ledger is setup for the selected client:
'    NominalLedgerSetupForNewClient adoConn


'   szHeader$ = "<Code|<Name|<NLTypeCode|<TypeValue|<STName|>DebitBalance|>CreditBalance|ClientID|<Posting|DrCr|SubType"
'                   x    x        0           x         x         x              x          0        x      0     0
  'Comment out by anol 30 Nov 2015
  
'   If chkYtD.Value = 1 Then
'        szSQL = GetNominalBalancesforBudget(txtClientList.tag, txtPropertyName.tag, txtFundName.TAG, periodFrom, periodTo, True, cmbFinancialYear.Value)
'   Else
'        szSQL = GetNominalBalancesforBudget(txtClientList.tag, txtPropertyName.tag, txtFundName.TAG, periodFrom, periodTo, False, cmbFinancialYear.Value)
'   End If
        Dim sessionID1 As String
        Dim reportingDate As String
        Dim periodFrom1 As String
        'Dim periodTo As String
        'Dim szSQL As String
   
   
        sessionID1 = GetTimeStamp
        reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
        If IsNull(cmbFinancialYear.Column(0)) Then Exit Function
        If IsNull(txtClientList.Tag) Then Exit Function
        If IsNull(txtFundName.Tag) Then Exit Function
        If chkYtD.Value = 1 Then
            periodFrom1 = Format(cmbFinancialYear.Column(3), "dd mmmm yyyy") ''Beginning of the Financial year
            periodTo = Format(cmbPeriodTo.Column(3), "dd mmmm yyyy")
        Else
            periodFrom1 = Format(cmbPeriodFrom.Column(2), "dd mmmm yyyy")
            periodTo = Format(cmbPeriodTo.Column(3), "dd mmmm yyyy")
        End If
        
'         adoConn.Execute "DELETE FROM ReportBudVsAE WHERE ReportingDate < #" & reportingDate & "# ;"
'         szSQL = GetBudVsAEQuery(cmbFinancialYear.Column(0), txtClientList.tag, txtPropertyName.tag, txtFundName.TAG, periodFrom1, periodTo, reportingDate, sessionID1, "1")
'            adoConn.Execute _
'      "INSERT INTO ReportBudVsAE " & _
'      "(ReportingDate, SessionID, CLIENTID,NC,tBalance,Budget,Variance) " & _
'        szSQL
'
'        szSQL = GetBudVsAEQuery(cmbFinancialYear.Column(0), txtClientList.tag, txtPropertyName.tag, txtFundName.TAG, periodFrom1, periodTo, reportingDate, sessionID1, "2")
'            adoConn.Execute _
'      "INSERT INTO ReportBudVsAE " & _
'      "(ReportingDate, SessionID, CLIENTID,NC,tBalance,Budget,Variance) " & _
'        szSQL
        
       adoconn.Execute "DELETE FROM ReportBudVsAE ;"
        szSQL = GetBudVsProftandloss(cmbFinancialYear.Column(0), txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom1, periodTo, reportingDate, sessionID1, 1)
            adoconn.Execute _
      "INSERT INTO ReportBudVsAE " & _
      "(ReportingDate, SessionID, CLIENTID,NC,tBalance,Budget,Variance) " & _
        szSQL
        szSQL = GetBudVsProftandloss(cmbFinancialYear.Column(0), txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom1, periodTo, reportingDate, sessionID1, 2)
            adoconn.Execute _
      "INSERT INTO ReportBudVsAE " & _
      "(ReportingDate, SessionID, CLIENTID,NC,tBalance,Budget,Variance) " & _
        szSQL
        
       szSQL = GetBudVsProftandloss(cmbFinancialYear.Column(0), txtClientList.Tag, txtPropertyName.Tag, txtFundName.Tag, periodFrom1, periodTo, reportingDate, sessionID1, 3)
            adoconn.Execute _
      "INSERT INTO ReportBudVsAE " & _
      "(ReportingDate, SessionID, CLIENTID,NC,tBalance,Budget,Variance) " & _
        szSQL
        
        
        
        
   Dim adoRst As New ADODB.Recordset
   
'   Debug.Print szSQL
   If chkExcludeInc.Value = 1 Then
    '   Exclude Income
        Dim rsCheck As New ADODB.Recordset
        rsCheck.Open "SELECT  R.*  from ReportBudVsAE R,NominalLedger N where N.Code=R.NC AND N.ClientID=R.ClientID AND N.sUBType IN ('cc1','cc2','CC1','CC2') AND SessionID='" & sessionID1 & "'", adoconn, adOpenStatic, adLockReadOnly
        While Not rsCheck.EOF
                adoconn.Execute "Delete from ReportBudVsAE where NC='" & rsCheck.Fields("NC").Value & "' AND SessionID='" & sessionID1 & "'"
        rsCheck.MoveNext
        Wend
        rsCheck.Close
        szSQL = "Select R.*,N.Name from ReportBudVsAE R,NominalLedger N where N.Code=R.NC AND N.ClientID=R.ClientID and N.TYPE=2 AND SessionID='" & sessionID1 & "' Order by R.NC" 'N.sUBType noT IN('cc1','cc2','CC1','CC2') and
   Else
        szSQL = "Select R.*,N.Name from ReportBudVsAE R,NominalLedger N where N.Code=R.NC AND N.ClientID=R.ClientID AND SessionID='" & sessionID1 & "' Order by R.NC"
   End If
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockOptimistic

   flxBudget.Rows = adoRst.RecordCount + 1
'   flxBudget.RowHeight(1) = RowHeight
   
   If adoRst.EOF Then
       adoRst.Close
       Set adoRst = Nothing
       txtTotalBudget.text = "0.00"
       txtTotalActual.text = "0.00"
       txtTotalVariance.text = "0.00"
       MsgBox "No records found.", vbInformation, "No records."
       Exit Function
   End If

   Dim i As Integer, j As Integer
   Dim budgetTotal As Double, actualTotal As Double
   Dim currentPnL As Double
   
   budgetTotal = 0
   actualTotal = 0
   currentPnL = 0
   txtTotalBudget.text = "0.00"
   txtTotalActual.text = "0.00"
   txtTotalVariance.text = "0.00"
   
'   Dim retainedEarningsControl As String
'
'   retainedEarningsControl = GetNominalCodeForControlAccount(adoConn, "Retained Earnings", txtClientList.tag)
'
'   For i = 0 To adoRst.RecordCount - 1
'
'      For j = 0 To adoRst.Fields.count - 1
'         flxBudget.TextMatrix(i + 1, j) = IIf(IsNull(adoRst.Fields(j)), "", adoRst.Fields(j))
'
'         If UCase(adoRst.Fields(j).Name) = "BUDGET" Then
'            budgetTotal = budgetTotal + IIf(IsNull(adoRst.Fields(j)), 0, adoRst.Fields(j))
'
''            If adoRst.Fields(0) <> retainedEarningsControl And adoRst.Fields(2) = "2" Then
''                currentPnL = currentPnL + IIf(IsNull(adoRst.Fields(j)), 0, adoRst.Fields(j))
''            End If
'
'         ElseIf UCase(adoRst.Fields(j).Name) = "ACTUAL" Then
'
'            actualTotal = actualTotal + IIf(IsNull(adoRst.Fields(j)), 0, adoRst.Fields(j))
''            If adoRst.Fields(0) <> retainedEarningsControl And adoRst.Fields(2) = "2" Then
''                currentPnL = currentPnL - IIf(IsNull(adoRst.Fields(j)), 0, adoRst.Fields(j))
''            End If
'         End If
'
'
'
'      Next j
'      adoRst.MoveNext
''      If Not adoRst.EOF Then flxBudget.AddItem ""
'   Next i
    Dim iKount As Integer
    iKount = 1
    With flxBudget
      While Not adoRst.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoRst.Fields.Item("NC").Value 'Code
         .TextMatrix(iKount, 1) = adoRst.Fields.Item("Name").Value 'Nominal Name
         .TextMatrix(iKount, 2) = Format(adoRst.Fields.Item("tBalance").Value, "0.00") 'Actual
          actualTotal = actualTotal + Val(adoRst.Fields.Item("tBalance").Value)
         .TextMatrix(iKount, 3) = Format(adoRst.Fields.Item("Budget").Value, "0.00") 'Budget
          budgetTotal = budgetTotal + Val(adoRst.Fields.Item("Budget").Value)
         .TextMatrix(iKount, 4) = Format(adoRst.Fields.Item("Variance").Value, "0.00") 'Variance
          currentPnL = currentPnL + Val(adoRst.Fields.Item("Variance").Value)
         adoRst.MoveNext
         iKount = iKount + 1
      Wend
   End With
   txtTotalBudget.text = Format(budgetTotal, "#,##0.00")
   txtTotalActual.text = Format(actualTotal, "#,##0.00")
   txtTotalVariance.text = Format(currentPnL, "#,##0.00")
'
'   'Multiplied by -1 to show the debit figure (expenditure) as negative
'   lblCurrentPnL = Format(currentPnL * -1, "#,##0.00")
'
'   If currentPnL * -1 > 0 Then
'        lblCurrentPnL.ForeColor = vbBlue
'   Else
'       lblCurrentPnL.ForeColor = vbRed
'   End If
       
   adoRst.Close
   Set adoRst = Nothing

   Exit Function

NewNominalCode:
   Set adoRst = Nothing
   MsgBox Err.Number & " " & Err.description, vbExclamation + vbOKOnly, "Generating Nominal Balances"
   
End Function

Private Sub Form_Load()
    Dim szSQL As String
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 14055
   Me.Height = 7830
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   chkYtD.BackColor = MODULEBACKCOLOR
   chkExcludeInc.BackColor = MODULEBACKCOLOR
   ConfigFlxBudget

   Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
   adoconn.Open getConnectionString

   'LoadFlxBudget adoConn
   'added by anol 09 08 2016
   createTable adoconn
'   LoadCboClient adoConn, cboClient
   
'   If cboClient.ListCount > 0 Then cboClient.ListIndex = 0

szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

  


    If Not adoRst.EOF Then
                txtClientList.Tag = adoRst.Fields("CLIENTID").Value
                txtClientList.text = adoRst.Fields("CLIENTNAME").Value
                adoRst.Close
   End If
   LoadCmbFinancialYear adoconn
'   LoadCmbFunds adoConn, cboFund
'
'   If cboFund.ListCount > 0 Then cboFund.ListIndex = 0

'   UpdateBudgetBalance adoConn
'   UpdateActualBalance adoConn
   
'   chkYtD.Value = True
   adoconn.Close
   Set adoconn = Nothing

'   UpdateVariance
   Call WheelHook(Me.hWnd)
End Sub

'Private Function LoadFlxBudget(adoConn As ADODB.Connection) As Boolean
'   Dim szSQL As String
'   Dim adoRst As New ADODB.Recordset
'
'   LoadFlxBudget = True
'   On Error GoTo NewNominalCode
'
''  Check: has the client id been setup with NC?
'   adoRst.Open "SELECT ClientID FROM NominalLedger;", adoConn, adOpenStatic, adLockReadOnly
'   adoRst.Close
'
''  ClientID column has been setup
'   szSQL = "SELECT N.Code, N.Name, '', '', '', N.ClientID " & _
'           "FROM NominalLedger AS N  " & _
'           "WHERE N.ClientID <> 'NONE' AND N.Type = 2 " & _
'           "ORDER BY N.Code;"
''Debug.Print szSQL
'   populateGridDefinedHeader adoConn, szSQL, flxBudget, 0
'
'   Set adoRst = Nothing
'   Exit Function
'
'NewNominalCode:
'   Set adoRst = Nothing
'   LoadFlxBudget = False
'End Function

'Private Sub UpdateBudgetBalance(adoConn As ADODB.Connection)
'   Dim adoRst     As New ADODB.Recordset
'   Dim szSQL      As String
'   Dim iRow       As Integer
'   Dim cTotal     As Currency
'
'   On Error GoTo ErrorHandler
'
'   szSQL = "SELECT C.NC, C.BudgetAmt " & _
'           "FROM GlobalSC AS P, GlobalSCDtls AS C, NominalLedger AS N " & _
'           "WHERE P.BudgetID = C.BudgetID AND " & _
'                 "P.PropertyID = '" & txtPropertyName.tag & "' AND " & _
'                 "P.Fund = " & Val(txtFundName.TAG) & " AND " & _
'                 "C.NC = N.Code AND N.Type = 2 AND " & _
'                 "N.ClientID = '" & txtClientList.tag & "';"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'   cTotal = 0
'   While Not adoRst.EOF
'      For iRow = 1 To flxBudget.Rows - 1
'         If flxBudget.RowHeight(iRow) = 240 And flxBudget.TextMatrix(iRow, 0) = adoRst.Fields.Item("NC").Value And Not IsNull(adoRst.Fields.Item("BudgetAmt").Value) Then
'            If flxBudget.TextMatrix(iRow, 3) = "" Then
'               flxBudget.TextMatrix(iRow, 3) = Format(adoRst.Fields.Item("BudgetAmt").Value, "0.00")
'            Else
'               flxBudget.TextMatrix(iRow, 3) = Format(Val(flxBudget.TextMatrix(iRow, 3)) + Val(adoRst.Fields.Item("BudgetAmt").Value), "0.00")
'            End If
'            cTotal = cTotal + Val(adoRst.Fields.Item("BudgetAmt").Value)
'         End If
'      Next iRow
'      adoRst.MoveNext
'   Wend
'
'   txtTotalBudget.text = Format(cTotal, "0.00")
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   Set adoRst = Nothing
'End Sub

'Private Sub UpdateActualBalance(adoConn As ADODB.Connection)
'   Dim iRow       As Integer
'   Dim cBalDr     As Currency
'   Dim cBalCr     As Currency
'   Dim cTotal     As Currency
'
'   For iRow = 1 To flxBudget.Rows - 1
'      If flxBudget.RowHeight(iRow) = 240 And flxBudget.TextMatrix(iRow, 0) <> "" Then
'         cBalDr = CalcuateBalanceDr_PnL(flxBudget.TextMatrix(iRow, 0), txtClientList.tag, adoConn)
'         cBalCr = CalcuateBalanceCr_PnL(flxBudget.TextMatrix(iRow, 0), txtClientList.tag, adoConn)
'
'         If cBalDr - cBalCr <> 0 Then
'            flxBudget.TextMatrix(iRow, 2) = Format(cBalDr - cBalCr, "0.00")
'            cTotal = cTotal + Val(flxBudget.TextMatrix(iRow, 2))
'         End If
'      End If
'   Next iRow
'
'   txtTotalActual.text = Format(cTotal, "0.00")
'End Sub
'
'Private Sub UpdateVariance()
'   Dim iRow       As Integer
'
'   For iRow = 1 To flxBudget.Rows - 1
'      If flxBudget.RowHeight(iRow) = 240 And flxBudget.TextMatrix(iRow, 0) <> "" And (flxBudget.TextMatrix(iRow, 3) <> "" Or flxBudget.TextMatrix(iRow, 2) <> "") Then
'         flxBudget.TextMatrix(iRow, 4) = Format(Val(flxBudget.TextMatrix(iRow, 3)) - Val(flxBudget.TextMatrix(iRow, 2)), "0.00")
'      End If
'   Next iRow
'
'   txtTotalVariance.text = Format(Val(txtTotalBudget.text) - Val(txtTotalActual.text), "0.00")
'End Sub

Private Function LoadCboClient(adoconn As ADODB.Connection, cboC As Control) As Boolean
 
 Dim adoRst As New ADODB.Recordset
 Dim szSQL As String

   LoadCboClient = False
   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT C.CLIENTID, C.CLIENTNAME, S.Code AS NCode " & _
           "FROM CLIENT AS C LEFT JOIN NominalLedger AS S ON C.ClientID = S.ClientID " & _
           "Where S.CAType = 'R' " & _
           "ORDER BY CLIENTNAME;"

   szSQL = "SELECT C.CLIENTID, C.CLIENTNAME, Q.Code AS NCode " & _
           "FROM CLIENT AS C LEFT JOIN (" & _
               "SELECT ClientID, Code " & _
               "FROM NominalLedger AS N " & _
               "WHERE N.CAType = 'R'" & _
              ") AS Q ON C.ClientID = Q.ClientID " & _
           "ORDER BY C.CLIENTID;"


   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

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

   LoadCboClient = True
   Exit Function

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Function

'Private Sub LoadFund(adoConn As ADODB.Connection)
'   Dim rRow As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT FundID, FundName FROM Fund WHERE CategoryCode = 2;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "A fund has not been created for this company.", vbExclamation, "Load Fund in Global"
'   Else
'      ReDim Data(2, adoRst.RecordCount - 1) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
'         Data(1, rRow) = adoRst.Fields.Item("FundName").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'      cboFund.Clear
'      cboFund.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Private Function LoadCmbFunds(adoconn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbFunds = False
   On Error GoTo ErrorHandler

   szSQL = "SELECT FundID, FundCode " & _
           "FROM Fund WHERE CategoryCode = 2 ORDER BY FundCode;"


   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

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
   ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Function LoadCmbProperties(adoconn As ADODB.Connection, cboC As Control) As Boolean
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   LoadCmbProperties = False
   On Error GoTo ErrorHandler

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property WHERE CLIENTID = '" & txtClientList.Tag & "' " & _
           "ORDER BY PropertyID;"


   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   Dim Data() As String

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
   
   cboC.Column() = Data()
   cboC.ListIndex = 0
   
   LoadCmbProperties = True
   Exit Function

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   ShowMsgInTaskBar "Nominal Ledger will not be loaded, as no client has been setup", "Y", "N"

   Exit Function

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub LoadCmbFinancialYear(adoconn As ADODB.Connection)
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
           "ORDER BY FY_StDate Desc;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

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
               dtStartPnL = CDate(adoRst.Fields("FY_StDate").Value)
               dtStartBS = CDate("01 January 2000")
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
   ShowMsgInTaskBar "Financial year has not been created for the client", "Y", "N"
   Exit Sub
End Sub

Private Sub ConfigFlxBudget()
   Dim szHeader      As String
   Dim iCol          As Integer
   Dim iLeft         As Integer

   flxBudget.Clear
   flxBudget.Cols = 6
   flxBudget.Rows = 2
   flxBudget.RowHeight(0) = 0
   szHeader$ = "<Code|<Name|>Actual|>Budget|>Variance|ClientID"

   flxBudget.FormatString = szHeader$

   flxBudget.ColWidth(0) = Label2(2).Left - Label2(1).Left      'Code
   flxBudget.ColWidth(1) = Label2(3).Left - Label2(2).Left      'Nominal Name
   flxBudget.ColWidth(2) = Label2(4).Left - Label2(3).Left      'Actual
   flxBudget.ColWidth(3) = Label2(5).Left - Label2(4).Left      'Budget
   flxBudget.ColWidth(4) = flxBudget.Width + flxBudget.Left - Label2(5).Left - 300     'Variance
   flxBudget.ColWidth(5) = 0

   txtTotalActual.Width = flxBudget.ColWidth(2)
   txtTotalBudget.Width = flxBudget.ColWidth(3)
   txtTotalVariance.Width = flxBudget.ColWidth(4)

   iLeft = flxBudget.Left
   iLeft = iLeft + flxBudget.ColWidth(0)        'NC
   iLeft = iLeft + flxBudget.ColWidth(1)        'NN
   txtTotalActual.Left = iLeft
   iLeft = iLeft + flxBudget.ColWidth(2)        'B
   txtTotalBudget.Left = iLeft
   iLeft = iLeft + flxBudget.ColWidth(3)        'V
   txtTotalVariance.Left = iLeft
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
    UnLoadForm Me
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
         ' PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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

