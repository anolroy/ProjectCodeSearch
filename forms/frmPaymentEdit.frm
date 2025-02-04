VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPaymentEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Payment"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16065
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaymentEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   16065
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   8010
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   41
      Top             =   0
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
         TabIndex        =   45
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   44
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   49
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   48
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   90
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
         Left            =   1665
         TabIndex        =   46
         Top             =   90
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         Height          =   280
         Index           =   15
         Left            =   45
         Top             =   50
         Width           =   5850
      End
   End
   Begin VB.Frame fraEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3660
      Left            =   720
      TabIndex        =   29
      Top             =   90
      Width           =   5940
      Begin VB.TextBox txtProperty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   52
         Top             =   495
         Width           =   4095
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
         Height          =   255
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   525
         Width           =   280
      End
      Begin VB.CommandButton cmdClient 
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
         Height          =   255
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   280
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3285
         Width           =   1290
      End
      Begin VB.CommandButton cmdUpate 
         Caption         =   "&OK"
         Height          =   300
         Left            =   3210
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3285
         Width           =   1335
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1245
         Width           =   2655
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   5
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   2250
         Width           =   2655
      End
      Begin VB.CommandButton cmdSupplier 
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
         Height          =   255
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   900
         Width           =   280
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   8
         Top             =   2955
         Width           =   4455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   51
         Top             =   540
         Width           =   615
      End
      Begin MSForms.ComboBox txtSPSupplier 
         Height          =   315
         Left            =   1440
         TabIndex        =   50
         Top             =   855
         Width           =   4455
         VariousPropertyBits=   1753237535
         DisplayStyle    =   3
         Size            =   "7858;556"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin MSForms.ComboBox txtClient 
         Height          =   315
         Left            =   1440
         TabIndex        =   40
         Top             =   90
         Width           =   4455
         VariousPropertyBits=   1753237535
         DisplayStyle    =   3
         Size            =   "7858;556"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   39
         Top             =   135
         Width           =   435
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   4140
         TabIndex        =   38
         Top             =   1260
         Width           =   210
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "379;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   2250
         Width           =   735
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   1590
         Width           =   360
      End
      Begin MSForms.ComboBox cmbFund 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1590
         Width           =   4455
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7858;503"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   2955
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2580
         Width           =   1095
      End
      Begin MSForms.ComboBox cboBC 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   2580
         Width           =   4455
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "7858;556"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;3527"
      End
      Begin MSForms.ComboBox cmbSPAmtType 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   3285
         Width           =   1710
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3016;556"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;3527"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   3285
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   2240
      TabIndex        =   13
      Top             =   5850
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   3915
      TabIndex        =   14
      Top             =   5850
      Width           =   1335
   End
   Begin VB.PictureBox picSupList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   3030
      Left            =   8010
      ScaleHeight     =   3000
      ScaleWidth      =   5880
      TabIndex        =   20
      Top             =   4905
      Visible         =   0   'False
      Width           =   5910
      Begin VB.TextBox txtSupplierSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1455
         TabIndex        =   17
         Top             =   270
         Width           =   4335
      End
      Begin VB.TextBox txtSupplierSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   1260
      End
      Begin VB.CommandButton cmdGridUnitClose2 
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
         Index           =   1
         Left            =   5610
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   15
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierList 
         Height          =   2370
         Left            =   45
         TabIndex        =   18
         Top             =   600
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   195
         Index           =   2
         Left            =   1530
         TabIndex        =   21
         Top             =   60
         Width           =   1035
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   45
         Top             =   60
         Width           =   5565
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   615
      TabIndex        =   12
      Top             =   5850
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cl&ose"
      Height          =   495
      Left            =   6255
      TabIndex        =   15
      Top             =   5850
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPaymentSplit 
      Height          =   1665
      Left            =   120
      TabIndex        =   23
      Top             =   4125
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2937
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   16777215
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   28
      Top             =   3870
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   240
      TabIndex        =   27
      Top             =   3870
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   26
      Top             =   3870
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5400
      TabIndex        =   25
      Top             =   3870
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   24
      Top             =   3870
      Width           =   555
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   120
      Top             =   3810
      Width           =   7695
   End
End
Attribute VB_Name = "frmPaymentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sTextBox As String
Private iRecords As Integer
Public TransactionID As Long
Public InvoiceNO As String
Public BoolReconciled As Boolean
'Public isEditPPR As Boolean
'Private Sub PrepareCboClient(adoConn As ADODB.Connection)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count - 1
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow - 1
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClient.Column() = Data()
'   'cboClient.ListIndex = -1
'   adoRst.Close
'
'NoRes:
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub cboClient_Change()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call LoadBankAccountInComboClient(adoconn)
    adoconn.Close
    
End Sub





Private Sub cboBC_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
       FocusControl txtReference
    End If
End Sub

Private Sub cmbFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtDetails
    End If
End Sub



Private Sub cmbSPAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
          FocusControl cmdUpate
    End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frmPurchaseExpense.Enabled = True
End Sub

Private Sub cmdCancel1_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    ConfigFlxPaymentSplit
    LoadFlxPaymentSplit adoconn
    adoconn.Close
    cmdUpate.Enabled = True
End Sub

Private Sub cmdDelete_Click()
   'Modified by anol 30 Aug 2015
   'issue 571 note 1152
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim dAmt As Double
   Dim iSelection As Integer
   If txtAmount.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and The line cannot be deleted", "Y", "N"
      Exit Sub
   End If
   If frmPurchaseExpense.tabPurExp.Tab = 1 Then
           dAmt = flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 6)
           If iRecords < 2 Then 'if there is only one split line
             If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                adoconn.Open getConnectionString
                adoconn.BeginTrans
                ShowMsgInTaskBar "It is not possible to delete the last line of a purchase payment. The payment amount will be automatically set to zero.", "Y", "N"
                adoconn.Execute "Update  tlbPaymentSplit set amount=0 ,osamount=0 WHERE TransactionID = '" & flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 0) & "';"
                
                 If frmPurchaseExpense.sEditPPR = 1 Then 'when clicking Upper Grid on PI form
                    'type 24
                     adoconn.Execute "UPDATE tlbPayment SET    Amount   = 0, OSAmount = 0, NLPost=false " & _
                        "WHERE TransactionID = " & TransactionID & ";"
                     adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                        "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                        " and N.TRANSACTION_TYPE = 24 "
                        
                        
'                        UpdateBankAcBal_Plus adoConn, dAmt, frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 21), txtClient.text
'                        'adding to the seleted bank on combo because this is refund (and for delete this is reverse)
'                        UpdateBankAcBal_Minus adoConn, dAmt, cboBC.Value, txtClient.text
                        
                End If
                If frmPurchaseExpense.sEditPPR = 2 Then 'when clicking lower Grid on PI form
                    'type is 8,9
                     adoconn.Execute "UPDATE tlbPayment SET    Amount   = 0, OSAmount = 0, NLPost=false " & _
                         "WHERE TransactionID = " & TransactionID & ";"
                     adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                         "and (N.TRANSACTION_TYPE = 8 or N.TRANSACTION_TYPE=9)"
                     
                             'this is just opposite in the save
'                          UpdateBankAcBal_Minus adoConn, dAmt, _
'                                frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 15), txtClient.text
'                          UpdateBankAcBal_Plus adoConn, dAmt, cboBC.Value, txtClient.text
                     
               End If
              
                
                If SiPi_Check1(adoconn, "PI") = False Then
                     adoconn.RollbackTrans
                     adoconn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit."
                     Exit Sub
                Else
                    adoconn.CommitTrans
                    LoadFlxPaymentSplit adoconn
                    If frmPurchaseExpense.sEditPPR = 2 Then
                         frmPurchaseExpense.LoadFlxSCrPoA adoconn
                    Else
                         frmPurchaseExpense.LoadFlxSPayment adoconn
                    End If
                    adoconn.Close
                    cmdDelete.Enabled = False
                    cmdSave.Enabled = False
                    txtAmount.text = "0.00"
        '            flxPaymentSplit.TextMatrix(1, 6) = "0.00"
                End If
                adoconn.Open getConnectionString
                Export_PPnPPR_2_NL adoconn
                HighLightRowFlxGrid flxPaymentSplit, 1
                adoconn.Close
                Set adoconn = Nothing
                frmPurchaseExpense.cmdEditPayment.Enabled = False
                Exit Sub
           End If
          'now code start for where split line is two or more
           If flxPaymentSplit.row <= 0 Then
              ShowMsgInTaskBar "Select a line to delete."
              Exit Sub
           End If
           iSelection = flxPaymentSplit.row
           If MsgBox("Do you wish to delete the line?", vbQuestion + vbYesNo, "Deleting a line") = vbNo Then Exit Sub
        
        
           adoconn.Open getConnectionString
           adoconn.BeginTrans
           szSQL = "DELETE * " & _
                   "FROM tlbPaymentSplit " & _
                   "WHERE TransactionID = '" & _
                       flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 0) & "';"
           adoconn.Execute szSQL
           
           If frmPurchaseExpense.sEditPPR = 1 Then 'upper grid
             'type 24
                adoconn.Execute "UPDATE tlbPayment " & _
                        "SET    Amount   = Amount   - " & Val(txtAmount.text) & ", " & _
                               "OSAmount = OSAmount - " & Val(txtAmount.text) & ", NLPost=false " & _
                        "WHERE TransactionID = " & _
                               TransactionID & ";"
              adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                        "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                        " and N.TRANSACTION_TYPE = 24 "
'              UpdateBankAcBal_Plus adoConn, dAmt, frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 21), txtClient.text
'                        'adding to the seleted bank on combo because this is refund (and for delete this is reverse)
'              UpdateBankAcBal_Minus adoConn, dAmt, cboBC.Value, txtClient.text
         
           End If
           
           If frmPurchaseExpense.sEditPPR = 2 Then 'lower grid
           'type 8,9
                adoconn.Execute "UPDATE tlbPayment " & _
                        "SET    Amount   = Amount   - " & Val(txtAmount.text) & ", " & _
                               "OSAmount = OSAmount - " & Val(txtAmount.text) & ", NLPost=false " & _
                        "WHERE TransactionID = " & _
                               TransactionID & ";"
                              
                adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                         "and (N.TRANSACTION_TYPE = 8 or N.TRANSACTION_TYPE=9)"
                         
                 'this is just opposite in the save
'                          UpdateBankAcBal_Minus adoConn, dAmt, _
'                                frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 15), txtClient.text
'                          UpdateBankAcBal_Plus adoConn, dAmt, cboBC.Value, txtClient.text
            End If
           
            
           
          
           
                         
           If SiPi_Check1(adoconn, "PI") = False Then
                adoconn.RollbackTrans
                adoconn.Close
                MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit."
                Exit Sub
           Else
                adoconn.CommitTrans
                dAmt = txtAmount.text ' this will hold the amount which is currently deleting. After loading the amount will berestting to current row
                LoadFlxPaymentSplit adoconn
                'Below part I don't need because grid is loading amounts and osamount from database
                If frmPurchaseExpense.sEditPPR = 2 Then 'reducing the amout in header  in the grid
                    frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 7) = Format(Val(frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 7)) - Val(dAmt), "0.00")
                    frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 8) = Format(Val(frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 8)) - Val(dAmt), "0.00")
                Else
                    frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 8) = Format(Val(frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 8)) - Val(dAmt), "0.00")
                    frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 9) = Format(Val(frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 9)) - Val(dAmt), "0.00")
                End If
                adoconn.Close
           End If
           
           
           adoconn.Open getConnectionString
           Export_PPnPPR_2_NL adoconn
            
           ShowMsgInTaskBar "The line has been deleted.", "Y", "P"
        
           
           If iSelection <= iRecords Then
                flxPaymentSplit.row = iSelection
                HighLightRowFlxGrid flxPaymentSplit, flxPaymentSplit.row
           Else
                flxPaymentSplit.row = iRecords
                HighLightRowFlxGrid flxPaymentSplit, iRecords
           End If
           txtAmount.text = Format(flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 6), "0.00")
           adoconn.Close
           Set adoconn = Nothing
           frmPurchaseExpense.cmdEditPayment.Enabled = False
   End If
   If frmPurchaseExpense.tabPurExp.Tab = 3 Then
        If iRecords < 2 Then 'if there is only one split line
             If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                adoconn.Open getConnectionString
                adoconn.BeginTrans
                ShowMsgInTaskBar "It is not possible to delete the last line of a purchase payment. The payment amount will be automatically set to zero.", "Y", "N"
                adoconn.Execute "Update  tlbPaymentSplit set amount=0 ,osamount=0 WHERE TransactionID = '" & flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 0) & "';"
            
               'type is 8,9,24
                adoconn.Execute "UPDATE tlbPayment SET    Amount   = 0, OSAmount = 0, NLPost=false " & _
                    "WHERE TransactionID = " & TransactionID & ";"
                adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                    "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                    "and (N.TRANSACTION_TYPE = 8 or N.TRANSACTION_TYPE=9 OR N.TRANSACTION_TYPE = 24)"
              
              
                
                If SiPi_Check1(adoconn, "PI") = False Then
                     adoconn.RollbackTrans
                     adoconn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit."
                     Exit Sub
                Else
                    adoconn.CommitTrans
                    frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 9) = "0.00"
                    frmPurchaseExpense.LoadFlxSPayment adoconn
                    frmPurchaseExpense.LoadFlxSCrPoA adoconn
                    'LoadflxPaymentSplit adoconn
                    LoadFlxPaymentSplit adoconn
                    adoconn.Close
                    cmdDelete.Enabled = False
                    cmdSave.Enabled = False
                    txtAmount.text = "0.00"
        '            flxPaymentSplit.TextMatrix(1, 6) = "0.00"
                End If
                adoconn.Open getConnectionString
                Export_PPnPPR_2_NL adoconn
                HighLightRowFlxGrid flxPaymentSplit, 1
                adoconn.Close
                Set adoconn = Nothing
                Exit Sub
           End If
          'now code start for where split line is two or more
           If flxPaymentSplit.row <= 0 Then
              ShowMsgInTaskBar "Select a line to delete."
              Exit Sub
           End If
           iSelection = flxPaymentSplit.row
           If MsgBox("Do you wish to delete the line?", vbQuestion + vbYesNo, "Deleting a line") = vbNo Then Exit Sub
        
        
           adoconn.Open getConnectionString
           adoconn.BeginTrans
           szSQL = "DELETE * " & _
                   "FROM tlbPaymentSplit " & _
                   "WHERE TransactionID = '" & _
                       flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 0) & "';"
           adoconn.Execute szSQL
             'type 8,9,24
                adoconn.Execute "UPDATE tlbPayment " & _
                        "SET    Amount   = Amount   - " & Val(txtAmount.text) & ", " & _
                               "OSAmount = OSAmount - " & Val(txtAmount.text) & ", NLPost=false " & _
                        "WHERE TransactionID = " & _
                               TransactionID & ";"
                adoconn.Execute "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                        "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                        " and (N.TRANSACTION_TYPE = 8 or N.TRANSACTION_TYPE=9 OR N.TRANSACTION_TYPE = 24) "
           
                         
           If SiPi_Check1(adoconn, "PI") = False Then
                adoconn.RollbackTrans
                adoconn.Close
                MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit."
                Exit Sub
           Else
                adoconn.CommitTrans
                'LoadflxPaymentSplit adoconn
                LoadFlxPaymentSplit adoconn
                frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 9) = Format(frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 9) - Val(txtAmount.text), "0.00")
                frmPurchaseExpense.LoadFlxSPayment adoconn
                frmPurchaseExpense.LoadFlxSCrPoA adoconn
                adoconn.Close
           End If
           
           
           adoconn.Open getConnectionString
           Export_PPnPPR_2_NL adoconn
            
           ShowMsgInTaskBar "The line has been deleted.", "Y", "P"
        
           
           If iSelection <= iRecords Then
                flxPaymentSplit.row = iSelection
                HighLightRowFlxGrid flxPaymentSplit, flxPaymentSplit.row
           Else
                flxPaymentSplit.row = iRecords
                HighLightRowFlxGrid flxPaymentSplit, iRecords
           End If
           txtAmount.text = Format(flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 6), "0.00")
           adoconn.Close
           Set adoconn = Nothing
   End If
End Sub



Private Sub cmdEdit_Click()
   cmdSave.Enabled = False

   If flxPaymentSplit.row <= 0 Then
      ShowMsgInTaskBar "Select a split to edit."
      Exit Sub
   End If

   If cmdEdit.Caption = "&Cancel" Then
'      fraEdit.Enabled = False
      flxPaymentSplit.Enabled = True
      cmdSave.Enabled = False
      cmdEdit.Caption = "&Edit"
      Exit Sub
   End If

   fraEdit.Enabled = True
   flxPaymentSplit.Enabled = False

   cmdEdit.Caption = "&Cancel"
End Sub

Private Sub cmdGridUnitClose2_Click(Index As Integer)
   picSupList.Visible = False
End Sub

Private Sub cmdPicCLose_Click()
     picClient.Visible = False
     fraEdit.Enabled = True
End Sub

Private Function SiPi_Check1(adoconn As ADODB.Connection, szSiPi As String) As Boolean
   Dim szSQL      As String
   Dim adoRst     As New ADODB.Recordset
   Dim szTran2Fix As String

   If szSiPi = "PI" Then
      szSQL = "SELECT  P.TransactionID " & _
               "FROM tlbPayment AS P, (" & _
                     "SELECT PayHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                     "From tlbPaymentSplit " & _
                     "Group by PayHeader " & _
                     ") AS Q " & _
               "WHERE P.TransactionID = Q.PayHeader AND P.Amount <> P.OSAmount AND " & _
                     "ROUND(P.Amount - P.OSAmount, 2) <> Q.T;"
'Debug.Print szSQL
      adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)

         adoRst.MoveNext
      Wend

      adoRst.Close
   End If

   If szSiPi = "SI" Then
      szSQL = "SELECT  R.TransactionID " & _
               "FROM tlbReceipt AS R, (" & _
                     "SELECT RptHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                     "From tlbReceiptSplit " & _
                     "Group by RptHeader " & _
                     ") AS Q " & _
               "WHERE R.TransactionID = Q.RptHeader AND R.Amount <> R.OSAmount AND " & _
                     "ROUND(R.Amount - R.OSAmount, 2) <> Q.T;"
'Debug.Print szSQL
      adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)

         adoRst.MoveNext
      Wend

      adoRst.Close
   End If

   Set adoRst = Nothing

   If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)

   If Len(szTran2Fix) > 0 Then
        SiPi_Check1 = False
   Else
        SiPi_Check1 = True
        'MsgBox "HI"
   End If
      
End Function

Private Sub cmdProperty_Click()
    picClient.Left = 880
    picClient.Top = 45
    sTextBox = "2"
    LoadflxProperty
'    fraDemandType.Enabled = False
'    fraCommands.Enabled = False
    picClient.Visible = True
'    fraEdit.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdSave_Click()
'In this procedure we are not creating new SPlit for payment
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, dAmt As Double, i As Integer
    If lblPostingDate.ToolTipText = "" Then
        MsgBox "Posting date not found"
        Exit Sub
    End If
'    If IsNull(cboClient.Column(0)) Then
'        MsgBox "Please select a proper client"
'        Exit Sub
'    End If
    
    If txtClient.text = "" Then
        MsgBox "Please select a valid client"
        Exit Sub
    End If
    If DateDiff("d", lblPostingDate.ToolTipText, txtDate.text) > 0 Then
        MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
        Exit Sub
    End If
    If frmPurchaseExpense.tabPurExp.Tab = 1 Then
                adoconn.Open getConnectionString
                If frmPurchaseExpense.sEditPPR = 1 Then 'UPPER GRID
                '----------------------------------------------------Type 24--------------------------------
                   adoconn.BeginTrans 'implemented by anol 20170424 issue 357
                   szSQL = "SELECT * " & _
                           "FROM tlbPaymentSplit " & _
                           "WHERE PayHeader = " & _
                               TransactionID & ";"
                   'Here PayHeader is a foreign key of transaction ID from tlbpayment table
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                
                   dAmt = 0
                   For i = 1 To flxPaymentSplit.Rows - 1
                            If adoRst.EOF Then
                                adoRst.AddNew
                                adoRst.Fields.Item("PayHeader").Value = Trim(TransactionID)
                                adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                            End If
                      adoRst.Fields.Item("FundID").Value = flxPaymentSplit.TextMatrix(i, 1)
                      adoRst.Fields.Item("Amount").Value = flxPaymentSplit.TextMatrix(i, 6)
                      adoRst.Fields.Item("OSAmount").Value = flxPaymentSplit.TextMatrix(i, 6)
                      dAmt = dAmt + flxPaymentSplit.TextMatrix(i, 6)
                      adoRst.Fields.Item("Description").Value = flxPaymentSplit.TextMatrix(i, 4)
                      adoRst.Fields.Item("Nominal_Code").Value = flxPaymentSplit.TextMatrix(i, 11)
                      adoRst.Update
                   Next i
                   adoRst.Close
                
                '  updating the header record
                   szSQL = "SELECT * " & _
                           "FROM tlbPayment " & _
                           "WHERE TransactionID = " & _
                               TransactionID & ";"
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                
                 '   Issue: 0000571 note 1148. Modified by anol
                   adoconn.Execute "UPDATE NLPosting AS N " & _
                            "SET    N.DeleteFlag = TRUE " & _
                            "WHERE  N.PARENT_RECORD = '" & TransactionID & "' AND " & _
                            "N.TRANSACTION_TYPE = 24 "
                     'end of modification
                            
                   With adoRst.Fields
                      .Item("SageAccountNumber").Value = txtSPSupplier.text
                      .Item("BankCode").Value = cboBC.Value
                      'issue 784 solved by anol 2019-06-10
                      .Item("NominalCode").Value = cboBC.Value
                      .Item("PayAmtType").Value = cmbSPAmtType.Value ''RptAmtType
                      .Item("Amount").Value = txtAmount.text
                      .Item("OSAmount").Value = .Item("Amount").Value
                      .Item("PDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                      .Item("ExtRef").Value = txtReference.text
                      'added by anol 23 aug 2015
                      'PRESTIGE VALIDATION ALL MODULES AND FORMS Note 1145
                       .Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                       
                      .Item("FundID").Value = cmbFund.Column(0)
                      .Item("Details").Value = txtDetails.text
                       '  Issue: 0000571 note 1148. Modified by anol
                      .Item("NLPost").Value = False
                      .Item("ClientID").Value = txtClient.text 'cboClient.Column(0)
                      .Item("UnitID").Value = txtProperty.text 'it is property ID
                      .Item("LastModifiedBy").Value = User
                      .Item("LastModifiedDate").Value = Now
                      'end of addition
                   End With
                   adoRst.Update
                   adoRst.Close
                
                   Set adoRst = Nothing
                   'check consistency
                   If SiPi_Check1(adoconn, "PI") = False Then
                        adoconn.RollbackTrans
                        adoconn.Close
                        MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit."
                        Exit Sub
                   Else
                        adoconn.CommitTrans
                   End If
                   
'                   If cboBC.Value <> frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 21) Or _
'                         Val(txtAmount.text) <> dAmt Then
'                      UpdateBankAcBal_Minus adoConn, dAmt, frmPurchaseExpense.flxSPayment.TextMatrix(frmPurchaseExpense.flxSPayment.row, 21), txtClient.text
'                      'adding to the seleted bank on combo because this is refund
'                      UpdateBankAcBal_Plus adoConn, dAmt, cboBC.Value, txtClient.text
'                   End If
                
                   cmdEdit.Caption = "&Edit"
                   frmPurchaseExpense.AfterEditPayment adoconn 'refresh flxsPaymengGrid
                   cmdCancel_Click
                   adoconn.Close
                   Set adoconn = Nothing
                   
                   adoconn.Open getConnectionString
                   Export_PPnPPR_2_NL adoconn
                   adoconn.Close
                   frmPurchaseExpense.cmdEditPayment.Enabled = False
           End If
           If frmPurchaseExpense.sEditPPR = 2 Then 'LOWER GRID PP AND PA
        '-------------------------------------------Type 8,9-----------------------------------------------
                   adoconn.BeginTrans
                '  updating all splits FIRST
                   szSQL = "SELECT * " & _
                           "FROM tlbPaymentSplit " & _
                           "WHERE PayHeader = " & _
                               TransactionID & ";"
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                
                   dAmt = 0
                   'Because PA can have only one split
                   For i = 1 To flxPaymentSplit.Rows - 1
                   
                      If Left(InvoiceNO, 2) <> "PA" Then _
                         adoRst.Find "TransactionID = '" & flxPaymentSplit.TextMatrix(i, 0) & "'", , , 1
                    If adoRst.EOF Then
                                adoRst.AddNew
                                adoRst.Fields.Item("PayHeader").Value = Trim(TransactionID)
                                adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                     End If
                      adoRst.Fields.Item("FundID").Value = flxPaymentSplit.TextMatrix(i, 1)
                      adoRst.Fields.Item("Amount").Value = flxPaymentSplit.TextMatrix(i, 6)
                      adoRst.Fields.Item("OSAmount").Value = flxPaymentSplit.TextMatrix(i, 6)
                      dAmt = dAmt + flxPaymentSplit.TextMatrix(i, 6)
                      adoRst.Fields.Item("Description").Value = flxPaymentSplit.TextMatrix(i, 4)
                      adoRst.Fields.Item("Nominal_code").Value = flxPaymentSplit.TextMatrix(i, 11)
                      adoRst.Update
                   Next i
                   adoRst.Close
                   
                
                '  updating the header record
                   szSQL = "SELECT * " & _
                           "FROM tlbPayment " & _
                           "WHERE TransactionID = " & _
                               TransactionID & ";"
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                 '   Issue: 0000571 note 1148. Modified by anol
                   szSQL = "UPDATE NLPosting AS N " & _
                            "SET    N.DeleteFlag = TRUE " & _
                            "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                            " AND (N.TRANSACTION_TYPE = 8 OR  N.TRANSACTION_TYPE = 9)"
                            adoconn.Execute szSQL
                     'end of modification
                   With adoRst.Fields
                      .Item("SageAccountNumber").Value = txtSPSupplier.text
                      .Item("BankCode").Value = cboBC.Value
                      'added by anol issue 345: wrong bank account is being posted to in the nominal ledger 20170419
                      .Item("NominalCode").Value = cboBC.Value
                      .Item("PayAmtType").Value = cmbSPAmtType.Value
                      .Item("Amount").Value = AllSplitTotal
                      .Item("OSAmount").Value = .Item("Amount").Value
                      .Item("PDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                      'added by anol 23 aug 2015
                      'PRESTIGE VALIDATION ALL MODULES AND FORMS Note 1145
                      .Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                      .Item("ExtRef").Value = txtReference.text
                      
                      '  Issue: 0000571 note 1148. Modified by anol 25 Aug 2015
                      .Item("NLPost").Value = False
                      .Item("ClientID").Value = txtClient.text
                      .Item("UnitID").Value = txtProperty.text
                      'end of addition
                      'If Left(InvoiceNO, 2) = "PA" Then
                         .Item("FundID").Value = cmbFund.Column(0)
                         .Item("Details").Value = txtDetails.text
                      'End If
                      .Item("LastModifiedBy").Value = User
                      .Item("LastModifiedDate").Value = Now
                   End With
                   adoRst.Update
                   adoRst.Close
                
                   Set adoRst = Nothing
           'End If
        'check consistency
           If SiPi_Check1(adoconn, "PI") = False Then
                adoconn.RollbackTrans
                adoconn.Close
                MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit !"
                Exit Sub
           Else
                adoconn.CommitTrans
           End If
'          If cboBC.Value <> frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 15) Or _
'                 Val(txtAmount.text) <> dAmt Then
'        'issue 523
'        'Client Id not found for passing in the function
'
''              UpdateBankAcBal_Plus adoConn, dAmt, _
''                    frmPurchaseExpense.flxSCrPoA.TextMatrix(frmPurchaseExpense.flxSCrPoA.row, 15), txtClient.text
''              UpdateBankAcBal_Minus adoConn, dAmt, cboBC.Value, txtClient.text
'           End If
        
           cmdEdit.Caption = "&Edit"
           frmPurchaseExpense.AfterEditPayment adoconn
           cmdCancel_Click
        
           adoconn.Close
           Set adoconn = Nothing
           
           adoconn.Open getConnectionString
           Export_PPnPPR_2_NL adoconn
           adoconn.Close
           frmPurchaseExpense.cmdEditPayment.Enabled = False
        End If 'END IF FORsppredit
   End If 'END IF FOR TAB 1
   If frmPurchaseExpense.tabPurExp.Tab = 3 Then
                    adoconn.Open getConnectionString
                    adoconn.BeginTrans
                '  updating all splits FIRST
                   szSQL = "SELECT * " & _
                           "FROM tlbPaymentSplit " & _
                           "WHERE PayHeader = " & _
                               TransactionID & ";"  'MY_ID/transaction ID
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                
                   dAmt = 0
                   
                   'Because PA can have only one split
                   For i = 1 To flxPaymentSplit.Rows - 1
                        If flxPaymentSplit.TextMatrix(i, 1) <> "" Then 'this is checking fund not null
                             If Left(InvoiceNO, 2) <> "PA" Then _
                                adoRst.Find "TransactionID = '" & flxPaymentSplit.TextMatrix(i, 0) & "'", , , 1
                                
                            'adoRST.Find "TransactionID = " & Val(flxPaymentSplit.TextMatrix(i, 0)) & "", , , 1
                             If adoRst.EOF Then
                                       adoRst.AddNew
                                       adoRst.Fields.Item("PayHeader").Value = Trim(TransactionID)  'MY_ID/transaction ID
                                       adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                                       adoRst.Fields.Item("SplitID").Value = i
                                       'We are not modifiying the split ID, or creating new split ID
                             End If
                             
                             adoRst.Fields.Item("FundID").Value = flxPaymentSplit.TextMatrix(i, 1)
                             adoRst.Fields.Item("Amount").Value = flxPaymentSplit.TextMatrix(i, 6)
                             adoRst.Fields.Item("OSAmount").Value = flxPaymentSplit.TextMatrix(i, 6)
                             dAmt = dAmt + flxPaymentSplit.TextMatrix(i, 6)
                             adoRst.Fields.Item("Description").Value = flxPaymentSplit.TextMatrix(i, 4)
                             adoRst.Update
                        End If
                   Next i
                   adoRst.Close
                   
                
                '  updating the header record
                   szSQL = "SELECT * " & _
                           "FROM tlbPayment " & _
                           "WHERE TransactionID = " & _
                               TransactionID & ";"
                   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                 '   Issue: 0000571 note 1148. Modified by anol
                   szSQL = "UPDATE NLPosting AS N " & _
                            "SET    N.DeleteFlag = TRUE " & _
                            "WHERE  N.PARENT_RECORD = '" & TransactionID & "' " & _
                            " AND (N.TRANSACTION_TYPE = 8 OR  N.TRANSACTION_TYPE = 9 OR N.TRANSACTION_TYPE = 24)"
                            adoconn.Execute szSQL
                     'end of modification
                   With adoRst.Fields
                      .Item("SageAccountNumber").Value = txtSPSupplier.text
                      .Item("BankCode").Value = cboBC.Value
                      'added by anol issue 345: wrong bank account is being posted to in the nominal ledger 20170419
                      .Item("NominalCode").Value = cboBC.Value
                      .Item("PayAmtType").Value = cmbSPAmtType.Value
                      .Item("Amount").Value = AllSplitTotal
                      .Item("OSAmount").Value = .Item("Amount").Value
                      .Item("PDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                      'added by anol 23 aug 2015
                      'PRESTIGE VALIDATION ALL MODULES AND FORMS Note 1145
                      .Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                      .Item("ExtRef").Value = txtReference.text
                      
                      '  Issue: 0000571 note 1148. Modified by anol 25 Aug 2015
                      .Item("NLPost").Value = False
                      .Item("ClientID").Value = txtClient.text
                      .Item("UnitID").Value = txtProperty.text
                      'end of addition
'                      If Left(InvoiceNO, 2) = "PA" Then
                         .Item("FundID").Value = cmbFund.Column(0)
                         .Item("Details").Value = txtDetails.text
'                      End If
                     .Item("LastModifiedBy").Value = User
                     .Item("LastModifiedDate").Value = Now
                   End With
                   adoRst.Update
                   adoRst.Close
                
                   Set adoRst = Nothing
           
        'check consistency
           If SiPi_Check1(adoconn, "PI") = False Then
                adoconn.RollbackTrans
                adoconn.Close
                MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Payment Edit !"
                Exit Sub
           Else
                adoconn.CommitTrans
           End If
'          If cboBC.Value <> frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 15) Or _
'                 Val(txtAmount.text) <> dAmt Then
'            'issue 523
'            'flxPurchPPHistory.row, 15 is bank code. The idea is that if you change the amount or bank code then the bank bank balance is changing accordingly
'            'Client Id not found for passing in the function
'
'              UpdateBankAcBal_Plus adoConn, dAmt, _
'                    frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 15), txtClient.text
'              UpdateBankAcBal_Minus adoConn, dAmt, cboBC.Value, txtClient.text
'           End If
        
           cmdEdit.Caption = "&Edit"
           'frmPurchaseExpense.AfterEditPayment adoconn
           'you need to refresh the grid line here
           
           frmPurchaseExpense.flxPurchPPHistory.TextMatrix(frmPurchaseExpense.flxPurchPPHistory.row, 9) = Format(dAmt, "0.00")
'           If dAmt > 0 Then
'                frmPurchaseExpense.LoadFlxSPayment adoconn
'                frmPurchaseExpense.LoadFlxSCrPoA adoconn
'           End If
           adoconn.Close
           Set adoconn = Nothing
           
           adoconn.Open getConnectionString
           Export_PPnPPR_2_NL adoconn
           adoconn.Close
           
       End If
       cmdSave.Enabled = False
       FocusControl cmdCancel
End Sub

Private Function AllSplitTotal() As Currency
   If flxPaymentSplit.Rows = 2 Then
      AllSplitTotal = CCur(txtAmount.text)
      Exit Function
   End If
   
   Dim iRow As Integer

   For iRow = 1 To flxPaymentSplit.Rows - 1
      AllSplitTotal = AllSplitTotal + CCur(flxPaymentSplit.TextMatrix(iRow, 6))
   Next iRow
End Function



Private Sub cmdSupplier_Click()
   picSupList.Top = 480
   picSupList.Left = 840
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   Call LoadflxSupplierList(adoconn)
   adoconn.Close
   Set adoconn = Nothing
   picSupList.Visible = True
   picSupList.ZOrder 0
   FocusControl txtSupplierSearchID
End Sub
Private Sub LoadflxSupplierList(adoconn As ADODB.Connection)
    Dim adoRst As New ADODB.Recordset
    Dim szSQL As String
    Dim i As String
    If frmPurchaseExpense.txtSupplierType.Tag = "ALL" Then
        szSQL = "SELECT SupplierID, SupplierName FROM Supplier ORDER BY SupplierName;"
   Else
        szSQL = "SELECT SupplierID, SupplierName  FROM Supplier where TYPE='" & frmPurchaseExpense.txtSupplierType.Tag & "'" & _
                "ORDER BY SupplierName;"
   End If
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   flxSupplierList.Rows = adoRst.RecordCount + 1
   i = 1
   While Not adoRst.EOF
      flxSupplierList.TextMatrix(i, 0) = ""
      flxSupplierList.TextMatrix(i, 1) = adoRst.Fields.Item(0).Value
      flxSupplierList.TextMatrix(i, 2) = adoRst.Fields.Item(1).Value
      adoRst.MoveNext
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing
End Sub
Private Sub cmdUpate_Click()
   If cboBC.ListIndex = -1 Then
      ShowMsgInTaskBar "Please select a bank account.", "Y", "N"
      cboBC.SetFocus
      Exit Sub
   End If
   If cmbSPAmtType.text = "" Then
      ShowMsgInTaskBar "Please select the Payment Type.", "Y", "N"
      cmbSPAmtType.SetFocus
      Exit Sub
   End If
   If cmbFund.ListIndex = -1 Then
      cmbFund.Locked = False
      ShowMsgInTaskBar "Please select the fund.", "Y", "N"
      FocusControl cmbFund
      Exit Sub
   End If
   If cmbFund.text = "" Then
      cmbFund.Locked = False
      ShowMsgInTaskBar "Please select the fund.", "Y", "N"
      FocusControl cmbFund
      Exit Sub
   End If
   If txtAmount.text = "" Then
        MsgBox "Please enter amount", vbInformation, "Warning"
        FocusControl txtAmount
        Exit Sub
   End If
   With flxPaymentSplit
       If .row = 0 Then
            .row = 1
        End If
      .TextMatrix(.row, 1) = cmbFund.Column(0)
      .TextMatrix(.row, 3) = cmbFund.text
      .TextMatrix(.row, 4) = txtDetails.text
      .TextMatrix(.row, 5) = txtDate.text
      .TextMatrix(.row, 6) = Format(txtAmount.text, "0.00")
      .TextMatrix(.row, 11) = cboBC.Value
   End With
   'fraEdit.Enabled = false
   fraEdit.Enabled = True
   cmdSave.Enabled = True
   cmdSave.SetFocus
   flxPaymentSplit.Enabled = True
   flxPaymentSplit.row = 0
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub flxClient_Click()
    If sTextBox = "1" Then
            txtClient.text = flxClient.TextMatrix(flxClient.row, 0)
            cboClient_Change
    Else
            txtProperty.text = flxClient.TextMatrix(flxClient.row, 0)
    End If
    fraEdit.Enabled = True
    picClient.Visible = False
     
End Sub



Private Sub flxPaymentSplit_Click()
    If flxPaymentSplit.Rows > 1 Then
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub flxPaymentSplit_RowColChange()
   If flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 0) = "" Then
      flxPaymentSplit.row = flxPaymentSplit.row - 1
   End If

   HighLightRowFlxGrid flxPaymentSplit, flxPaymentSplit.row

   cmbFund.ListIndex = FindComboIndex(cmbFund, flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 1), 0)
   txtDetails.text = flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 4)
   txtAmount.text = Format(flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 6), "0.00")
   txtProperty.text = flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 10)
   If flxPaymentSplit.TextMatrix(flxPaymentSplit.row, 8) <> "" Then
      txtAmount.Locked = True
   Else
      txtAmount.Locked = False
   End If
    cmdDelete.Enabled = True
   fraEdit.Enabled = True
End Sub

Private Sub flxSupplierList_Click()
   If flxSupplierList.TextMatrix(flxSupplierList.row, 1) = "" Then Exit Sub
   txtSPSupplier.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
   picSupList.Visible = False
   FocusControl txtDate
End Sub
Private Sub cmdClient_Click()
    picClient.Left = 880
    picClient.Top = 45
    sTextBox = "1"
    LoadflxClient
'    fraDemandType.Enabled = False
'    fraCommands.Enabled = False
    picClient.Visible = True
'    fraEdit.Enabled = False
    txtSearchClientID.SetFocus

End Sub

Private Sub txtAmount_Click()
    SelTxtInCtrl txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.text = Format(Val(txtAmount.text), "0.00")
End Sub



Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtAmount
    End If
End Sub



Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmbSPAmtType
    End If
End Sub

Private Sub txtSearchClientID_Change()
        'Updated by anol 22 Dec 2015
   Dim i As Integer
'    txtSearchClientID.text = UCase(txtSearchClientID.text)
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

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        If Len(txtSearchClientID) > 0 Then
            flxClient.SetFocus
        Else
            txtSearchClientName.SetFocus
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

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        flxClient.SetFocus
    End If
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0


   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010

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
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345

   'lblJobName.Visible = False
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly


      If sTextBox = "1" Then
'           flxClient.TextMatrix(1, 0) = "ALL"
'           flxClient.TextMatrix(1, 1) = "All Client"
'           flxClient.TextMatrix(1, 2) = ""
'           flxClient.AddItem ""
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub LoadflxProperty()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0


   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010

   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345

   'lblJobName.Visible = False
   adoconn.Open getConnectionString
   szSQL = "SELECT PropertyID, PropertyName,clientID FROM   Property where clientID='" & txtClient.text & "' ORDER BY PropertyID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly


      If sTextBox = "2" Then
'           flxClient.TextMatrix(1, 0) = "ALL"
'           flxClient.TextMatrix(1, 1) = "All Client"
'           flxClient.TextMatrix(1, 2) = ""
'           flxClient.AddItem ""
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub ConfigFlxPaymentSplit()
   Dim szHeader As String
   
   With flxPaymentSplit
      .Clear
      .FormatString = szHeader$
      .Rows = 2
      .RowHeight(0) = 0
      szHeader$ = "TransactionID|<FundID|<No|<FundName|<Description|<DueDate|>Amount||ReconNow"
'                       0           1     2      3           4          5        6   7    8
      .Cols = 12
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0                                      'No
      .ColWidth(3) = Label2(3).Left - Label2(2).Left        'Fund Name
      .ColWidth(4) = Label2(4).Left - Label2(3).Left        'Description
      .ColWidth(5) = Label2(5).Left - Label2(4).Left        'Due Date
      .ColWidth(6) = .Left + .Width - Label2(5).Left - 300  'Amount
      .ColWidth(7) = 0                                      'Flag for edit
      .ColWidth(8) = 0                                      'Reconciled?
      .ColWidth(9) = 0                                      'posting date
      .ColWidth(10) = 0                                    'Property ID
      .ColWidth(11) = 0                                    'Nominal Code
 End With
End Sub

Private Sub flxSupplierList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplierList_Click
    End If
End Sub

Private Sub Form_Load()
   Dim adoconn As New ADODB.Connection

   Me.Height = 7005
   Me.Width = 8040
   Me.BackColor = MODULEBACKCOLOR
   fraEdit.BackColor = MODULEBACKCOLOR

   ConfigflxSupplierList
   ConfigFlxPaymentSplit

'   Set the ADO Connections to the dataset
   adoconn.Open getConnectionString
   'Below line added by anol 25 aug 2015
   'issue 571 note 1148
'    PrepareCboClient adoConn
'   Load All supplier in the dropdown combo
   ConfigflxSupplierList
   loadFund adoconn
  ' LoadAllSupplierFlxGrd adoconn
   LoadBankAccountInCombo adoconn
   LoadPayAmtType "RECEIPT AMOUNT TYPE", adoconn
   'LoadFlxPaymentSplit adoConn

   adoconn.Close
   Set adoconn = Nothing

   flxPaymentSplit.row = 1
   flxPaymentSplit_RowColChange
'   If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
        Call WheelHook(Me.hWnd)
'   End If
End Sub

Public Sub LoadFlxPaymentSplit(adoconn As ADODB.Connection)
        Dim szSQL As String
        If BoolReconciled Then 'if transaction is reconciled only some part of this line can be edited issue 515
           txtAmount.Locked = True
           cmdClient.Enabled = False
           cmdSupplier.Enabled = False
           txtDate.Locked = True
           cmbFund.Locked = False
           cboBC.Locked = True
           txtProperty.Locked = True
        Else
           txtAmount.Locked = False
           cmdClient.Enabled = True
           cmdSupplier.Enabled = True
           txtDate.Locked = False
           cmbFund.Locked = False
           cboBC.Locked = False
           txtProperty.Locked = False
        End If
        If frmPurchaseExpense.tabPurExp.Tab = 1 Then
             If frmPurchaseExpense.sEditPPR = 1 Then
                    szSQL = "SELECT P.TransactionID, P.FundID, '' AS No, F.FundName, P.Details AS Description, " & _
                           "P.PDate AS DueDate, FORMAT(P.Amount,'0.00') AS AMOUNT, P.ReconNow, P.POSTINGDATE,P.UnitID as PropertyID " & _
                       "FROM tlbPayment AS P INNER JOIN Fund AS F ON P.FundID = F.FundID " & _
                       "WHERE P.TransactionID = " & TransactionID & " AND " & _
                           "P.Type = 24 " & _
                       "ORDER BY P.TransactionID;"
            'Debug.Print szSQL
               iRecords = populateGridDefinedHeader(adoconn, szSQL, flxPaymentSplit)
               If iRecords > 0 Then
                    txtAmount.text = Format(Val(flxPaymentSplit.TextMatrix(1, 6)), "0.00")
                    txtDetails.text = flxPaymentSplit.TextMatrix(1, 4)
                    lblPostingDate.ToolTipText = flxPaymentSplit.TextMatrix(1, 9)
                    txtProperty.text = flxPaymentSplit.TextMatrix(1, 10)
               End If
        
          End If
          If frmPurchaseExpense.sEditPPR = 2 Then
                If frmPurchaseExpense.flxSCrPoA.row = 0 Then
                      MsgBox "Please select a payment line to edit"
                      Unload Me
                End If
               
                 szSQL = "SELECT S.TransactionID, S.FundID, S.SplitID AS No, F.FundName, S.Description, " & _
                             "S.DueDate, FORMAT(S.Amount,'0.00') AS AMOUNT, '', P.ReconNow, P.POSTINGDATE,P.UnitID as PropertyID " & _
                         "FROM (tlbPaymentSplit AS S INNER JOIN Fund AS F ON S.FundID = F.FundID) INNER JOIN " & _
                             "tlbPayment AS P ON S.PayHeader = P.TransactionID " & _
                         "WHERE P.TransactionID = " & TransactionID & " AND " & _
                             "(P.Type = 8 OR P.Type = 9) " & _
                         "ORDER BY S.SplitID;"
              'End If
           'Debug.Print szSQL
              iRecords = populateGridDefinedHeader(adoconn, szSQL, flxPaymentSplit)
               If iRecords > 0 Then
                    txtAmount.text = Format(Val(flxPaymentSplit.TextMatrix(1, 6)), "0.00")
                    txtDetails.text = flxPaymentSplit.TextMatrix(1, 4)
                    lblPostingDate.ToolTipText = flxPaymentSplit.TextMatrix(1, 9)
                    txtProperty.text = flxPaymentSplit.TextMatrix(1, 10)
               End If
        End If
    End If
    If frmPurchaseExpense.tabPurExp.Tab = 3 Then
           'This shall include 3 types (24,8,9)
           szSQL = "SELECT S.TransactionID, S.FundID, S.SplitID AS No, F.FundName, S.Description, " & _
                          "S.DueDate, FORMAT(S.Amount,'0.00') AS AMOUNT, '', P.ReconNow, P.POSTINGDATE,P.UnitID as PropertyID " & _
                      "FROM (tlbPaymentSplit AS S INNER JOIN Fund AS F ON S.FundID = F.FundID) INNER JOIN " & _
                          "tlbPayment AS P ON S.PayHeader = P.TransactionID " & _
                      "WHERE P.TransactionID = " & TransactionID & " AND " & _
                          "(P.Type = 8 OR P.Type = 9 OR P.Type = 24) " & _
                      "ORDER BY S.SplitID;"
           iRecords = populateGridDefinedHeader(adoconn, szSQL, flxPaymentSplit)
            If iRecords > 0 Then
                    txtAmount.text = Format(Val(flxPaymentSplit.TextMatrix(1, 6)), "0.00")
                    txtDetails.text = flxPaymentSplit.TextMatrix(1, 4)
                    lblPostingDate.ToolTipText = flxPaymentSplit.TextMatrix(1, 9)
               End If
    End If
End Sub

Private Sub LoadPayAmtType(szValue As String, adoconn As ADODB.Connection)
   Dim SQLStr1 As String, szaData() As String, i As Integer
   Dim adoRst As New ADODB.Recordset

   SQLStr1 = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
             "FROM PrimaryCode, SecondaryCode " & _
             "WHERE PrimaryCode.Value = '" & szValue & "' AND " & _
                  "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
             "ORDER BY SecondaryCode.Value;"

   adoRst.Open SQLStr1, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRst.RecordCount - 1) As String

   cmbSPAmtType.Clear
   i = 0
   While Not adoRst.EOF
      szaData(0, i) = adoRst!c
      szaData(1, i) = adoRst!V
      adoRst.MoveNext
      i = i + 1
   Wend
   adoRst.Close
   Set adoRst = Nothing

   cmbSPAmtType.Column() = szaData()
End Sub

Private Sub LoadBankAccountInCombo(ByVal adoconn As ADODB.Connection)
   On Error GoTo Error_Handler

   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, Data() As String, j As Integer
   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = '" & frmPurchaseExpense.txtClientIDPurPay.text & "' AND tlbClientBanks.CLIENT_ID=NominalLedger.ClientID order by NominalCode;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   iTotalRow = adoRst.RecordCount
   iTotalCol = adoRst.Fields.Count
   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String

   For i = 0 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboBC.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

Error_Handler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   Set adoRst = Nothing
End Sub
Private Sub LoadBankAccountInComboClient(ByVal adoconn As ADODB.Connection)
   On Error GoTo Error_Handler

   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, Data() As String, j As Integer
   Dim i As Integer, iTotalCol As Integer, iTotalRow As Integer
   If txtClient.text = "" Then Exit Sub
   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = '" & txtClient.text & "' AND tlbClientBanks.CLIENT_ID=NominalLedger.ClientID order by NominalCode;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   iTotalRow = adoRst.RecordCount
   iTotalCol = adoRst.Fields.Count
   ReDim Data(iTotalCol - 1, iTotalRow - 1) As String

   For i = 0 To iTotalRow
       For j = 0 To iTotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboBC.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing
   Exit Sub

Error_Handler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   Set adoRst = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
    frmPurchaseExpense.Enabled = True
    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    If IsNull(frmPurchaseExpense.txtClientIDPurPay.text) = True Then
       ShowMsgInTaskBar "Please select a client from purchase and expense payment", "Y"
       'cboClient.SetFocus
       Exit Sub
    End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, frmPurchaseExpense.txtClientIDPurPay.text
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 45 Then
        KeyAscii = 0
   End If
   If txtAmount.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and the amount cannot be edited", "Y", "N"
      Exit Sub
   End If
   If KeyAscii = 13 Then
        FocusControl cboBC
   End If
   DigitTextKeyPress txtAmount, KeyAscii

   If KeyAscii = 27 And txtAmount.text <> "" Then
      txtAmount.text = ""
   End If
   If KeyAscii = 27 And txtAmount.text = "" Then Unload Me
   
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
   lblPostingDate.ToolTipText = txtDate.text
End Sub

Private Sub txtDate_GotFocus()
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If txtAmount.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and the date cannot be edited", "Y", "N"
      Exit Sub
   End If
   If KeyAscii = 13 Then
        FocusControl cmbFund
   End If
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   TextBoxFormatDate txtDate
        If IsDate(lblPostingDate.ToolTipText) = True Then
              Dim adoconn As New ADODB.Connection
              Dim szSQL As String
              If IsNull(frmPurchaseExpense.txtClientIDPurPay.text) Then
                  ShowMsgInTaskBar "Please select a client on Purchase Expense Payment", "Y", "N"
                  'frmPurchaseExpense.cboClient.Value.SetFocus
                  Exit Sub
              End If
              adoconn.Open getConnectionString
              If IsPeriodStatus(lblPostingDate.ToolTipText, frmPurchaseExpense.txtClientIDPurPay.text, adoconn) = 0 Then
                  ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
                  adoconn.Close
                  Exit Sub
              ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, frmPurchaseExpense.txtClientIDPurPay.text, adoconn) = 9 Then
                  ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
                  adoconn.Close
                  Exit Sub
              End If
           End If
End Sub
'Private Function SupplierType() As String
'    On Error GoTo Err
'    'mark 4
'    Dim szType As String
'    If UCase(frmPurchaseExpense.cmdACType.Column(0)) = UCase("Managing agent") Then
'         SupplierType = UCase("AGENT")
'    ElseIf UCase(frmPurchaseExpense.cmdACType.Column(0)) = UCase("Client") Then
'        SupplierType = UCase("Client")
'    ElseIf UCase(frmPurchaseExpense.cmdACType.Column(0)) = UCase("Landlord") Then
'        SupplierType = UCase("Landlord")
'    Else
'        SupplierType = UCase("ALL")
'    End If
'    Exit Function
'Err:
'   SupplierType = UCase("ALL")
'End Function
'Private Sub LoadAllSupplierFlxGrd(ByVal adoconn As ADODB.Connection)
'   Dim adoRst  As New ADODB.Recordset
'   Dim adoLL   As New ADODB.Recordset
'   Dim adoC    As New ADODB.Recordset
'   Dim adoMA   As New ADODB.Recordset
'
'   Dim szSQL      As String
'   Dim iTotalRow  As Integer
'   Dim j          As Integer
'   Dim i          As Integer
'   Dim iTotalCol  As Integer
'   Dim Data()     As String
'
'    'mark 3
'   On Error GoTo ErrorHandler
''Modified by anol 28 Oct 2015
'
'   If SupplierType = "ALL" Then
'        szSQL = "SELECT SupplierID, SupplierName  " & _
'           "FROM Supplier ORDER BY SupplierName;"
'   Else
'        szSQL = "SELECT SupplierID, SupplierName  " & _
'           "FROM Supplier where TYPE='" & SupplierType & "'" & _
'           "ORDER BY SupplierName;"
'   End If
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT DISTINCT L.LandlordID, L.LandlordName " & _
'           "FROM   PropertyLandlord AS PL, Landlord AS L " & _
'           "WHERE  PL.LandlordID = L.LandlordID AND 1=2;"
''Debug.Print szSQL
'   adoLL.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT   DISTINCT C.ClientID, C.ClientName " & _
'           "FROM     Client AS C where 1=2 " & _
'           "ORDER BY C.ClientName;"
''Debug.Print szSQL
'   adoC.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT   DISTINCT A.AgentID, A.AgentName " & _
'           "FROM     Agent AS A where 1=2 " & _
'           "ORDER BY A.AgentName;"
''Debug.Print szSQL
'   adoMA.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'
'   iTotalRow = adoRst.RecordCount + adoLL.RecordCount + adoC.RecordCount + adoMA.RecordCount
'   If iTotalRow = 0 Then
'        adoRst.Close
'        adoLL.Close
'        adoC.Close
'        adoMA.Close
'        GoTo NoRes
'   End If
'
'   iTotalCol = adoRst.Fields.count
'
'   ReDim Data(iTotalCol, iTotalRow - 1) As String
'
'   If adoRst.RecordCount > 0 Then
'      For i = 0 To iTotalRow
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoRst.Fields.Item(j).Value), "", adoRst.Fields.Item(j).Value)
'          Next j
''          Data(j, i) = Format(GetSupplierBalance(adoRst.Fields.Item(0).Value), "0.00")
'          adoRst.MoveNext
'          If adoRst.EOF Then Exit For
'      Next i
'   End If
'
'   If adoLL.RecordCount > 0 Then
'      For i = i + 1 To iTotalRow - 1
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoLL.Fields.Item(j).Value), "", adoLL.Fields.Item(j).Value)
'          Next j
'          adoLL.MoveNext
'          If adoLL.EOF Then Exit For
'      Next i
'   End If
'
'   If adoC.RecordCount > 0 Then
'      For i = i + 1 To iTotalRow - 1
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoC.Fields.Item(j).Value), "", adoC.Fields.Item(j).Value)
'          Next j
'          adoC.MoveNext
'          If adoC.EOF Then Exit For
'      Next i
'   End If
'
'   If adoMA.RecordCount > 0 Then
'      For i = i + 1 To iTotalRow - 1
'          For j = 0 To iTotalCol - 1
'              Data(j, i) = IIf(IsNull(adoMA.Fields.Item(j).Value), "", adoMA.Fields.Item(j).Value)
'          Next j
'          adoMA.MoveNext
'          If adoMA.EOF Then Exit For
'      Next i
'   End If
'
'   cmbSPSupplier.Column() = Data()
'
''  LoadFlxSupplier ---->>
'   adoRst.MoveFirst
'   If adoLL.RecordCount > 0 Then adoLL.MoveFirst
'   If adoC.RecordCount > 0 Then adoC.MoveFirst
'   If adoMA.RecordCount > 0 Then adoMA.MoveFirst
'   i = 1
'
'   While Not adoRst.EOF
'      flxSupplierList.TextMatrix(i, 0) = "" 'frmPurchaseExpense.cmdACType.Column(0)
'      flxSupplierList.TextMatrix(i, 1) = adoRst.Fields.Item(0).Value
'      flxSupplierList.TextMatrix(i, 2) = adoRst.Fields.Item(1).Value
''      flxSupplierList.TextMatrix(i, 3) = Format(GetSupplierBalance(adoRst.Fields.Item(0).Value), "0.00")
'      adoRst.MoveNext
'      If Not adoRst.EOF Then flxSupplierList.AddItem ""
'      i = i + 1
'   Wend
'   adoRst.Close
'
'   While Not adoLL.EOF
'      If Not adoLL.EOF Then flxSupplierList.AddItem ""
'      flxSupplierList.TextMatrix(i, 0) = "Landlord"
'      flxSupplierList.TextMatrix(i, 1) = adoLL.Fields.Item(0).Value
'      flxSupplierList.TextMatrix(i, 2) = adoLL.Fields.Item(1).Value
''      flxSupplierList.TextMatrix(i, 3) = Format(GetLandlordBalance(adoLL.Fields.Item(0).Value), "0.00")
'      adoLL.MoveNext
'      i = i + 1
'   Wend
'   adoLL.Close
'
'   While Not adoC.EOF
'      If Not adoC.EOF Then flxSupplierList.AddItem ""
'      flxSupplierList.TextMatrix(i, 0) = "Client"
'      flxSupplierList.TextMatrix(i, 1) = adoC.Fields.Item(0).Value
'      flxSupplierList.TextMatrix(i, 2) = adoC.Fields.Item(1).Value
''      flxSupplierList.TextMatrix(i, 3) = Format(GetClientBalance(adoC.Fields.Item(0).Value), "0.00")
'      adoC.MoveNext
'      i = i + 1
'   Wend
'   adoC.Close
'
'   While Not adoMA.EOF
'      If Not adoMA.EOF Then flxSupplierList.AddItem ""
'      flxSupplierList.TextMatrix(i, 0) = "Managing Agent"
'      flxSupplierList.TextMatrix(i, 1) = adoMA.Fields.Item(0).Value
'      flxSupplierList.TextMatrix(i, 2) = adoMA.Fields.Item(1).Value
''      flxSupplierList.TextMatrix(i, 3) = Format(GetAgentBalance(adoMA.Fields.Item(0).Value), "0.00")
'      adoMA.MoveNext
'      i = i + 1
'   Wend
'   adoMA.Close
'
'NoRes:
'   szSQL = "SELECT FundID, FundName FROM Fund;"
'   If adoRst.State = 1 Then
'        adoRst.Close
'   End If
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   ReDim szaData(1, adoRst.RecordCount) As String
'
'   i = 1
'   szaData(0, 0) = "0"
'   szaData(1, 0) = "Not Found"
'   While Not adoRst.EOF
'      szaData(0, i) = adoRst.Fields.Item("FundID").Value
'      szaData(1, i) = adoRst.Fields.Item("FundName").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cmbFund.Clear
'   cmbFund.Column() = szaData()
'   adoRst.Close
'
'   Set adoRst = Nothing
'   Set adoLL = Nothing
'   Set adoC = Nothing
'   Set adoMA = Nothing
'   Exit Sub
'
'ErrorHandler:
'   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
'
'   Set adoRst = Nothing
'   Set adoLL = Nothing
'   Set adoC = Nothing
'   Set adoMA = Nothing
'End Sub
Private Sub loadFund(adoconn As ADODB.Connection)
    Dim szSQL As String
    Dim adoRst As New Recordset
    Dim i As Integer
    szSQL = "SELECT FundID, FundName FROM Fund;"
   
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    ReDim szaData(1, adoRst.RecordCount) As String

    i = 0
'    szaData(0, 0) = "0"
'    szaData(1, 0) = "Not Found"
    While Not adoRst.EOF
        szaData(0, i) = adoRst.Fields.Item("FundID").Value
        szaData(1, i) = adoRst.Fields.Item("FundName").Value
        i = i + 1
        adoRst.MoveNext
    Wend
    
    cmbFund.Clear
    cmbFund.Column() = szaData()
    adoRst.Close
End Sub

Private Sub ConfigflxSupplierList()
   Dim szHeader As String

   flxSupplierList.Clear
   flxSupplierList.Rows = 2
   flxSupplierList.Cols = 4
'   flxSupplierList.Width = 5805
   szHeader$ = "<|<|<"
   flxSupplierList.FormatString = szHeader$

   flxSupplierList.RowHeight(0) = 0
   flxSupplierList.ColWidth(0) = 120 'Label20(1).Left - Label20(0).Left
   flxSupplierList.ColWidth(1) = txtSupplierSearchID.Width + 40 'Label20(2).Left - Label20(1).Left
   flxSupplierList.ColWidth(2) = txtSupplierSearchName.Width
   flxSupplierList.ColWidth(3) = 0
   flxSupplierList.Width = 160 + txtSupplierSearchID.Width + txtSupplierSearchName.Width
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
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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


Private Sub txtSupplierSearchID_Change()
    Dim i As Integer
'   txtSupplierSearchID.text = UCase(txtSupplierSearchID.text)
   If Len(txtSupplierSearchID.text) > 0 Then
        txtSupplierSearchName.text = ""
   End If

   For i = flxSupplierList.Rows - 1 To 1 Step -1
      flxSupplierList.RowHeight(i) = 240
      
      If InStr(1, UCase(flxSupplierList.TextMatrix(i, 1)), UCase(txtSupplierSearchID.text), vbTextCompare) = 0 Then
            flxSupplierList.RowHeight(i) = 0
      End If
      If flxSupplierList.RowHeight(i) = 240 Then
            flxSupplierList.row = i
      End If
   Next i
End Sub

Private Sub txtSupplierSearchID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        flxSupplierList.SetFocus
    End If
End Sub

Private Sub txtSupplierSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtSupplierSearchID.text) > 0 Then
            flxSupplierList.SetFocus
        Else
            txtSupplierSearchName.SetFocus
        End If
    End If
End Sub

Private Sub txtSupplierSearchName_Change()
     Dim i As Integer
   txtSupplierSearchName.text = UCase(txtSupplierSearchName.text)
   If Len(txtSupplierSearchName.text) > 0 Then
        txtSupplierSearchID.text = ""
   End If

   For i = flxSupplierList.Rows - 1 To 1 Step -1
      flxSupplierList.RowHeight(i) = 240
      
      If InStr(1, UCase(flxSupplierList.TextMatrix(i, 2)), UCase(txtSupplierSearchName.text), vbTextCompare) = 0 Then
            flxSupplierList.RowHeight(i) = 0
      End If
      If flxSupplierList.RowHeight(i) = 240 Then
            flxSupplierList.row = i
      End If
   Next i
End Sub

Private Sub txtSupplierSearchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxSupplierList.SetFocus
    End If
End Sub

