VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDemandTypesSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generating Demands - Options"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemandTypesSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   9615
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   5085
      ScaleHeight     =   450
      ScaleWidth      =   2655
      TabIndex        =   31
      Top             =   5175
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading ...."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   270
         TabIndex        =   32
         Top             =   135
         Width           =   2745
      End
   End
   Begin VB.TextBox txtClientSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   2
      Top             =   405
      Width           =   1575
   End
   Begin VB.TextBox txtPostingDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7230
      MaxLength       =   10
      TabIndex        =   1
      Top             =   80
      Width           =   1575
   End
   Begin VB.Frame fraSC_St 
      Height          =   615
      Left            =   60
      TabIndex        =   26
      Top             =   8730
      Width           =   4215
      Begin VB.TextBox txtSCDateTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         Top             =   255
         Width           =   1575
      End
      Begin VB.TextBox txtSCDateFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   260
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   255
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   260
         Width           =   495
      End
      Begin MSForms.CheckBox chkSCS 
         Height          =   375
         Left            =   75
         TabIndex        =   29
         Top             =   -45
         Width           =   2415
         BackColor       =   -2147483633
         ForeColor       =   16711680
         DisplayStyle    =   4
         Size            =   "4260;661"
         Value           =   "0"
         Caption         =   "Service Charge Statement:"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
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
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   3855
      TabIndex        =   24
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
         TabIndex        =   25
         Top             =   15
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdGDPOk 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5895
      TabIndex        =   9
      Top             =   9315
      Width           =   1440
   End
   Begin VB.Frame fraTransType 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4395
      TabIndex        =   22
      Top             =   8550
      Width           =   3975
      Begin VB.OptionButton optAutoGenConsolidated 
         Caption         =   "Consolidated Transactions"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1725
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
      Begin VB.OptionButton optAutoGenSig 
         Caption         =   "Single Transactions"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.TextBox txtInitialIssueDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   0
      Top             =   80
      Width           =   1575
   End
   Begin VB.CheckBox chkDT 
      Appearance      =   0  'Flat
      Caption         =   "Select All Demand Types"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3810
      Width           =   2055
   End
   Begin VB.CheckBox chkProp 
      Appearance      =   0  'Flat
      Caption         =   "Select All Properties"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6630
      TabIndex        =   4
      Top             =   450
      Width           =   1815
   End
   Begin VB.CommandButton cmdGDPCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7695
      TabIndex        =   10
      Top             =   9315
      Width           =   1440
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
      Height          =   3000
      Left            =   60
      TabIndex        =   3
      Top             =   750
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   5292
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   3000
      Left            =   4365
      TabIndex        =   5
      Top             =   750
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5292
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypes 
      Height          =   4185
      Left            =   3330
      TabIndex        =   7
      Top             =   4140
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   4
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
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCategory 
      Height          =   4185
      Left            =   60
      TabIndex        =   11
      Top             =   4140
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   7382
      _Version        =   393216
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
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6225
      TabIndex        =   30
      Top             =   75
      Width           =   945
   End
   Begin MSForms.CheckBox chkSentAutoEmail 
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   8805
      Width           =   3255
      VariousPropertyBits=   746588179
      BackColor       =   12648447
      ForeColor       =   16711680
      DisplayStyle    =   4
      Size            =   "5741;873"
      Value           =   "0"
      Caption         =   "Automatically Send Demands by Email"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Category:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   23
      Top             =   3900
      Width           =   1290
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Freq:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Index           =   32
      Left            =   150
      TabIndex        =   21
      Top             =   8415
      Width           =   360
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   60
      TabIndex        =   20
      Top             =   75
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Demand Types:"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3420
      TabIndex        =   19
      Top             =   3870
      Width           =   1065
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client (Search)"
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
      Left            =   60
      TabIndex        =   18
      Top             =   405
      Width           =   1035
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   450
      Width           =   645
   End
   Begin MSForms.ComboBox cboGDPFreq 
      Height          =   270
      Left            =   570
      TabIndex        =   8
      Top             =   8385
      Width           =   3135
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5530;476"
      TextColumn      =   2
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
End
Attribute VB_Name = "frmDemandTypesSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iSelDemandCategory  As Integer
Public iSelDemandTypes     As Integer
Public iSelProperties      As Integer     'Number of properties selected by the user
Public szCallingFrom       As String

Private Type SendDemandByEmail
   szLesseeID    As String
   szLesseeEmail As String
   szClient      As String
   colAtt        As Collection
   colURN        As Collection
End Type

Private uLessee()   As SendDemandByEmail
Private iLes        As Integer, szEmailDmdIdList   As String
Private szaProp()   As String, iProperty           As Integer
Private szaDT()     As String, iDT                 As Integer
Private iSelClient  As Integer

Private Sub chkDT_Click()
   If iSelProperties = 0 Then
      chkDT.Value = 0
      ShowMsgInTaskBar "Please select a property", "Y", "N"
      Exit Sub
   End If

   Dim i As Integer

   iSelDemandTypes = 0

   If chkDT.Value = 1 Then
      For i = 1 To flxDemandTypes.Rows - 1
         If flxDemandTypes.TextMatrix(i, 0) <> "X" And flxDemandTypes.TextMatrix(i, 1) <> "" Then
            SelectFlxGridRow 0, flxDemandTypes, i
            iSelDemandTypes = iSelDemandTypes + 1
         End If
      Next i
   Else
      For i = 1 To flxDemandTypes.Rows - 1
         If flxDemandTypes.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxDemandTypes, i
         End If
      Next i
   End If
End Sub

Private Sub chkDT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDemandTypes.SetFocus
    End If
End Sub

Private Sub chkProp_Click()
   If iSelClient = 0 Then
      chkProp.Value = 0
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      Exit Sub
   End If

   Dim i As Integer

   If chkProp.Value = 1 Then
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) <> "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
      iSelProperties = flxProperties.Rows - 1
   Else
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
      iSelProperties = 0
   End If
   FilterDemandTypes
End Sub

Private Sub chkSCS_Click()
   If chkSCS.Value Then
      txtSCDateFrom.Locked = False
      txtSCDateFrom.SetFocus
      txtSCDateTo.Locked = False
   Else
      txtSCDateFrom.Locked = True
      txtSCDateTo.Locked = True
      txtSCDateFrom.text = ""
      txtSCDateTo.text = ""
   End If
End Sub

Private Sub cmdGDPCancel_Click()
   frmDemands3.Enabled = True
   Unload Me
End Sub

Private Sub flxCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxDemandTypes.MousePointer = vbArrow
End Sub

Private Sub flxCategory_RowColChange()
   SelectFlxGridRow 0, flxCategory, flxCategory.row
   If flxCategory.TextMatrix(flxCategory.row, 0) = "X" Then
      iSelDemandCategory = iSelDemandCategory + 1
   Else
      iSelDemandCategory = iSelDemandCategory - 1
   End If

   FilterDemandTypes
End Sub

Private Sub flxClients_Click()
   SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
   iSelClient = 1

   FilterProperties
   flxDemandTypes.Clear
End Sub

Private Sub flxClients_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
           SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
           iSelClient = 1
           FilterProperties
           flxDemandTypes.Clear
    End If
End Sub

Private Sub flxClients_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClients_Click
        flxProperties.SetFocus
    End If
End Sub

Private Sub flxDemandTypes_Click()
   'from colrow change to click by anol 16 aug 2016
   SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
   If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
      iSelDemandTypes = iSelDemandTypes + 1
   Else
      iSelDemandTypes = iSelDemandTypes - 1
   End If
End Sub

Private Sub flxDemandTypes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
         SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
        If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
           iSelDemandTypes = iSelDemandTypes + 1
        Else
           iSelDemandTypes = iSelDemandTypes - 1
        End If
    End If
End Sub

Private Sub flxDemandTypes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        flxDemandTypes_Click
        optAutoGenSig.SetFocus
    End If
End Sub

Private Sub flxProperties_Click()
    'from colrow change to click by anol 16 Aug 2016
   SelectFlxGridRow 0, flxProperties, flxProperties.row
'   If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'      iSelProperties = iSelProperties + 1
'   Else
'      iSelProperties = iSelProperties - 1
'   End If
            'modified by anol 20160920
            Dim i As Integer
            iSelProperties = 0
            For i = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(i, 0) = "X" Then
                    iSelProperties = iSelProperties + 1
                End If
            Next i

   FilterDemandTypes
End Sub

Private Sub flxProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
           SelectFlxGridRow 0, flxProperties, flxProperties.row
'           If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'              iSelProperties = iSelProperties + 1
'           Else
'              iSelProperties = iSelProperties - 1
'           End If
            'modified by anol 20160920
            Dim i As Integer
            iSelProperties = 0
            For i = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(i, 0) = "X" Then
                    iSelProperties = iSelProperties + 1
                End If
            Next i
           FilterDemandTypes
    End If
End Sub

Private Sub flxProperties_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        flxProperties_Click
        chkDT.SetFocus
    End If
End Sub

Private Sub Form_Activate()
'   If szCallingFrom = "Automatic Demand" Then
'      fraTransType.Visible = True
'   Else
'      fraTransType.Visible = False
'   End If
   If szCallingFrom = "Print All Demands" Then
      fraTransType.Visible = False
   End If
End Sub

Private Sub optAutoGenSig_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        cmdGDPOk.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        optAutoGenConsolidated.SetFocus
    End If
    
    If KeyCode = 13 Then
        cmdGDPOk.SetFocus
    End If
End Sub

Private Sub txtClientSearch_Change()
            'Updated by anol 22 Dec 2015
   Dim i As Integer
   For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.RowHeight(i) = 240
            If InStr(1, UCase(flxClients.TextMatrix(i, 1)), UCase(txtClientSearch.text), vbTextCompare) = 0 And txtClientSearch.text <> "" Then
                flxClients.RowHeight(i) = 0
            End If
       
      If flxClients.RowHeight(i) = 240 Then
            flxClients.row = i
      End If
   Next i
End Sub

Private Sub txtClientSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        flxClients.SetFocus
    End If
End Sub

Private Sub txtClientSearch_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        flxClients.SetFocus
'    End If
End Sub

Private Sub txtInitialIssueDate_Change()
   TextBoxChangeDate txtInitialIssueDate
End Sub

Private Sub txtInitialIssueDate_GotFocus()
   SelTxtInCtrl txtInitialIssueDate
End Sub

Private Sub txtInitialIssueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientSearch.SetFocus
    End If
   TextBoxKeyPrsDate txtInitialIssueDate, KeyAscii
End Sub

Private Sub txtInitialIssueDate_LostFocus()
    TextBoxFormatDate txtInitialIssueDate
   If IsDate(txtInitialIssueDate.text) = True Then
        txtPostingDate.text = txtInitialIssueDate.text
      'Modified by BOSL
      'issue 468
      'Modified by anol 01 Sep 2014
      If flxClients.row = 0 Then
        ShowMsgInTaskBar "Please select a client", "Y", "N"
'        flxClients.SetFocus
        Exit Sub
     End If
      If frmMMain.IsRibbonVersion And IsDate(txtPostingDate.text) = True Then
        Dim adoConn1 As New ADODB.Connection
        Dim szSQL As String
        If flxClients.TextMatrix(flxClients.row, 1) = "" Then Exit Sub
        adoConn1.Open getConnectionString
        If IsPeriodStatus(txtPostingDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 0 Then
            MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
            adoConn1.Close
            FocusControl txtInitialIssueDate
            Exit Sub
        ElseIf IsPeriodStatus(txtPostingDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 9 Then
            MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
            adoConn1.Close
            FocusControl txtInitialIssueDate
            Exit Sub
        End If
     End If
    'End of modification
   End If
End Sub

Private Sub ClearLocks(adoConn As ADODB.Connection)
   adoConn.Execute "UPDATE Property SET RunningAutoDemand = '' WHERE RunningAutoDemand = '" & WS_Name & "';"
End Sub

Private Function IsDemandRunning(szProperties As String, adoConn As ADODB.Connection) As Boolean
   Dim iProp   As Integer
   Dim szSQL   As String
   Dim adoRst  As New ADODB.Recordset

   On Error GoTo ErrLocking

   szSQL = ""
   IsDemandRunning = False
   adoConn.BeginTrans

   adoConn.Execute "UPDATE Property SET RunningAutoDemand = '' WHERE RunningAutoDemand = '" & WS_Name & "';"

   szSQL = "SELECT * FROM Property WHERE ClientID = '" & _
                     flxClients.TextMatrix(flxClients.row, 1) & "';"
   adoRst.Open szSQL, adoConn, adOpenForwardOnly, adLockOptimistic

   While Not adoRst.EOF
      For iProp = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(iProp, 0) = "X" Then
            If adoRst.Fields.Item("PropertyID").Value = flxProperties.TextMatrix(iProp, 1) Then
               If adoRst.Fields.Item("RunningAutoDemand").Value <> "" And _
                     adoRst.Fields.Item("RunningAutoDemand").Value <> WS_Name Then
'                  szProperties = szProperties + ", " + adoRst.Fields.Item("PropertyName").Value
                  szProperties = adoRst.Fields.Item("PropertyName").Value + Chr(13) + szProperties
                  IsDemandRunning = True
               Else
                  adoRst.Fields.Item("RunningAutoDemand").Value = WS_Name
                  adoRst.Update
               End If
            End If
         End If
      Next iProp
      adoRst.MoveNext
   Wend

   adoConn.CommitTrans

   adoRst.Close
   Set adoRst = Nothing
   
   Exit Function

ErrLocking:
   IsDemandRunning = True
   adoConn.RollbackTrans
   Set adoRst = Nothing
End Function
Private Function Check_Fund_LRentCharge(adoConn1 As ADODB.Connection) As Boolean
'2020-05-28 added by anol
    Dim rsLRentCharges As New ADODB.Recordset
    Dim rsLease As New ADODB.Recordset
    Dim szSQL As String
    Dim szSAGEID As String
    rsLRentCharges.Open "Select * from LRentCharges where (RentChargeDept='' or isnull(RentChargeDept)) AND isnull(Spare3)", adoConn1, adOpenStatic, adLockReadOnly
    If Not rsLRentCharges.EOF Then
        szSQL = "Select * from LeaseDetails L where L.LeaseID='" & rsLRentCharges("LeaseID").Value & "'"
        rsLease.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
        If Not rsLease.EOF Then
            szSAGEID = rsLease("SageAccountNumber").Value
        End If
        rsLease.Close
        MsgBox "There is a 'Fund Name'  missing against Lease ID : '" & szSAGEID & "' . Please correct this in Lease Details before generating your demands", vbCritical, "Fund Name missing  "
        Check_Fund_LRentCharge = True
    End If
    rsLRentCharges.Close
    
End Function

Private Sub cmdGDPOk_Click()
   Dim i As Integer
   Dim rsDemandType As New ADODB.Recordset
   If txtInitialIssueDate.text = "" Then
      ShowMsgInTaskBar "Please type the Issue date.", , "N"
      txtInitialIssueDate.SetFocus
      Exit Sub
   End If
   'Modified by BOSL
   'Issue 468
   'Modified by Anol 02 Sep 2014
   If IsDate(txtPostingDate.text) = False Then
        ShowMsgInTaskBar "Please enter posting date.", , "N"
        txtPostingDate.SetFocus
        Exit Sub
   End If
   If flxClients.row = 0 Then
        ShowMsgInTaskBar "Please select a client", "Y", "N"
        flxClients.SetFocus
        Exit Sub
     End If
    If frmMMain.IsRibbonVersion And IsDate(txtPostingDate.text) = True Then
        Dim adoConn1 As New ADODB.Connection
        Dim szSQL As String
        adoConn1.Open getConnectionString
        If Check_Fund_LRentCharge(adoConn1) = True Then
            adoConn1.Close
            Exit Sub
        End If
'        'added by anol 28 FEB 2016
'        'Adding FreqID column to LeaseDetails table
'        Call addFreqID(adoConn1)
'        'Updating frequency ID
'        adoConn1.Execute "Update LeaseDetails,LRentCharges,Frequencies set LeaseDetails.FreqID =LRentCharges.BRFrequency where Frequencies.ID=LRentCharges.BRFrequency AND LRentCharges.LeaseID=LeaseDetails.LeaseID"
'        adoConn1.Execute "Update LeaseDetails,LInsuranceCharges,Frequencies set LeaseDetails.FreqID =LInsuranceCharges.InsuranceFrequency where  Frequencies.ID=LInsuranceCharges.InsuranceFrequency AND LInsuranceCharges.LeaseID=LeaseDetails.LeaseID"
'        adoConn1.Execute "Update LeaseDetails,LServiceCharges,Frequencies set LeaseDetails.FreqID =LServiceCharges.SCFrequency where  Frequencies.ID=LServiceCharges.SCFrequency AND LServiceCharges.LeaseID=LeaseDetails.LeaseID"
'        'end of addition
        If IsPeriodStatus(txtInitialIssueDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 0 Then
            MsgBox "The issue date cannot fall within a closed financial period", vbInformation, "Warning"
            adoConn1.Close
            Exit Sub
        ElseIf IsPeriodStatus(txtInitialIssueDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 9 Then
            MsgBox "The issue date does not fall in any existing financial period", vbInformation, "Warning"
            adoConn1.Close
            FocusControl txtInitialIssueDate
            Exit Sub
        ElseIf IsPeriodStatus(txtPostingDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 0 Then
            MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
            adoConn1.Close
            FocusControl txtInitialIssueDate
            Exit Sub
        ElseIf IsPeriodStatus(txtPostingDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn1) = 9 Then
            MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
            adoConn1.Close
            FocusControl txtInitialIssueDate
            Exit Sub
        End If
     End If
   'End of modification
   If iSelDemandTypes = 0 Then
      MsgBox "Please select demand types.", vbInformation, "Warning"
      Exit Sub
   End If

   If fraSC_St.Visible And chkSCS.Value Then
      If txtSCDateFrom.text = "" Then
         MsgBox "Please enter the Start date for Service Charge Statement.", vbInformation, "Warning"
         txtSCDateFrom.SetFocus
         Exit Sub
      End If
      If txtSCDateTo.text = "" Then
         MsgBox "Please enter the End date for Service Charge Statement.", vbInformation, "Warning"
         txtSCDateTo.SetFocus
         Exit Sub
      End If
   End If
   If adoConn1.State = 0 Then
        adoConn1.Open getConnectionString
   End If
   For i = 1 To flxDemandTypes.Rows - 1
        If flxDemandTypes.TextMatrix(i, 0) = "X" And optAutoGenSig.Value = True Then
                rsDemandType.Open "Select * from DemandTypes where ID=" & flxDemandTypes.TextMatrix(i, 2) & " and Consolidated=true ", adoConn1, adOpenKeyset, adLockReadOnly
                If Not rsDemandType.EOF Then
                    MsgBox "The demand type '" & flxDemandTypes.TextMatrix(i, 3) & "' selected only supports consolidated demands"
                    rsDemandType.Close
                    adoConn1.Close
                    Exit Sub
                End If
                rsDemandType.Close
        Else
        End If
    Next
    adoConn1.Close
    
   SaveSetting "PropertyManagement", "ChoosedOption", "GenerateDemand-c" & CStr(SCID), IIf(optAutoGenSig.Value, "S", "C")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'##################################### DEMAND PREVIEW   ####################
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If szCallingFrom = "Demand Preview" Then
        Picture1.Visible = True
        Picture1.Refresh
        frmDemands3.GenerateDemandPreview
        Picture1.Visible = False
        Picture1.Refresh
        Exit Sub
   End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'##################################### AUTOMATIC DEMAND ####################
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       If szCallingFrom = "Automatic Demand" Then
          Dim adoConn As New ADODB.Connection
          Dim adoRst  As New ADODB.Recordset
          Dim b       As Byte
          Dim szPP    As String
    
       '   connect to database
          adoConn.Open getConnectionString
    
    '           The system will not let user to run automatic demands concurrently by more than
    '           one user.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
          If szCallingFrom = "Automatic Demand" Then
             If IsDemandRunning(szPP, adoConn) Then
                adoConn.Close
                Set adoConn = Nothing
    
                MsgBox "Automatic demands are currently being run for: " & Chr(13) & szPP
                Exit Sub
             End If
          End If
    
    '     Check the posting date period
          b = IsPeriodStatus(txtPostingDate.text, flxClients.TextMatrix(flxClients.row, 1), adoConn)      'Return value 0, 1, 9. 0 --> Close, 1 --> Open, 9 --> Not found
          If b <> 1 Then
    '####### Clear the lock for demand generation ###################################################
             ClearLocks adoConn
    
             adoConn.Close
             Set adoConn = Nothing
             If b = 0 Then
                If MsgBox("The period you wish to post to is closed." & Chr(13) & _
                          "Do you wish to change the posting date?", vbQuestion + vbYesNo, "Posting Period") = vbYes Then
                   FocusControl txtPostingDate
                Else
                   Unload Me
                End If
             End If
             If b = 9 Then
                If MsgBox("The period you wish to post to is not found." & Chr(13) & _
                          "Do you wish to change the posting date?", vbQuestion + vbYesNo, "Posting Period") = vbYes Then
                   FocusControl txtPostingDate
                Else
                   Unload Me
                End If
             End If
             Exit Sub
          End If
    
          If optAutoGenConsolidated.Value = True Then
             Call frmDemands3.GenAutoConDemands(adoConn)
          Else
             Call frmDemands3.GenAutoSngDemands(adoConn)
          End If
        
      'Fix the demadsplit that starts with 2 by anol 20160421
        BUGFIX_DemandSplitRecords_SplitID adoConn
    'end of fix
    '#### Clear the lock for demand generation ###################################################
          ClearLocks adoConn
    
       '  EXPORT all Invoices or Demands into tlbReceipt table *********************************************
          MigrateInvIntoReceipt adoConn
    
    '############################################# EMAIL DEMANDS TO LESSEE #######################
          fmeLoading.Visible = True
          fmeLoading.Refresh
    
          If chkSentAutoEmail.Value Then EmailingDemandAuto adoConn
    
          fmeLoading.Visible = False
    
    '############################################# REFRESH THE DEMAND GRID AFTER AUTO GENERATE #######################
          frmDemands3.LoadFlxGrid frmDemands3.flxDemands, False, adoConn, ""
    
          frmDemands3.flxDemands.row = 0
          frmDemands3.flxDemands.col = 0
    
          adoConn.Close
          Set adoConn = Nothing
       End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'###################################################### PRINT ALL DEMANDS ####################
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       If szCallingFrom = "Print All Demands" Then
          frmDemands3.szSelProperties = AllSelProperties()
    
          frmDemands3.PrintAllDemands
       End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
       frmDemands3.Enabled = True
   'Me.Hide
   
    'Resolved by BOSL.Added by anol
   'issue 530. 4. Demand invoice description not showing in batch receipts details field in batch receipts form.
   'Date 03 Feb 2015
   '  Bring all Invoices or Demands into tlbReceipt table *********************************************
'      MigrateInvIntoReceipt adoConn
       adoConn.Open getConnectionString
      adoConn.Execute "UPDATE demandrecords AS dr, demandsplitrecords AS ds " & _
                      "SET    dr.details = ds.Description " & _
                      "WHERE  dr.demandid=ds.demandid And isnull(dr.details) And ds.splitid=1;"
      adoConn.Execute "UPDATE demandrecords AS dr, demandsplitrecords AS ds " & _
                      "SET    dr.details = ds.Description " & _
                      "WHERE  dr.demandid=ds.demandid And isnull(dr.details) And ds.splitid=2;"

      adoConn.Execute "UPDATE demandrecords AS dr, tlbReceipt AS r " & _
                      "SET    r.details = dr.Details " & _
                      "WHERE  r.demandref=dr.demandid And isnull(r.details);"
                      adoConn.Close
     iSelDemandTypes = 0
'     Unload Me
End Sub

Private Function BUGFIX_DemandSplitRecords_SplitID(Conn1 As ADODB.Connection)
    Dim Rst1 As New ADODB.Recordset
    Rst1.Open "SELECT Q1.* " & _
             "FROM ( " & _
               "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
               "FROM DemandSplitRecords AS DS " & _
               "GROUP BY DS.DemandID " & _
             ") AS Q1 INNER JOIN " & _
             "( " & _
               "SELECT DS.DemandID, COUNT(DSR) AS C " & _
               "FROM DemandSplitRecords AS DS " & _
               "GROUP BY DS.DemandID " & _
             ") AS Q2 ON Q2.DemandID = Q1.DemandID " & _
             "WHERE Q1.M <> Q2.C", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then                                                 'There are demand split without SplitID
      Rst1.Close
      Conn1.Execute "UPDATE DemandSplitRecords " & _
                    "Set SplitID = SplitID - 1 " & _
                    "WHERE DSR IN ( " & _
                    "SELECT DS.DSR " & _
                    "FROM ((" & _
                       "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
                       "FROM DemandSplitRecords AS DS " & _
                       "GROUP BY DS.DemandID " & _
                    ") AS Q1 INNER JOIN " & _
                    "( " & _
                       "SELECT DS.DemandID, COUNT(DSR) AS C " & _
                       "FROM DemandSplitRecords AS DS " & _
                       "GROUP BY DS.DemandID " & _
                    ") AS Q2 ON Q2.DemandID = Q1.DemandID) INNER JOIN " & _
                       "DemandSplitRecords AS DS ON Q1.DemandID = DS.DemandID " & _
                    "WHERE Q1.M <> Q2.C);"

         Conn1.Execute "UPDATE tlbReceiptSplit AS R, DemandSplitRecords AS S " & _
                       "SET R.SplitID = S.SplitID " & _
                       "WHERE R.AllocTranID = S.DSR AND R.Amount = S.TotalAmount;"

       
         
         Exit Function
   End If
   Rst1.Close
   'Freqency ID is not writing for some Demand split records
   'Fixed by anol on 03 Jun 2016
   Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LServiceCharges LS SET DS.FrequencyID=LS.SCFrequency,DS.ChargingMethod=LS.ChargingMethod " & _
   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=2 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
   "AND isnull(LS.Spare3)"
   Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LRentCharges LS SET DS.FrequencyID=LS.BRFrequency,DS.ChargingMethod=LS.spare1 " & _
   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=1 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
   "AND isnull(LS.Spare3)"
    Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LInsuranceCharges LS SET DS.FrequencyID=LS.InsuranceFrequency,DS.ChargingMethod=LS.ChargingType " & _
   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=3 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
   "AND isnull(LS.Spare3)"
  
End Function
Private Function AllSelProperties() As String
   Dim iRow As Integer

   AllSelProperties = ""
   
   For iRow = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(iRow, 0) = "X" Then
         AllSelProperties = "'" & flxProperties.TextMatrix(iRow, 1) & "', " & AllSelProperties
      End If
   Next iRow
   If Len(AllSelProperties) > 0 Then AllSelProperties = Left(AllSelProperties, Len(AllSelProperties) - 2)
End Function

'######################  EmailingDemandAuto  ##############################################################
'  Check the email configuration
'
'  lArrDemand array contains all latest generated demands ID
'     lArrDemand was filled by frmDemands3.GenAutoConDemands/GenAutoSngDemands
'
'  Loop based on number of demands [UBound(lArrDemands)]
'     ? is lessee marked as send demand by email
'        generate the demand
'        create email list - the list does not contains the attachment.
'           the list has only lessee id, email address, demand id
'        convert the demand into pdf
'        add the pdf into the email list
'
'     ? is the lessee marked as send statement by email
'        Create statement in pdf
'        add the pdf into the email list
'
'  Run Send Demands by email     **>>**>>
'
'######################  ******************  ##############################################################
Private Sub EmailingDemandAuto(adoConn As ADODB.Connection)
   Dim szSQL         As String
   Dim iEmailDmd     As Integer
   Dim i             As Integer
   Dim bEmailResult  As Boolean, szLeseList  As String
   Dim rsMinDueDate As New ADODB.Recordset

   Dim adoRstDRecords   As New ADODB.Recordset
   Dim adoRstDRCurrent  As New ADODB.Recordset
   Dim reportApp        As New CRAXDRT.Application
   Dim Report           As CRAXDRT.Report

   On Error GoTo ErrHandling

   If szFromEmail = "" Or szSMTPserver = "" Then
      ShowMsgInTaskBar "Company email or SMTP server IP has not been setup.", "Y", "N"
      Exit Sub
   End If
   If UBound(lArrDemand) < 1 Then
      Exit Sub
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

   MousePointer = vbHourglass
   iLes = 0

   With adoRstDRCurrent
   '*********************************************************************************************
      szLeseList = ""
      For i = 0 To UBound(lArrDemand) - 1
         szSQL = "SELECT DR.DEMANDID AS D_ID, DR.BATCHID AS B_ID, " & _
                        "DR.SAGEACCOUNTNUMBER AS S_AC, " & _
                        "DR.UNITNUMBER AS U_NUM, DR.SOURCE AS SOU, " & _
                        "DR.TRANSACTIONTYPE AS T_TYPE, DSR.DSR, " & _
                        "DR.ISSUEDATE AS IDT, DR.GenerateDemand AS CS, DSR.A_M AS A_M, " & _
                        "DSR.splitid as S_ID, DSR.NOMINALCODEFORAMOUNT AS NCA, " & _
                        "DSR.NOMINALNAMEFORAMOUNT AS NNA, DSR.NOMINALCODEFORVAT AS NCV, " & _
                        "DSR.NOMINALNAMEFORVAT AS NNV, DSR.NOMINALCODEFORTOTAL AS NCT, " & _
                        "DSR.NOMINALNAMEFORTOTAL AS NNT, " & _
                        "DSR.AMOUNT AS AMT, DSR.VATAMOUNT AS VAMT, " & _
                        "DSR.TOTALAMOUNT AS TAMT, DSR.SAGEREF AS SREF, DSR.DUEDATE AS DDT, " & _
                        "DSR.VATMONTH AS VMTH, DSR.TYPEOFDEMAND AS TDM, " & _
                        "DSR.DESCRIPTION AS DESCR, DSR.DEMANDSTATEMENT AS D_ST, " & _
                        "DSR.DATEFROM AS D_FROM, DSR.VAT_CODE, " & _
                        "DSR.DATETO AS D_TO, T.spare1, " & _
                        "DSR.ChargingFigure, DemandTypes.DTGroup, " & _
                        "DemandTypes.EmailInvoiceTemplate AS DRN, DemandTypes.CategoryCode AS CC, " & _
                        "IIF(T.InvoiceTo = 'B', T.Email2, T.Email1) AS EMAIL, T.CombEmail, T.EmailSC "
         szSQL = szSQL & _
                        "FROM DemandRecords as DR, DemandSplitRecords as DSR, Tenants AS T, DemandTypes "
         szSQL = szSQL & _
                        "WHERE DR.DEMANDID = " & CStr(lArrDemand(i)) & " AND " & _
                           "DSR.DEMANDID = DR.DEMANDID AND " & _
                           "T.SAGEACCOUNTNUMBER = DR.SAGEACCOUNTNUMBER AND " & _
                           "DSR.TYPEOFDEMAND = DemandTypes.ID And T.EmailDmd = True AND " & _
                           "IIF(T.InvoiceTo = 'B', T.Email2, T.Email1) <> '' " & _
                        "ORDER BY DSR.splitid;"
'Debug.Print szSQL
         adoRstDRecords.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

         If Not adoRstDRecords.EOF Then
            adoConn.Execute "DELETE * FROM tlbDRCURRENTPRINT;"
            szSQL = "SELECT * FROM tlbDRCURRENTPRINT;"
            .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

'         GET N UPDATE THE INVOICE NUMBER
            iEmailDmd = -1

            While Not adoRstDRecords.EOF
               If ((IsNull(adoRstDRecords.Fields.Item("spare1").Value) Or _
                     InStr(adoRstDRecords.Fields.Item("spare1").Value, "NotD") = 0)) Then
                  .AddNew

                  .Fields.Item("UniqueRefNumber").Value = adoRstDRecords.Fields.Item("D_ID").Value
                  .Fields.Item("Batch").Value = adoRstDRecords.Fields.Item("B_ID").Value
                  .Fields.Item("AutomaticManual").Value = adoRstDRecords.Fields.Item("A_M").Value
                  .Fields.Item("SageAccountNumber").Value = adoRstDRecords.Fields.Item("S_AC").Value
                  szLeseList = szLeseList & "#" & adoRstDRecords.Fields.Item("S_AC").Value & "#"
                  .Fields.Item("UnitNumber").Value = adoRstDRecords.Fields.Item("U_NUM").Value
                  .Fields.Item("NominalCodeforAmount").Value = adoRstDRecords.Fields.Item("NCA").Value
                  .Fields.Item("NominalNameforAmount").Value = adoRstDRecords.Fields.Item("NNA").Value
                  .Fields.Item("NominalCodeforVAT").Value = adoRstDRecords.Fields.Item("NCV").Value
                  .Fields.Item("NominalNameforVAT").Value = adoRstDRecords.Fields.Item("NNV").Value
                  .Fields.Item("NominalCodeforTotal").Value = adoRstDRecords.Fields.Item("NCT").Value
                  .Fields.Item("NominalNameforTotal").Value = adoRstDRecords.Fields.Item("NNT").Value
                  .Fields.Item("Source").Value = adoRstDRecords.Fields.Item("DTGroup").Value
                  .Fields.Item("TransactionType").Value = adoRstDRecords.Fields.Item("T_TYPE").Value
                  .Fields.Item("IssueDate").Value = adoRstDRecords.Fields.Item("IDT").Value
                  .Fields.Item("VATMonth").Value = adoRstDRecords.Fields.Item("VMTH").Value
                  .Fields.Item("Typeofdemand").Value = adoRstDRecords.Fields.Item("TDM").Value
                  .Fields.Item("Reference").Value = adoRstDRecords.Fields.Item("SREF").Value
                  .Fields.Item("Description").Value = adoRstDRecords.Fields.Item("DESCR").Value
                  .Fields.Item("DemandReportName").Value = adoRstDRecords.Fields.Item("DRN").Value
                  .Fields.Item("SptSlNo").Value = adoRstDRecords.Fields.Item("S_ID").Value
                  .Fields.Item("DTCategory").Value = adoRstDRecords.Fields.Item("CC").Value

                  If adoRstDRecords.Fields.Item("D_ST").Value Then
                     .Fields.Item("DueDate").Value = adoRstDRecords.Fields.Item("DDT").Value
                      rsMinDueDate.Open "Select MIN(DUEDATE) as mindate From DemandSplitRecords DSR where DSR.DEMANDID = " + CStr(lArrDemand(i)) + "", adoConn, adOpenKeyset, adLockReadOnly
                      If Not rsMinDueDate.EOF Then
                          .Fields("MinDueDate").Value = rsMinDueDate.Fields("mindate").Value
                      End If
                      rsMinDueDate.Close

                     .Fields.Item("TotalAmount").Value = adoRstDRecords.Fields.Item("TAMT").Value
                     .Fields.Item("VATAmount").Value = adoRstDRecords.Fields.Item("VAMT").Value
                     .Fields.Item("Amount").Value = adoRstDRecords.Fields.Item("AMT").Value
                     .Fields.Item("DemandFrom").Value = adoRstDRecords.Fields.Item("D_FROM").Value
                     .Fields.Item("DemandTo").Value = adoRstDRecords.Fields.Item("D_TO").Value
                  End If

                  .Fields.Item("ChargingFigure").Value = adoRstDRecords.Fields.Item("ChargingFigure").Value
                  .Fields.Item("spare1").Value = adoRstDRecords.Fields.Item("DSR").Value
                  .Fields.Item("spare2").Value = adoRstDRecords.Fields.Item("VAT_CODE").Value
                  .Fields.Item("CombEmail").Value = adoRstDRecords.Fields.Item("CombEmail").Value
                  .Fields.Item("EmailSC").Value = adoRstDRecords.Fields.Item("EmailSC").Value
                  If InStr(InStr(szLeseList, "#" & adoRstDRecords.Fields.Item("S_AC").Value & "#") + 1, _
                                 szLeseList, "#" & adoRstDRecords.Fields.Item("S_AC").Value & "#") = 0 Then
                     .Fields.Item("PrintBBF").Value = True
                  End If
                  .Update

'                  Creating the list of Lessee ID & Email, Client
'                  ReceiverEmailList procedure list all lessee with their email address. only one email address.
'                  ReceiverEmailList procedure does not save the attachment.
                  If adoRstDRecords.Fields.Item("S_ID").Value = "1" Then
                     If adoRstDRecords.Fields.Item("U_NUM").Value <> "" Then
                        iEmailDmd = ReceiverEmailList(adoRstDRecords.Fields.Item("S_AC").Value, _
                                                      adoRstDRecords.Fields.Item("EMAIL").Value, _
                                                      adoRstDRecords.Fields.Item("D_ID").Value, _
                                                      GetClientNameByUnit(adoRstDRecords.Fields.Item("U_NUM").Value, adoConn))
                     Else
                        iEmailDmd = ReceiverEmailList(adoRstDRecords.Fields.Item("S_AC").Value, _
                                                      adoRstDRecords.Fields.Item("EMAIL").Value, _
                                                      adoRstDRecords.Fields.Item("D_ID").Value)
                     End If
                  End If
               End If
               adoRstDRecords.MoveNext
            Wend
            adoRstDRecords.Close
            .Close

'           Converting the demand into pdf
'              if >> Creating Statement for attachment
'              if >> Creating SC Statement for attachment
            If iEmailDmd >= 0 Then
               szSQL = "SELECT   DemandReportName, Source, SageAccountNumber, CombEmail, EmailSC, UniqueRefNumber " & _
                       "FROM     tlbDRCurrentPrint " & _
                       "GROUP BY DemandReportName, Source, SageAccountNumber, CombEmail, EmailSC, UniqueRefNumber;"
               .Open szSQL, adoConn, adOpenStatic, adLockReadOnly

               While Not .EOF
'                 Passing the FROM and TO date values to Crystal Reports
                  Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & .Fields.Item("DemandReportName").Value)
                  Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

                  Report.EnableParameterPrompting = False
                  If Report.HasSavedData Then Report.DiscardSavedData

                  Report.ParameterFields(1).AddCurrentValue .Fields.Item("Source").Value

'                  Transfer report into PDF file
'                    Create the pdf file name of the report
                  szSQL = .Fields.Item("SageAccountNumber").Value & "_" & UniqueID() & ".pdf"

'                 Transfer demand report into PDF file
                  Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\" & szSQL
                  Report.ExportOptions.DestinationType = crEDTDiskFile
                  Report.ExportOptions.FormatType = crEFTPortableDocFormat
                  Report.ExportOptions.PDFExportAllPages = True
                  Report.Export False
                  Set Report = Nothing

'                    Attaching the demand to the email
                  SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL, .Fields.Item("SageAccountNumber").Value

'                  Statement will go with demand?
                  If .Fields.Item("CombEmail").Value Then
'                    Statement will go with demand
'                    Create the pdf of the statement
                     szSQL = .Fields.Item("SageAccountNumber").Value & "_LEEST_" & UniqueID() & ".pdf"

                     CreatePDF_Statement .Fields.Item("SageAccountNumber").Value, DB_PATH & "\AllStuff\Temp\" & szSQL, adoConn

'                    Attaching the statement to the email
                     SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL, .Fields.Item("SageAccountNumber").Value
                  End If
                  EmailDelay 10

'                 SC Statement will go with demand?
                  If .Fields.Item("EmailSC").Value And chkSCS.Value Then
'                     Statement will go with demand
'                     If its single demand then check ?has SC St generated and attached?
                     If optAutoGenSig.Value Then
                        If Not SC_St_Attached(.Fields.Item("SageAccountNumber").Value) Then

'                       Create the pdf of the statement
                           szSQL = .Fields.Item("SageAccountNumber").Value & "_SCST_" & UniqueID() & ".pdf"

                           CreateSC_Statement .Fields.Item("SageAccountNumber").Value, DB_PATH & "\AllStuff\Temp\" & szSQL

'                             Attaching the statement to the email
                           SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL, .Fields.Item("SageAccountNumber").Value
'                           ConnectStatementWithDemand .Fields.Item("UniqueRefNumber").Value, szSQL, adoConn
                        End If
                     Else
                        szSQL = .Fields.Item("SageAccountNumber").Value & "_SCST_" & UniqueID() & ".pdf"

                        CreateSC_Statement .Fields.Item("SageAccountNumber").Value, DB_PATH & "\AllStuff\Temp\" & szSQL

'                          Attaching the statement to the email
                        SaveAttachment DB_PATH & "\AllStuff\Temp\" & szSQL, .Fields.Item("SageAccountNumber").Value
'                        ConnectStatementWithDemand .Fields.Item("UniqueRefNumber").Value, szSQL, adoConn
                     End If
                  End If
                  .MoveNext
               Wend
               .Close
            End If
            EmailDelay 10
         Else
            adoRstDRecords.Close
         End If               'adoRstDRecords.EOF
      Next i
   End With

   Set adoRstDRCurrent = Nothing

'  Finally send demands by email
   bEmailResult = SendDemandByE_Mail(adoConn)

   MousePointer = vbDefault
   If iLes > 0 And bEmailResult Then
      MsgBox "Email sent.", vbInformation + vbOKOnly, "Demands"
   Else
      MsgBox "No email sent.", vbExclamation + vbOKOnly, "Demands"
   End If
   Exit Sub
   
ErrHandling:
   
End Sub

Private Sub ConnectStatementWithDemand(lDemandID As Long, szStatement As String, adoConn As Connection)
'we have decided that we will not save it.
'when user will send manually then system will regenerate the sc statement.
'system will ask sc statement's st and end date.
'   adoConn.Execute "UPDATE DemandRecords " & _
                   "SET SentStName = '" & szStatement & "' " & _
                   "WHERE DemandID = " & lDemandID & ";"
End Sub

Private Function SC_St_Attached(szLessee As String) As Boolean
   SC_St_Attached = False

   Dim i          As Integer
   Dim j          As Integer
   Dim szaTemp()  As String

   On Error GoTo DeclareArray

   For i = 0 To iLes - 1
      If uLessee(i).szLesseeID = szLessee Then
         For j = 1 To uLessee(i).colAtt.Count
            szaTemp = Split(uLessee(i).colAtt(j), "\")
            szaTemp = Split(szaTemp(UBound(szaTemp)), "_")
            If szaTemp(1) = "SCST" Then
               SC_St_Attached = True
               Exit For
            End If
         Next j

         If SC_St_Attached Then Exit For
      End If
   Next i

   Exit Function
DeclareArray:
   Set uLessee(i).colAtt = New Collection
End Function

Private Sub CreateSC_Statement(szLessee As String, szFileName As String)
   Dim reportApp  As New CRAXDRT.Application
   Dim Report     As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SC_St_AutoEmail.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue szLessee
   Report.ParameterFields(2).AddCurrentValue CDate(txtSCDateFrom.text)
   Report.ParameterFields(3).AddCurrentValue CDate(txtSCDateTo.text)

'   Transfer report into PDF file
   Report.ExportOptions.DiskFileName = szFileName
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing
End Sub

Private Sub CreatePDF_Statement(szLessee As String, szFileName As String, adoConn As ADODB.Connection)
   Dim szSQL      As String
   Dim dLesBal    As Double
   Dim adoRst     As New ADODB.Recordset
   Dim reportApp  As New CRAXDRT.Application
   Dim Report     As CRAXDRT.Report

   szSQL = "SELECT T.Name, U.UnitName, P.PropertyName, C.ClientName " & _
           "FROM   Tenants AS T,  LeaseDetails AS L, Units AS U, Property AS P, Client AS C " & _
           "WHERE  Status = TRUE AND T.SageAccountNumber = '" & szLessee & "' AND " & _
               "   T.SageAccountNumber = L.SageAccountNumber AND " & _
               "   L.UnitNumber = U.UnitNumber AND " & _
               "   U.PropertyID = P.PropertyID AND " & _
               "   P.ClientID = C.ClientID;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      Set adoRst = Nothing
      Exit Sub
   End If

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\LesseeAcHistory.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   If Report.HasSavedData Then Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue szLessee
   Report.ParameterFields(2).AddCurrentValue adoRst.Fields.Item("Name").Value
   Report.ParameterFields(3).AddCurrentValue adoRst.Fields.Item("ClientName").Value
   Report.ParameterFields(4).AddCurrentValue adoRst.Fields.Item("PropertyName").Value
   Report.ParameterFields(5).AddCurrentValue adoRst.Fields.Item("UnitName").Value
   dLesBal = CDbl(LesseeAccountBalance(adoConn, szLessee))
   Report.ParameterFields(6).AddCurrentValue dLesBal

   adoRst.Close
   Set adoRst = Nothing

   'Transfer report into PDF file
   Report.ExportOptions.DiskFileName = szFileName
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False
   Set Report = Nothing
End Sub

Private Function SendDemandByE_Mail(adoConn As ADODB.Connection) As Boolean
   Dim i As Integer
   Dim szSub As String, szBody As String

   For i = 0 To iLes - 1
      szSub = "Rent and Service Charge demands from " & uLessee(i).szClient
      szBody = "Please find attachment your rent and service charge demands for payment." & (Chr(13) + Chr(10)) & _
               (Chr(13) + Chr(10)) & _
               "Yours sincerely," & (Chr(13) + Chr(10)) & _
               (Chr(13) + Chr(10)) & _
               uLessee(i).szClient
'SendEmail(ByVal strSender As String, _
'          ByVal strRecipient As String, _
'          ByVal strSubject As String, _
'          ByVal strBody As String, _
'          Optional ByVal strCc As String, _
'          Optional ByVal strBcc As String, _
'          Optional ByVal colAttachments As Collection, _
'          Optional ByVal szRecipID As String, _
'          Optional ByVal szEmailType As String, _
'          Optional ByVal szRef As String) As Boolean
      SendDemandByE_Mail = SendEmail(szFromEmail, Trim(uLessee(i).szLesseeEmail), _
                                     szSub, _
                                     szBody, , , _
                                     uLessee(i).colAtt, uLessee(i).szLesseeID, "SI")

      adoConn.Execute "UPDATE DemandRecords " & _
                      "SET SentByEmail =" & IIf(SendDemandByE_Mail, "1", "0") & " " & _
                      "WHERE DemandID IN (" & ListDemandID(i) & ");"
   Next i
End Function

Private Function ListDemandID(i As Integer) As String
   Dim j As Integer

   For j = 1 To uLessee(i).colURN.Count
      ListDemandID = ListDemandID & CStr(uLessee(i).colURN(j)) & ", "
   Next j

   ListDemandID = Left(ListDemandID, Len(ListDemandID) - 2)
End Function

Private Function ReceiverEmailList(szLessee As String, szEmail As String, lURN As Long, Optional szClient As String) As Integer
   Dim i As Integer

   ReceiverEmailList = 0
   If iLes = 0 Then
      iLes = 1
      ReDim uLessee(0) As SendDemandByEmail
      Set uLessee(0).colURN = New Collection

      uLessee(0).szLesseeID = szLessee
      uLessee(0).szLesseeEmail = szEmail
      uLessee(0).szClient = szClient
      uLessee(0).colURN.Add lURN
      Exit Function
   End If
   
   For i = 0 To UBound(uLessee)
      If uLessee(i).szLesseeID = szLessee Then
         uLessee(i).colURN.Add lURN
         Exit For
      End If
   Next i

   If i > UBound(uLessee) Then
      ReDim Preserve uLessee(iLes) As SendDemandByEmail
      Set uLessee(iLes).colURN = New Collection

      uLessee(iLes).szLesseeID = szLessee
      uLessee(iLes).szLesseeEmail = szEmail
      uLessee(iLes).szClient = szClient
      uLessee(iLes).colURN.Add lURN
      ReceiverEmailList = iLes
      iLes = iLes + 1
   End If
End Function

Private Sub SaveAttachment(szFile As String, szLessee As String)
   Dim i As Integer

   On Error GoTo DeclareArray
   For i = 0 To iLes - 1
      If uLessee(i).szLesseeID = szLessee Then
         uLessee(i).colAtt.Add szFile
         Exit For
      End If
   Next i
   Exit Sub

DeclareArray:
   Set uLessee(i).colAtt = New Collection
   uLessee(i).colAtt.Add szFile
End Sub

Private Sub flxClients_RowColChange()
'   SelectFlxGridRow 0, flxClients, flxClients.row
   SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
   iSelClient = 1

   FilterProperties
   flxDemandTypes.Clear
End Sub

Private Sub flxDemandTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   flxDemandTypes.MousePointer = vbArrow
End Sub

Private Sub ConfigFlxGrids()
   Dim szHeader As String

   flxClients.RowHeight(0) = 0
   flxClients.ColWidth(0) = 300
   flxClients.ColWidth(1) = 1000
   flxClients.ColWidth(2) = 2350

   szHeader$ = "<|<|<|<"
   flxProperties.FormatString = szHeader
   flxProperties.Cols = 4
   flxProperties.RowHeight(0) = 0
   flxProperties.ColWidth(0) = 300                   '"X"
   flxProperties.ColWidth(1) = 1000                'Property ID
   flxProperties.ColWidth(2) = 2350                'Property Name
   flxProperties.ColWidth(3) = 0                   'Client ID

   flxDemandTypes.Cols = 5
   flxDemandTypes.RowHeight(0) = 0
   flxDemandTypes.ColWidth(0) = 300                  '"X"
   flxDemandTypes.ColWidth(1) = 900                  'Property ID
   flxDemandTypes.ColWidth(2) = 0               'Demand Type ID
   flxDemandTypes.ColAlignment(2) = vbRightJustify
   flxDemandTypes.ColWidth(3) = 4000               'Demand Type Name
   flxDemandTypes.ColWidth(4) = 0                  'Demand Category

   flxCategory.RowHeight(0) = 0
   flxCategory.ColWidth(0) = 0
   flxCategory.ColWidth(1) = 0
   flxCategory.ColWidth(2) = flxCategory.Width - 250
End Sub

Private Sub flxDemandTypes_RowColChange()
'   SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
'   If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
'      iSelDemandTypes = iSelDemandTypes + 1
'   Else
'      iSelDemandTypes = iSelDemandTypes - 1
'   End If
'Debug.Print iSelDemandTypes
End Sub

Private Sub flxProperties_RowColChange()
'   SelectFlxGridRow 0, flxProperties, flxProperties.row
'   If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'      iSelProperties = iSelProperties + 1
'   Else
'      iSelProperties = iSelProperties - 1
'   End If
'
'   FilterDemandTypes
End Sub

Private Sub Form_Load()
   Dim szChoice As String

'   Me.Height = 8550
'   Me.Width = 7710
'   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fraSC_St.BackColor = Me.BackColor
   chkProp.BackColor = Me.BackColor
   chkDT.BackColor = Me.BackColor
   optAutoGenSig.BackColor = Me.BackColor
   optAutoGenConsolidated.BackColor = Me.BackColor
   optAutoGenConsolidated.BackColor = fraSC_St.BackColor
   chkSCS.BackColor = Me.BackColor

   ConfigFlxGrids          'flxClients, flxProperties, flxDemandTypes, flxCategory

   LoadFlxGrids            'flxClients, flxProperties, flxDemandTypes, flxCategory, cboGDPFreq

   iSelClient = 0

   txtInitialIssueDate.text = Format(Date, "dd/mm/yyyy")
   txtPostingDate.text = txtInitialIssueDate.text
   If Not frmMMain.IsRibbonVersion Then
      txtPostingDate.Visible = False
      Label19(0).Visible = False
   End If

   szChoice = GetSetting("PropertyManagement", "ChoosedOption", "GenerateDemand-c" & CStr(SCID))
   If szChoice <> "S" And szChoice <> "C" Then
      optAutoGenSig.Value = True
   Else
      optAutoGenSig.Value = IIf(szChoice = "S", True, False)
      optAutoGenConsolidated.Value = IIf(szChoice = "C", True, False)
   End If
   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadFlxGrids()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT;"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   r = 1
   flxClients.Rows = 1
   flxClients.ColAlignment(1) = vbLeftJustify

   While Not adoRst.EOF
      flxClients.AddItem ""
      flxClients.TextMatrix(r, 1) = adoRst.Fields.Item("CLIENTID").Value
      flxClients.TextMatrix(r, 2) = adoRst.Fields.Item("CLIENTNAME").Value
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
'------------------------------------------------------------------------------------------
   szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY " & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 0
   ReDim szaProp(adoRst.RecordCount) As String

   While Not adoRst.EOF
      szaProp(r) = adoRst.Fields.Item("PROPERTYID").Value & " ## " & _
                   adoRst.Fields.Item("PROPERTYNAME").Value & " ## " & _
                   adoRst.Fields.Item("ClientID").Value
      r = r + 1
      adoRst.MoveNext
   Wend
   iProperty = r

   adoRst.Close
'------------------------------------------------------------------------------------------
   szSQL = "SELECT PropertyID, ID, Type, CategoryCode " & _
           "FROM   DemandTypes " & _
           "ORDER BY ID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 0
   ReDim szaDT(adoRst.RecordCount) As String

   While Not adoRst.EOF
      szaDT(r) = adoRst.Fields.Item(0).Value & " ## " & _
                 adoRst.Fields.Item(1).Value & " ## " & _
                 adoRst.Fields.Item(2).Value & " ## " & _
                 adoRst.Fields.Item(3).Value
                 
                   Debug.Print szaDT(r)
      r = r + 1
      adoRst.MoveNext
   Wend
   iDT = r

   adoRst.Close
'------------------------------------------------------------------------------------------
   szSQL = "SELECT Code, Value " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'DCTG' " & _
           "ORDER BY Code;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 1
   flxCategory.Rows = 1

   While Not adoRst.EOF
      flxCategory.AddItem ""
      flxCategory.TextMatrix(r, 1) = adoRst.Fields.Item(0).Value
      flxCategory.TextMatrix(r, 2) = adoRst.Fields.Item(1).Value
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   For r = 1 To flxCategory.Rows - 1
      SelectFlxGridRow 0, flxCategory, r
   Next r
   iSelDemandCategory = r - 1
'------------------------------------------------------------------------------------------
   FillFrequency adoConn

   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub FilterProperties()
   Dim i As Integer, j As Integer, r As Integer
   Dim szaTemp() As String

   flxProperties.Clear
   iSelProperties = 0
   chkProp.Value = 0
   chkDT.Value = 0

   flxProperties.Rows = 1
   r = 1
   For i = 0 To iProperty - 1
      szaTemp = Split(szaProp(i), " ## ")
      
      For j = 0 To flxClients.Rows - 1
         If flxClients.TextMatrix(j, 0) = "X" And flxClients.TextMatrix(j, 1) = szaTemp(2) Then
            flxProperties.AddItem ""
            flxProperties.TextMatrix(r, 1) = szaTemp(0)
            flxProperties.TextMatrix(r, 2) = szaTemp(1)
            flxProperties.TextMatrix(r, 3) = szaTemp(2)
            'added by anol 16 Aug 2016
            flxProperties.row = 1
            r = r + 1
         End If
      Next j
   Next i
End Sub

Private Sub flxClients_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxClients.MousePointer = vbArrow
End Sub

Private Sub FilterDemandTypes()
   Dim i As Integer, j As Integer, r As Integer
   Dim szaTemp() As String

   flxDemandTypes.Clear
   chkDT.Value = 0

   flxDemandTypes.Rows = 1
   r = 1
   For i = 0 To iDT - 1
      szaTemp = Split(szaDT(i), " ## ")
        Debug.Print szaDT(i)
      For j = 1 To flxProperties.Rows - 1
        'Debug.Print szaTemp(0) & szaTemp(1) & szaTemp(1)
        If szaTemp(0) = "23BAR" Then
            Debug.Print szaTemp(0) & szaTemp(1) & szaTemp(1)
        End If
         If flxProperties.TextMatrix(j, 0) = "X" And (flxProperties.TextMatrix(j, 1) = szaTemp(0) Or szaTemp(0) = "ALL") Then
            If DemandCategorySelected(CInt(szaTemp(3))) Then
               flxDemandTypes.AddItem ""
               flxDemandTypes.TextMatrix(r, 1) = szaTemp(0)
               flxDemandTypes.TextMatrix(r, 2) = szaTemp(1)
               flxDemandTypes.TextMatrix(r, 3) = szaTemp(2)
               flxDemandTypes.TextMatrix(r, 4) = szaTemp(3)
               'added by anol 16 Aug 2016
               flxDemandTypes.row = 1
               r = r + 1
               If szaTemp(0) = "ALL" Then Exit For
            End If
         End If
      Next j
   Next i
End Sub

Private Function DemandCategorySelected(iCat As Integer) As Boolean
   Dim i As Integer

   DemandCategorySelected = False

   For i = 1 To flxCategory.Rows - 1
      If flxCategory.TextMatrix(i, 0) = "X" And flxCategory.TextMatrix(i, 1) = iCat Then
         DemandCategorySelected = True
         Exit For
      End If
   Next i
End Function

Public Sub FillFrequency(Conn1 As ADODB.Connection)
   Dim i As Integer, SQLStr1 As String
   Dim Rst1 As New ADODB.Recordset
   Dim Data() As String

   SQLStr1 = "SELECT * FROM Frequencies"
   Rst1.Open SQLStr1, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      ReDim Preserve Data(1, Rst1.RecordCount) As String
      Data(0, 0) = 0
      Data(1, 0) = "ALL Frequencies"
      i = 1
      While Rst1.EOF = False
         Data(0, i) = Rst1!ID
         Data(1, i) = Rst1!Frequency
         i = i + 1
         Rst1.MoveNext
      Wend
      cboGDPFreq.Clear

      cboGDPFreq.Column() = Data()
      cboGDPFreq.ListIndex = 0
   End If

   Rst1.Close
   Set Rst1 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmDemands3.Enabled = True
'   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtPostingDate_Change()
   TextBoxChangeDate txtPostingDate
End Sub

Private Sub txtPostingDate_GotFocus()
   SelTxtInCtrl txtPostingDate
End Sub

Private Sub txtPostingDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtPostingDate, KeyAscii
End Sub

Private Sub txtPostingDate_LostFocus()
   TextBoxFormatDate txtPostingDate
End Sub

Private Sub txtSCDateFrom_Change()
   TextBoxChangeDate txtSCDateFrom
End Sub

Private Sub txtSCDateFrom_GotFocus()
   SelTxtInCtrl txtSCDateFrom
End Sub

Private Sub txtSCDateFrom_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSCDateFrom, KeyAscii
End Sub

Private Sub txtSCDateFrom_LostFocus()
   TextBoxFormatDate txtSCDateFrom
End Sub

Private Sub txtSCDateTo_Change()
   TextBoxChangeDate txtSCDateTo
End Sub

Private Sub txtSCDateTo_GotFocus()
   SelTxtInCtrl txtSCDateTo
End Sub

Private Sub txtSCDateTo_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSCDateTo, KeyAscii
End Sub

Private Sub txtSCDateTo_LostFocus()
   TextBoxFormatDate txtSCDateTo
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
