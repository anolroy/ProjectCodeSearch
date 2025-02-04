VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBankTransfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Transfer"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   2925
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   34
      Top             =   4455
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   37
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   36
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
         TabIndex        =   35
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
      BackColor       =   &H00E0E0E0&
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
      Height          =   4650
      Index           =   1
      Left            =   40
      TabIndex        =   14
      Top             =   40
      Width           =   11295
      Begin VB.CommandButton cmdFund 
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
         Left            =   5175
         TabIndex        =   4
         Top             =   3555
         Width           =   300
      End
      Begin VB.CommandButton cmdAccountTo 
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
         Left            =   7560
         TabIndex        =   3
         Top             =   1440
         Width           =   300
      End
      Begin VB.CommandButton cmdAccountFrom 
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
         Left            =   2880
         TabIndex        =   2
         Top             =   1440
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
         Left            =   5175
         TabIndex        =   0
         Top             =   135
         Width           =   300
      End
      Begin VB.TextBox txtBkTrAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   4335
         Width           =   1095
      End
      Begin MSForms.TextBox txtTBalance 
         Height          =   285
         Left            =   6480
         TabIndex        =   50
         Top             =   3060
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBalance 
         Height          =   285
         Left            =   1800
         TabIndex        =   49
         Top             =   3060
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTAccountName 
         Height          =   285
         Left            =   6480
         TabIndex        =   48
         Top             =   2655
         Width           =   3735
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "6588;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountName 
         Height          =   285
         Left            =   1800
         TabIndex        =   47
         Top             =   2655
         Width           =   3465
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "6112;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTSortCode 
         Height          =   285
         Left            =   6480
         TabIndex        =   46
         Top             =   2250
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSortCode 
         Height          =   285
         Left            =   1800
         TabIndex        =   45
         Top             =   2250
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTNominalCode 
         Height          =   285
         Left            =   6480
         TabIndex        =   44
         Top             =   1440
         Width           =   1035
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1826;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNominalCode 
         Height          =   285
         Left            =   1800
         TabIndex        =   43
         Top             =   1440
         Width           =   1080
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "1905;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   42
         Top             =   990
         Width           =   1200
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Account From:"
         Size            =   "2117;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtFund 
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Top             =   3555
         Width           =   3420
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6032;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountTo 
         Height          =   285
         Left            =   6480
         TabIndex        =   40
         Top             =   1845
         Width           =   3690
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6509;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAccountFrom 
         Height          =   285
         Left            =   1800
         TabIndex        =   39
         Top             =   1845
         Width           =   3420
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6032;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   135
         Width           =   3420
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6032;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   3020
         TabIndex        =   30
         Top             =   480
         Width           =   225
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "397;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   2
         Left            =   6480
         TabIndex        =   29
         Top             =   990
         Width           =   1110
         BackColor       =   16768960
         VariousPropertyBits=   8388627
         Caption         =   "Account To:"
         Size            =   "1958;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   255
         Index           =   14
         Left            =   600
         TabIndex        =   28
         Top             =   2670
         Width           =   1080
         BackColor       =   16768960
         VariousPropertyBits=   8388627
         Caption         =   "Account Name"
         Size            =   "1905;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   27
         Top             =   2250
         Width           =   720
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Sort Code"
         Size            =   "1270;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   26
         Top             =   1425
         Width           =   1005
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Nominal Code"
         Size            =   "1773;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   25
         Top             =   1860
         Width           =   840
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Account No"
         Size            =   "1482;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblHClientID 
         Height          =   315
         Left            =   7680
         TabIndex        =   24
         Top             =   195
         Visible         =   0   'False
         Width           =   3135
         BackColor       =   16777215
         Size            =   "5530;556"
         BorderStyle     =   1
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBkTrDes 
         Height          =   345
         Left            =   7440
         TabIndex        =   23
         Top             =   4335
         Visible         =   0   'False
         Width           =   3525
         VariousPropertyBits=   -1400879077
         Size            =   "6218;609"
         Value           =   "BANK TRANSFER"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBkTrRef 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   3945
         Width           =   3495
         VariousPropertyBits=   746604571
         MaxLength       =   20
         Size            =   "6165;503"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBkTrDate 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "2143;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   22
         Top             =   3555
         Width           =   405
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Fund:"
         Size            =   "714;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   21
         Top             =   4335
         Width           =   450
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Value:"
         Size            =   "794;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   390
         Index           =   9
         Left            =   6600
         TabIndex        =   20
         Top             =   4245
         Visible         =   0   'False
         Width           =   885
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Hidden Description:"
         Size            =   "1561;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   19
         Top             =   3945
         Width           =   765
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Reference:"
         Size            =   "1349;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   390
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Date:"
         Size            =   "688;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   390
         Index           =   12
         Left            =   6960
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   675
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Hidden Client ID:"
         Size            =   "1191;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   480
         BackColor       =   16768960
         VariousPropertyBits=   276824083
         Caption         =   "Client:"
         Size            =   "847;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   15
         Top             =   3090
         Width           =   600
         BackColor       =   16768960
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "1058;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   2175
         Index           =   22
         Left            =   480
         Top             =   1290
         Width           =   10335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00404040&
         BorderWidth     =   3
         Height          =   2175
         Index           =   23
         Left            =   480
         Top             =   1290
         Width           =   10335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   3
      Left            =   40
      TabIndex        =   13
      Top             =   4680
      Width           =   11295
      Begin VB.CommandButton cmdBTSave 
         BackColor       =   &H00F0F0F0&
         Caption         =   "&Save"
         Height          =   400
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton cmdCloseBk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C&lose"
         Height          =   400
         Index           =   1
         Left            =   9360
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1450
      End
   End
   Begin VB.Label lblOverDraftAmount 
      Height          =   150
      Left            =   90
      TabIndex        =   32
      Top             =   5490
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label isOverDratAllowed 
      Height          =   150
      Left            =   855
      TabIndex        =   31
      Top             =   5490
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "frmBankTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FrmBankTransfer_CALLING_FROM    As String
Private szaClientID()                  As String
Dim sTextBox As String
'Private Sub cboANF_Click()
'   On Error GoTo ErrorHandler
'
'   If txtAccountFrom.text = txtAccountTo.text Then
'      MsgBox "Bank Transfter can not possible between same account.", vbCritical + vbOKOnly, "Bank transfer"
'      txtNominalCode.text = ""
'      txtBalance.text = ""
'      txtSortCode.text = ""
'      txtAccountName.text = ""
'      txtAccountFrom.text = ""
'      isOverDratAllowed.Caption = ""
'      lblOverDraftAmount.Caption = ""
'   Else
'      txtNominalCode.text = IIf(IsNull(cboANF.Column(1)), "", cboANF.Column(1)) 'Bank ID
'      'txtBalance.text = IIf(IsNull(cboANF.Column(2)), "", cboANF.Column(2))
'      txtSortCode.text = IIf(IsNull(cboANF.Column(3)), "", cboANF.Column(3))
'      isOverDratAllowed.Caption = IIf(IsNull(cboANF.Column(5)), "", cboANF.Column(5))
'      lblOverDraftAmount.Caption = IIf(IsNull(cboANF.Column(6)), "", cboANF.Column(6))
'      'Resolved by BOSL
'      '0000546: Bank Transfers not showing bank balance
'      'Modified by anol 12 Mar 2015
'      Dim adoConn As New ADODB.Connection
'      adoConn.Open getConnectionString
'      txtBalance.text = Format(BankAccBalance(adoConn, cboANF.Column(1), txtClientList.Tag), "0.00") 'For finding bank balance
'      adoConn.Close
'      txtAccountName.text = IIf(IsNull(cboANF.Column(4)), "", cboANF.Column(4))
'      cmdAccountTo.SetFocus
'   End If
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'End Sub

'Private Sub cboANT_Click()
'   On Error GoTo ErrorHandler
'
'   cboFundBankTransf.Enabled = True
'
'   If cboANF.text = cboANT.text Then
'      MsgBox "Bank Transfter can not possible between same account.", vbCritical + vbOKOnly, "Bank transfer"
'      txtTNominalCode.text = ""
'      txtTBalance.text = ""
'      txtTSortCode.text = ""
'      txtTAccountName.text = ""
'      cboANT.text = ""
'   Else
'      txtTNominalCode.text = IIf(IsNull(cboANT.Column(1)), "", cboANT.Column(1))
'      'txtTBalance.text = IIf(IsNull(cboANT.Column(2)), "", cboANT.Column(2))
'      'Resolved by BOSL
'      '0000546: Bank Transfers not showing bank balance
'      'Modified by anol 12 Mar 2015
'      Dim adoConn As New ADODB.Connection
'      adoConn.Open getConnectionString
'      txtTBalance.text = Format(BankAccBalance(adoConn, cboANT.Column(1), txtClientList.Tag), "0.00") 'For finding bank balance
'      adoConn.Close
'      txtTSortCode.text = IIf(IsNull(cboANT.Column(3)), "", cboANT.Column(3))
'      txtTAccountName.text = IIf(IsNull(cboANT.Column(4)), "", cboANT.Column(4))
'   End If
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'End Sub

Private Sub cmdAccountFrom_Click()
         picClient.Left = 269.029
         picClient.Top = 855.299
         sTextBox = "2"
         LoadAccountGrid
         Frame5(1).Enabled = False
         Frame5(3).Enabled = False
         picClient.Visible = True
         txtSearchClientID.SetFocus
End Sub

Private Sub cmdAccountTo_Click()
         picClient.Left = 5069.029
         picClient.Top = 855.299
         sTextBox = "3"
         LoadAccountGrid
         Frame5(1).Enabled = False
         Frame5(3).Enabled = False
         picClient.Visible = True
         txtSearchClientID.SetFocus
End Sub

Private Sub cmdClientList_Click()
         picClient.Left = 269.029
         picClient.Top = 255.299
         sTextBox = "1"
         LoadflxClient
         Frame5(1).Enabled = False
         Frame5(3).Enabled = False
         picClient.Visible = True
         txtSearchClientID.SetFocus
End Sub
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
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new

   
   adoConn.Open getConnectionString
    szSQL = "SELECT  tlbClientBanks.CLIENT_ID, Client.ClientName From tlbClientBanks, Client " & _
            "Where tlbClientBanks.CLIENT_ID = Client.ClientID " & _
            "GROUP BY  tlbClientBanks.CLIENT_ID, Client.ClientName " & _
            "HAVING COUNT(Client.ClientName) > 1 ORDER BY tlbClientBanks.CLIENT_ID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'           flxClient.TextMatrix(1, 0) = ""
'           flxClient.TextMatrix(1, 1) = "ALL"
'           flxClient.TextMatrix(1, 2) = "All Client"
'           flxClient.AddItem ""
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

Private Sub cmdFund_Click()
         picClient.Left = 269.029
         picClient.Top = 855.299
         sTextBox = "4"
         loadflxFund
         Frame5(1).Enabled = False
         Frame5(3).Enabled = False
         picClient.Visible = True
         txtSearchClientID.SetFocus
End Sub
Private Sub loadflxFund()

    Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 50
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
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
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
   'New
   
'   picClient.Width = 5295
'   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   'End of new
   
   adoConn.Open getConnectionString
    szSQL = "SELECT FundID, FundCode,FundName FROM FUND Order by FundCode;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item("FundID").Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundCode").Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundName").Value
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
        Frame5(1).Enabled = True
        Frame5(3).Enabled = True
        picClient.Visible = False
        
        cmdClientList.SetFocus
End Sub

Private Sub cmdtAccountTo_Click()
         picClient.Left = 5069.029
         picClient.Top = 855.299
         sTextBox = "3"
         LoadAccountGrid
         Frame5(1).Enabled = False
         Frame5(3).Enabled = False
         picClient.Visible = True
         txtSearchClientID.SetFocus
End Sub

Private Sub txtFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 13 Then
            cmdFund.SetFocus
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
'    If KeyAscii = 27 Then
'         picClient.Visible = False
'          Frame1.Enabled = True
'          Frame2.Enabled = True
'          If sTextBox = "1" Then
'                 cmdClientList.SetFocus
''           ElseIf sTextBox = "2" Then
''                cmdproperty.SetFocus
''           ElseIf sTextBox = "3" Then
''                cmdFundLookUp.SetFocus
'           End If
'    End If
If KeyAscii = 27 Then
        picClient.Visible = False
         Frame5(1).Enabled = True
         Frame5(3).Enabled = True
         If sTextBox = "1" Then
            cmdClientList.SetFocus
         ElseIf sTextBox = "2" Then
            cmdAccountFrom.SetFocus
         ElseIf sTextBox = "3" Then
            cmdAccountTo.SetFocus
         ElseIf sTextBox = "4" Then
           cmdFund.SetFocus
         
         ElseIf sTextBox = "5" Then
           ' cmdNC.SetFocus
         ElseIf sTextBox = "6" Then
            'cmdFund.SetFocus
         ElseIf sTextBox = "7" Then
            'cmdVATCode.SetFocus
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

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub

            
Private Sub flxClient_Click()
    Frame5(1).Enabled = True
    Frame5(3).Enabled = True
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    If sTextBox = "1" Then
            txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
            txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
            'cboClientName_Click
            txtNominalCode.text = ""
           txtBalance.text = ""
           txtSortCode.text = ""
           txtAccountName.text = ""
           txtAccountFrom.text = ""
           isOverDratAllowed.Caption = ""
           lblOverDraftAmount.Caption = ""
           txtTNominalCode.text = ""
            txtTBalance.text = ""
            txtTSortCode.text = ""
            txtTAccountName.text = ""
            txtAccountTo.text = ""
            txtAccountFrom.text = ""
            txtBkTrDate.SetFocus
    ElseIf sTextBox = "2" Then
            'txtAccountFrom.Tag = flxClient.TextMatrix(flxClient.row, 2)
         txtAccountFrom.text = flxClient.TextMatrix(flxClient.row, 2)
'         If Trim(txtAccountFrom.text) = "" Then
'            MsgBox "No clients found with more than one bank account. Please go to Clients > Bank Details Tab to add more bank accounts.", vbInformation, "No Clients with Multiple Bank Accounts"
'            Exit Sub
'         End If
        If txtAccountFrom.text = txtAccountTo.text Then
           MsgBox " A Bank Transfer is not possible between the same Bank Account. Please select another Bank Account", vbCritical + vbOKOnly, "Bank transfer"
           txtNominalCode.text = ""
           txtBalance.text = ""
           txtSortCode.text = ""
           txtAccountName.text = ""
           txtAccountFrom.text = ""
           isOverDratAllowed.Caption = ""
           lblOverDraftAmount.Caption = ""
           cmdAccountTo.SetFocus
        Else
        ''2-BANK_AC_NUM, 1-NominalCode, 2-CurrentBalance, 3-BANK_SC, 4-Bank_AC_Name, 5-AllowOverDraft, 6-OverDraftLimit
           txtNominalCode.text = flxClient.TextMatrix(flxClient.row, 1) 'NominalCode
           'txtBalance.text = IIf(IsNull(cboANF.Column(2)), "", cboANF.Column(2))
           txtSortCode.text = flxClient.TextMatrix(flxClient.row, 4)
           isOverDratAllowed.Caption = flxClient.TextMatrix(flxClient.row, 6)
           lblOverDraftAmount.Caption = flxClient.TextMatrix(flxClient.row, 7)
           'Resolved by BOSL
           '0000546: Bank Transfers not showing bank balance
           'Modified by anol 12 Mar 2015
           txtBalance.text = Format(BankAccBalance(adoConn, flxClient.TextMatrix(flxClient.row, 1), txtClientList.Tag), "0.00")   'For finding bank balance
           txtAccountName.text = IIf(IsNull(flxClient.TextMatrix(flxClient.row, 5)), "", flxClient.TextMatrix(flxClient.row, 5))
           FocusControl cmdAccountTo
        End If
        cmdAccountTo.SetFocus
    ElseIf sTextBox = "3" Then
           ' txtAccountTo.Tag = flxClient.TextMatrix(flxClient.row, 2)
            txtAccountTo.text = flxClient.TextMatrix(flxClient.row, 2)
'            If Trim(txtAccountFrom.text) = "" Then
'                MsgBox "No clients found with more than one bank account. Please go to Clients > Bank Details Tab to add more bank accounts.", vbInformation, "No Clients with Multiple Bank Accounts"
'                Exit Sub
'            End If
          If txtAccountFrom.text = txtAccountTo.text Then
               MsgBox " A Bank Transfer is not possible between same Bank Account. Please select another Bank Account", vbCritical + vbOKOnly, "Bank transfer"
               txtTNominalCode.text = ""
               txtTBalance.text = ""
               txtTSortCode.text = ""
               txtTAccountName.text = ""
               txtAccountTo.text = ""
               cmdAccountTo.SetFocus
            Else
               txtTNominalCode.text = flxClient.TextMatrix(flxClient.row, 1) 'NominalCode
               'txtTBalance.text = IIf(IsNull(cboANT.Column(2)), "", cboANT.Column(2))
               'Resolved by BOSL
               '0000546: Bank Transfers not showing bank balance
               'Modified by anol 12 Mar 2015
               txtTBalance.text = Format(BankAccBalance(adoConn, flxClient.TextMatrix(flxClient.row, 1), txtClientList.Tag), "0.00")  'For finding bank balance
               
               txtTSortCode.text = flxClient.TextMatrix(flxClient.row, 4)
               txtTAccountName.text = flxClient.TextMatrix(flxClient.row, 5)
               cmdFund.SetFocus
            End If
    ElseIf sTextBox = "4" Then
            txtFund.Tag = flxClient.TextMatrix(flxClient.row, 0)
            txtFund.text = flxClient.TextMatrix(flxClient.row, 2)
            txtBkTrRef.SetFocus
    End If
    adoConn.Close
    picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        picClient.Visible = False
         Frame5(1).Enabled = True
         Frame5(3).Enabled = True
         If sTextBox = "1" Then
            cmdClientList.SetFocus
         ElseIf sTextBox = "2" Then
            cmdAccountFrom.SetFocus
         ElseIf sTextBox = "3" Then
            cmdAccountTo.SetFocus
         ElseIf sTextBox = "4" Then
           cmdFund.SetFocus
         
         ElseIf sTextBox = "5" Then
           ' cmdNC.SetFocus
         ElseIf sTextBox = "6" Then
            'cmdFund.SetFocus
         ElseIf sTextBox = "7" Then
            'cmdVATCode.SetFocus
         End If
        
    End If
    If KeyAscii = 13 Then
         flxClient_Click
    End If
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   Me.Height = 6105
   Me.Width = 11415
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame5(3).BackColor = vbButtonFace 'Me.BackColor
   Frame5(1).BackColor = vbButtonFace ' Me.BackColor
'   connect to database
   adoConn.Open getConnectionString

   PrepareListBankTransf adoConn

   adoConn.Close
   Set adoConn = Nothing

   txtBkTrDate.text = Format(Date, "dd/mm/yyyy")
   lblPostingDate.ToolTipText = txtBkTrDate.text
   Call WheelHook(Me.hWnd)
End Sub

Private Sub PrepareListBankTransf(adoConn As ADODB.Connection)
   'If cboClientName.ListCount > 0 Then Exit Sub

   Dim adoRst As New ADODB.Recordset
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String, i As Integer, j As Integer
   Dim szSQL As String

   'On Error GoTo ErrorHandler
'   adoConn.Open getConnectionString

'*************************************** CLIENT BANK TRANSFER COMBO ******************************************
 'Resolved by BOSL
 'Issue No: 0000467
 'Load only those clients who have more than one bank accounts
 'Modified By: Asif. 04 Sep 2014
 
'   szSQL = "SELECT DISTINCT tlbClientBanks.CLIENT_ID, Client.ClientName " & _
'           "FROM tlbClientBanks, Client " & _
'           "WHERE tlbClientBanks.CLIENT_ID = Client.ClientID " & _
'           "ORDER BY CLIENT_ID;"

    szSQL = "SELECT  tlbClientBanks.CLIENT_ID, Client.ClientName From tlbClientBanks, Client " & _
            "Where tlbClientBanks.CLIENT_ID = Client.ClientID " & _
            "GROUP BY  tlbClientBanks.CLIENT_ID, Client.ClientName " & _
            "HAVING COUNT(Client.ClientName) > 1 ORDER BY tlbClientBanks.CLIENT_ID;"

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If adoRst.RecordCount > 0 Then
        txtClientList.Tag = adoRst.Fields.Item("CLIENT_ID").Value
        txtClientList.text = adoRst.Fields.Item("ClientName").Value
'       ReDim szaClientID(adoRst.RecordCount - 1) As String
'
'       i = 0
'       If adoRst.EOF Then GoTo NoRes
'       While Not adoRst.EOF
'          cboClientName.AddItem adoRst.Fields.Item("ClientName").Value
'          szaClientID(i) = adoRst.Fields.Item(0).Value
'          i = i + 1
'          adoRst.MoveNext
'       Wend
   Else
        MsgBox "No clients found with more than one bank account. Please go to Clients > Bank Details Tab to add more bank accounts.", vbInformation, "No Clients with Multiple Bank Accounts"
   End If
   adoRst.Close
'   i = 0
   
   'End of Modification
'*************************************** FUND ******************************************
 'Resolved by BOSL
 'Issue No: 0000467
 'Modified By: Asif. 04 Sep 2014
 
'    szSQL = "SELECT FundID, FundName " & _
'           "FROM Fund " & _
'           "ORDER BY FundID;"

'   szSQL = "SELECT FundID, FundCode " & _
'           "FROM Fund " & _
'           "ORDER BY FundID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount
'   TotalCol = adoRst.Fields.count
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'
'   cboFundBankTransf.Column() = Data()

'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'   adoConn.Close
'   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

'   adoRst.Close
'   Set adoRst = Nothing
''   adoConn.Close
'   Set adoConn = Nothing
End Sub
Private Sub LoadAccountGrid()
    flxClient.Clear
    flxClient.RowHeight(0) = 0
    flxClient.Cols = 8
    flxClient.ColWidth(0) = 100
    flxClient.ColWidth(1) = 1500
    flxClient.ColWidth(2) = 4500
    flxClient.ColWidth(3) = 0
    flxClient.ColWidth(4) = 0
    flxClient.ColWidth(5) = 0
    flxClient.ColWidth(6) = 0
    flxClient.ColWidth(7) = 0
    
'    flxClient.Height = 3345
'    flxClient.Width = 5175
    flxClient.Rows = 2
    flxClient.ColAlignment(0) = vbLeftJustify
    flxClient.ColAlignment(1) = vbLeftJustify
    flxClient.ColAlignment(2) = vbLeftJustify
    
    lblClientID.Caption = "Nominal Code"
    lblClientName.Caption = "Bank AC Number"
    lblClientID.Width = 1400
    lblClientID.Left = 50
    lblClientName.Width = 2600
    
    txtSearchClientName.Left = 1620
    txtSearchClientID.Width = 1530
    txtSearchClientName.Visible = True
    txtSearchClientName.text = ""
    txtSearchClientID.text = ""
    txtSearchClientID.Left = 45
    txtSearchClientName.Width = 3240
    'cmdPicCLose.Left = 5010
    lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   ''0-BANK_AC_NUM, 1-NominalCode, 2-CurrentBalance, 3-BANK_SC, 4-Bank_AC_Name, 5-AllowOverDraft, 6-OverDraftLimit
   
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    Dim szSQL As String
    Dim rRow As Integer
     szSQL = "SELECT DISTINCT CLIENT_ID, " & _
                 "BANK_AC_NUM, NominalCode, CurrentBalance, " & _
                 "Bank_AC_Name, BANK_SC, AllowOverDraft, OverDraftLimit " & _
                "FROM tlbClientBanks AS CB " & _
                "WHERE CLIENT_ID = '" & txtClientList.Tag & "' " & _
                "ORDER BY CLIENT_ID;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
         adoRst.Close
         adoConn.Close
        Exit Sub
   End If

   'cboANF.Clear

   If adoRst.RecordCount < 2 Then
      MsgBox "This Client has only one account. It's not possible make the transfer.", vbCritical + vbOKOnly, "Account Number Selection"
      cmdClientList.SetFocus
      adoRst.Close
      adoConn.Close
      Set adoRst = Nothing
      Set adoConn = Nothing
      Exit Sub
   Else
'      cboANF.Enabled = True
'      cboANT.Enabled = True
      lblHClientID.Caption = txtClientList.Tag
        'Modified by anol 18 Mar 2015
        'Added 2 dimension for keeping value of overdraft amount
'      ReDim szaData(6, adoRst.RecordCount - 1) As String
'      iRec = 0
      '0-BANK_AC_NUM, 1-NominalCode, 2-CurrentBalance, 3-BANK_SC, 4-Bank_AC_Name, 5-AllowOverDraft, 6-OverDraftLimit
      rRow = 1
     
      While Not adoRst.EOF
'         szaData(0, iRec) = adoRst.Fields.Item("BANK_AC_NUM").Value
'         szaData(1, iRec) = adoRst.Fields.Item("NominalCode").Value
'         szaData(2, iRec) = adoRst.Fields.Item("CurrentBalance").Value
'         szaData(3, iRec) = adoRst.Fields.Item("BANK_SC").Value
'         szaData(4, iRec) = adoRst.Fields.Item("Bank_AC_Name").Value
'         szaData(5, iRec) = IIf(IsNull(adoRst.Fields.Item("AllowOverDraft").Value), "", adoRst.Fields.Item("AllowOverDraft").Value) 'adoRst.Fields.Item("AllowOverDraft").Value
'         szaData(6, iRec) = IIf(IsNull(adoRst.Fields.Item("OverDraftLimit").Value), "0", adoRst.Fields.Item("OverDraftLimit").Value) 'adoRst.Fields.Item("OverDraftLimit").Value
'         iRec = iRec + 1
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("NominalCode").Value
               flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("BANK_AC_NUM").Value
               flxClient.TextMatrix(rRow, 3) = adoRst.Fields.Item("CurrentBalance").Value
               flxClient.TextMatrix(rRow, 4) = adoRst.Fields.Item("BANK_SC").Value
               flxClient.TextMatrix(rRow, 5) = adoRst.Fields.Item("Bank_AC_Name").Value
               flxClient.TextMatrix(rRow, 6) = IIf(IsNull(adoRst.Fields.Item("AllowOverDraft").Value), "", adoRst.Fields.Item("AllowOverDraft").Value) 'adoRst.Fields.Item("AllowOverDraft").Value
               flxClient.TextMatrix(rRow, 7) = IIf(IsNull(adoRst.Fields.Item("OverDraftLimit").Value), "0", adoRst.Fields.Item("OverDraftLimit").Value) 'adoRst.Fields.Item("OverDraftLimit").Value
               flxClient.RowHeight(rRow) = 280
               adoRst.MoveNext
               If Not adoRst.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
                
      Wend

'      cboANF.Column() = szaData()
'      cboANT.Column() = szaData()
   End If
       
    
        If IsLoadedAndVisible("frmAutoBankReconciliation") Then
           frmAutoBankReconciliation.LoadDataExternally adoConn
        End If
        adoRst.Close
        adoConn.Close
        Set adoRst = Nothing
        Set adoConn = Nothing
End Sub
'Private Sub cboClientName_Click()
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, szaData() As String, iRec As Integer
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   'On Error GoTo ErrorHandler
'   adoConn.Open getConnectionString
'Rem by anol
'''----------------------------------------------------------------Property
''   szSQL = "SELECT PropertyID, PropertyName " & _
''           "FROM Property " & _
''           "WHERE PropertyID <> '' AND ClientID = '" & txtClientList.tag & "' " & _
''           "ORDER BY PropertyID;"
'''   Debug.Print szSQL
''   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   If adoRst.EOF Then GoTo NoRes
''
''   TotalRow = adoRst.RecordCount
''   TotalCol = adoRst.Fields.count - 1
''
''   ReDim Data(TotalCol, TotalRow) As String
''
''   For i = 0 To TotalRow
''       For j = 0 To TotalCol
''           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
''       Next j
''       adoRst.MoveNext
''       If adoRst.EOF Then Exit For
''   Next i
''
''   cboPropF.Enabled = True
''   cboPropT.Enabled = True
''   cboPropF.Column() = Data()
''   cboPropT.Column() = Data()
''   cboPropF.SetFocus
''   adoRst.Close
'
''-----------------------------------------------------------------Account details
''   szSQL = "SELECT DISTINCT CLIENT_ID, " & _
''                 "BANK_AC_NUM, NominalCode, CurrentBalance, " & _
''                 "Bank_AC_Name, BANK_SC, AllowOverDraft, OverDraftLimit " & _
''           "FROM tlbClientBanks AS CB " & _
''           "WHERE CLIENT_ID = '" & txtClientList.Tag & "' " & _
''           "ORDER BY CLIENT_ID;"
'''Debug.Print szSQL
''   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''
''   If adoRst.EOF Then GoTo NoRes
''
''   cboANF.Clear
''
''   If adoRst.RecordCount < 2 Then
''      MsgBox "This Client has only one account. It's not possible make the transfer.", vbCritical + vbOKOnly, "Account Number Selection"
''      cmdClientList.SetFocus
''      adoRst.Close
''      adoConn.Close
''      Set adoRst = Nothing
''      Set adoConn = Nothing
''      Exit Sub
''   Else
''      cboANF.Enabled = True
''      cboANT.Enabled = True
''      lblHClientID.Caption = txtClientList.Tag
''        'Modified by anol 18 Mar 2015
''        'Added 2 dimension for keeping value of overdraft amount
''      ReDim szaData(6, adoRst.RecordCount - 1) As String
''      iRec = 0
''      '0-BANK_AC_NUM, 1-NominalCode, 2-CurrentBalance, 3-BANK_SC, 4-Bank_AC_Name, 5-AllowOverDraft, 6-OverDraftLimit
''      While Not adoRst.EOF
''         szaData(0, iRec) = adoRst.Fields.Item("BANK_AC_NUM").Value
''         szaData(1, iRec) = adoRst.Fields.Item("NominalCode").Value
''         szaData(2, iRec) = adoRst.Fields.Item("CurrentBalance").Value
''         szaData(3, iRec) = adoRst.Fields.Item("BANK_SC").Value
''         szaData(4, iRec) = adoRst.Fields.Item("Bank_AC_Name").Value
''         szaData(5, iRec) = IIf(IsNull(adoRst.Fields.Item("AllowOverDraft").Value), "", adoRst.Fields.Item("AllowOverDraft").Value) 'adoRst.Fields.Item("AllowOverDraft").Value
''         szaData(6, iRec) = IIf(IsNull(adoRst.Fields.Item("OverDraftLimit").Value), "0", adoRst.Fields.Item("OverDraftLimit").Value) 'adoRst.Fields.Item("OverDraftLimit").Value
''         iRec = iRec + 1
''         adoRst.MoveNext
''      Wend
''
''      cboANF.Column() = szaData()
''      cboANT.Column() = szaData()
''   End If
''
''   If IsLoadedAndVisible("frmAutoBankReconciliation") Then
''      frmAutoBankReconciliation.LoadDataExternally adoConn
''   End If
'
'NoRes:
'   adoRst.Close
'   adoConn.Close
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
''   adoRst.Close
''   adoConn.Close
''   Set adoRst = Nothing
''   Set adoConn = Nothing
'End Sub

Private Sub cmdBTSave_Click()
   If txtClientList.text = "" Then
      MsgBox "Please select the Client first.", vbInformation + vbOKOnly, "Transfer"
      cmdClientList.SetFocus
      Exit Sub
   End If
   If txtBkTrDate.text = "" Then
      MsgBox "Please enter the date.", vbInformation + vbOKOnly, "Transfer"
      txtBkTrDate.SetFocus
      Exit Sub
   End If
'   If cboPropF.text = "" Then
'      MsgBox "Please select the Property.", vbInformation + vbOKOnly, "Transfer"
'      cboPropF.SetFocus
'      Exit Sub
'   End If
'   If cboPropT.text = "" Then
'      MsgBox "Please select the Property.", vbInformation + vbOKOnly, "Transfer"
'      cboPropT.SetFocus
'      Exit Sub
'   End If
   If Trim(txtAccountFrom.text) = "" Then
      MsgBox "Please select the Account Number from first.", vbInformation + vbOKOnly, "Transfer"
      cmdAccountFrom.SetFocus
      Exit Sub
   End If
   If txtAccountTo.text = "" Then
      MsgBox "Please select the Account Number to first.", vbInformation + vbOKOnly, "Transfer"
      cmdAccountTo.SetFocus
      Exit Sub
   End If
   If txtAccountFrom.text = txtAccountTo.text Then
      MsgBox "A Bank Transfer is not possible between the same Bank Account. Please select another Bank Account.", vbCritical + vbOKOnly, "Account Number Selection"
      cmdAccountTo.SetFocus
      Exit Sub
   End If
   If txtFund.text = "" Then
      MsgBox "Please select the Fund first.", vbInformation + vbOKOnly, "Transfer"
      cmdFund.SetFocus
      Exit Sub
   End If
   If txtBkTrAmt.text = "" Or Val(txtBkTrAmt.text) = 0 Then
      MsgBox "Please insert the Payment Value first.", vbInformation + vbOKOnly, "Transfer"
      txtBkTrAmt.SetFocus
      Exit Sub
   End If
'Overdraft Checking By anol 18 Mar 2015
   If Val(txtBalance.text) - Val(txtBkTrAmt.text) < 0 Then 'Account balance-Current Amount
         If isOverDratAllowed.Caption = "True" Then
            If Val(lblOverDraftAmount.Caption) > 0 Then
               If (Val(txtBalance.text) - Val(txtBkTrAmt.text)) * (-1) > Val(lblOverDraftAmount.Caption) Then
                  If MsgBox("This Bank Account is over its overdraft limit. Do you wish to continue?", vbQuestion + vbYesNo, "Bank Overdrawn") = vbNo Then Exit Sub
               End If
            End If
         Else
            MsgBox "This Bank Account cannot go overdrawn", vbInformation + vbOKOnly, "Bank Overdraft"
            Exit Sub
         End If
   End If
'End of modification
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   'On Error GoTo ErrorHandler
   adoConn.Open getConnectionString
   
   szSQL = "SELECT * FROM tlbBankPayment"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

'Add New Records - BP
   adoRst.AddNew
   adoRst!My_ID = UniqueID()
   adoRst!TRAN_ID = SlNumber("BP", "tlbBankPayment", adoConn)
   adoRst!BANK_AC = txtNominalCode.text
   adoRst!TRAN_DATE = Format(IIf(txtBkTrDate.text = "", Now, txtBkTrDate.text), "DD MMMM YYYY")
   adoRst!TRANS = "BP"
   adoRst!Nominal_code = adoRst!BANK_AC
   adoRst!DEPT_ID = txtFund.Tag             'Fund
   adoRst!PROJ_REF = txtBkTrRef.text                          'Reference
   adoRst!description = txtBkTrDes.text
   adoRst!NET_AMOUNT = CCur(txtBkTrAmt.text)
   adoRst!TAX_CODE = "T9"
   adoRst!vat = 0
   adoRst!transactionType = 11     'Bank Receipt = sdoBR 12, Bank Payment = sdoBP 11
   adoRst!propertyID = "" 'cboPropF.Value
   adoRst!clientID = txtClientList.Tag
   adoRst!postingDate = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")

   'issue 523
   'added by anol 20 jan 2015
   UpdateBankAcBal_Minus adoConn, adoRst!NET_AMOUNT, adoRst!BANK_AC, txtClientList.Tag

   adoRst.Update

'Add New Records - BR
   adoRst.AddNew
   adoRst!My_ID = UniqueID()
   adoRst!TRAN_ID = SlNumber("BR", "tlbBankPayment", adoConn)
   adoRst!BANK_AC = txtTNominalCode.text
   adoRst!TRAN_DATE = Format(IIf(txtBkTrDate.text = "", Now, txtBkTrDate.text), "DD MMMM YYYY")
   adoRst!TRANS = "BR"
   adoRst!Nominal_code = adoRst!BANK_AC
   adoRst!DEPT_ID = txtFund.Tag                'Fund
   adoRst!PROJ_REF = txtBkTrRef.text                          'Reference
   adoRst!description = txtBkTrDes.text
   adoRst!NET_AMOUNT = CCur(txtBkTrAmt.text)
   adoRst!TAX_CODE = "T9"
   adoRst!vat = 0
   adoRst!transactionType = 12     'Bank Receipt = sdoBR 12, Bank Payment = sdoBP 11
   adoRst!propertyID = "" 'cboPropT.Value
   adoRst!clientID = txtClientList.Tag
   adoRst!postingDate = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")

  'issue 523
'added by anol 20 Jan 2015
   UpdateBankAcBal_Plus adoConn, adoRst!NET_AMOUNT, adoRst!BANK_AC, txtClientList.Tag

   adoRst.Update
   adoRst.Close

   ShowMsgInTaskBar "Bank Transfer has been saved successfully."

   txtClientList.text = ""
   lblHClientID.Caption = ""
   txtAccountFrom.text = ""
   txtSortCode.text = ""
   txtTAccountName.text = ""
   txtAccountFrom.text = ""
   txtTSortCode.text = ""
   txtTAccountName.text = ""
   txtTNominalCode.text = ""
   txtBkTrDate.text = ""
   txtBkTrRef.text = ""
   txtFund.text = ""
   txtBkTrDes.text = ""
   txtBkTrAmt.text = ""
  ' cmdCloseBk(1).SetFocus

   If FrmBankTransfer_CALLING_FROM = "frmBankTransactions" Then
      UpdateBankTransGrid adoConn
   End If

'--------------------------------------------------------------------------------------------
'  Export Transactions to Nominal Ledger (NLPosting table)
   Export_BPnBR_2_NL adoConn
      
'--------------------------------------------------------------------------------------

NoRes:
   
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
   Unload Me
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
  
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub UpdateBankTransGrid(adoConn As ADODB.Connection)
   With frmBankTransactions
      .flxBankPay.Clear
      .flxBankPay.Rows = 2

        Call .LoadFlxBankPay(adoConn, "")
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmBankTransactions.Enabled = True
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClientList.text = "" Then
          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
          Exit Sub
    End If
    If Trim(txtBkTrDate.text) = "" Then
        ShowMsgInTaskBar "Please enter a date.", "Y", "N"
        txtBkTrDate.SetFocus
         Exit Sub
    End If
    If frmMMain.IsRibbonVersion Then
        Dim adoConn As New ADODB.Connection
        Dim szSQL As String
        adoConn.Open getConnectionString
        If IsPeriodStatus(txtBkTrDate.text, txtClientList.Tag, adoConn) = 0 Then
           ShowMsgInTaskBar "The issue date cannot fall within a closed financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtBkTrDate.text, txtClientList.Tag, adoConn) = 9 Then
           ShowMsgInTaskBar "The issue date does not fall in any existing financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
    End If
    
   DispayCalendar Me, lblPostingDate.ToolTipText, txtBkTrDate.text, lblHClientID.Caption
   
End Sub

Private Sub txtBkTrAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        cmdBTSave.SetFocus
    End If
   DigitTextKeyPress txtBkTrAmt, KeyAscii, 2
End Sub

Private Sub txtBkTrAmt_LostFocus()
   If txtBkTrAmt.text = "" Then Exit Sub

   txtBkTrAmt.text = Format(txtBkTrAmt.text, "0.00")
End Sub

Private Sub txtBkTrDate_Change()
   TextBoxChangeDate txtBkTrDate
   lblPostingDate.ToolTipText = txtBkTrDate.text
End Sub

Private Sub txtBkTrDate_GotFocus()
   If Len(txtBkTrDate.text) < 10 Then txtBkTrDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtBkTrDate
End Sub

Private Sub txtBkTrDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdAccountFrom.SetFocus
   End If
   If KeyAscii = 27 And txtBkTrDate.text = "" Then txtBkTrDate.text = Format(Now, "dd/mm/yyyy")
   
End Sub

Private Sub txtBkTrDate_LostFocus()
   If txtBkTrDate.text <> "" Then
      If TextBoxFormatDate(txtBkTrDate) Then
         If lblPostingDate.ToolTipText = "" And lblPostingDate.ToolTipText <> txtBkTrDate.text Then
            lblPostingDate.ToolTipText = txtBkTrDate.text
         End If
      End If
   End If
End Sub

Private Sub cmdBTCancel_Click()
   txtClientList.text = ""
   lblHClientID.Caption = ""
   txtAccountFrom.text = ""
   txtSortCode.text = ""
   txtTAccountName.text = ""
   txtAccountFrom.text = ""
   txtTSortCode.text = ""
   txtTAccountName.text = ""
   txtTNominalCode.text = ""
   txtBkTrDate.text = ""
   txtBkTrRef.text = ""
   txtFund.text = ""
   txtBkTrDes.text = ""
   txtBkTrAmt.text = ""
'   cmdAccountFrom.Enabled = False
'   cmdAccountTo.Enabled = False
'   cboFundBankTransf.Enabled = False
End Sub

Private Sub cmdCloseBk_Click(Index As Integer)
   Unload Me
   Exit Sub
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        picClient.Visible = False
         Frame5(1).Enabled = True
         Frame5(3).Enabled = True
         If sTextBox = "1" Then
            cmdClientList.SetFocus
         ElseIf sTextBox = "2" Then
            cmdAccountFrom.SetFocus
         ElseIf sTextBox = "3" Then
            cmdAccountTo.SetFocus
         ElseIf sTextBox = "4" Then
           cmdFund.SetFocus
         
         ElseIf sTextBox = "5" Then
           ' cmdNC.SetFocus
         ElseIf sTextBox = "6" Then
            'cmdFund.SetFocus
         ElseIf sTextBox = "7" Then
            'cmdVATCode.SetFocus
         End If
        
    End If
End Sub
