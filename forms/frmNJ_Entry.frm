VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNJ_Entry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nominal Journal"
   ClientHeight    =   12195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15540
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNJ_Entry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12195
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picVatType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2430
      ScaleHeight     =   3180
      ScaleWidth      =   4260
      TabIndex        =   57
      Top             =   7785
      Visible         =   0   'False
      Width           =   4290
      Begin VB.CommandButton cmdBankClose 
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
         Left            =   3915
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   70
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxVatType 
         Height          =   2715
         Left            =   45
         TabIndex        =   59
         Top             =   405
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   4789
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
      Begin MSForms.Label Label8 
         Height          =   195
         Left            =   1680
         TabIndex        =   61
         Top             =   105
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Type"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label9 
         Height          =   195
         Left            =   180
         TabIndex        =   60
         Top             =   90
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "ID"
         Size            =   "2170;344"
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
         Height          =   285
         Index           =   0
         Left            =   90
         Top             =   90
         Width           =   3780
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   8280
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   31
      Top             =   8595
      Visible         =   0   'False
      Width           =   5295
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
         Left            =   5010
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   36
         Top             =   675
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5900
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
         TabIndex        =   39
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1650
         TabIndex        =   33
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   34
         Top             =   375
         Width           =   3420
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6032;450"
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
         Width           =   4950
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   45
      TabIndex        =   51
      Top             =   7155
      Width           =   15450
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "&Cancel"
         Height          =   355
         Left            =   9450
         TabIndex        =   56
         Top             =   270
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   355
         Left            =   10590
         TabIndex        =   19
         Top             =   255
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   225
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   13500
         TabIndex        =   20
         Top             =   285
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame fraHeader 
      Caption         =   "  Nominal Journal Header  "
      Height          =   1290
      Left            =   45
      TabIndex        =   21
      Top             =   0
      Width           =   15450
      Begin VB.CommandButton cmdPropID 
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
         Left            =   6345
         TabIndex        =   50
         Top             =   720
         Width           =   300
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
         Height          =   300
         Left            =   6345
         TabIndex        =   0
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8835
         TabIndex        =   1
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8835
         MaxLength       =   100
         TabIndex        =   2
         Top             =   735
         Width           =   5820
      End
      Begin VB.TextBox txtPropID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   735
         Width           =   5175
      End
      Begin VB.TextBox txtClient 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   375
         Width           =   5220
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   285
         Left            =   10290
         TabIndex        =   30
         Top             =   375
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
      Begin VB.Label lblNJ_Id 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10605
         TabIndex        =   29
         Top             =   375
         Visible         =   0   'False
         Width           =   4050
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   735
         Width           =   885
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   375
         Width           =   705
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   0
         Left            =   8280
         TabIndex        =   25
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   195
         Index           =   1
         Left            =   8280
         TabIndex        =   24
         Top             =   735
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   1275
      Left            =   40
      TabIndex        =   40
      Top             =   1230
      Width           =   15450
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   450
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdVatCode 
         Caption         =   ".."
         Enabled         =   0   'False
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
         Left            =   6525
         TabIndex        =   11
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtVatCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   765
         Width           =   1170
      End
      Begin VB.TextBox txtVatType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3555
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   765
         Width           =   1440
      End
      Begin VB.CommandButton cmdVatType 
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
         Left            =   4995
         TabIndex        =   9
         Top             =   765
         Width           =   300
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   355
         Left            =   12195
         TabIndex        =   15
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox txtNCCODE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   165
         Width           =   990
      End
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
         Left            =   12780
         TabIndex        =   5
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox txtFund 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10305
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   180
         Width           =   2475
      End
      Begin VB.TextBox txtNominal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   165
         Width           =   2430
      End
      Begin VB.CommandButton cmdNominal 
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
         Left            =   4635
         TabIndex        =   3
         Top             =   160
         Width           =   300
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6135
         MaxLength       =   200
         TabIndex        =   4
         Top             =   180
         Width           =   3435
      End
      Begin VB.TextBox txtDr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   770
         Width           =   1725
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   7740
         TabIndex        =   12
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox txtCr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   7
         Top             =   770
         Width           =   1605
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   765
         Width           =   1575
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   355
         Left            =   11100
         TabIndex        =   14
         Top             =   720
         Width           =   1020
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   355
         Left            =   13335
         TabIndex        =   16
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal A/C:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   49
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   7
         Left            =   5220
         TabIndex        =   48
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund :"
         Height          =   195
         Index           =   8
         Left            =   9810
         TabIndex        =   47
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Code:"
         Height          =   195
         Index           =   10
         Left            =   5865
         TabIndex        =   46
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Amount:"
         Height          =   195
         Index           =   11
         Left            =   7995
         TabIndex        =   45
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit (+)"
         Height          =   195
         Index           =   12
         Left            =   1185
         TabIndex        =   44
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit (-)"
         Height          =   195
         Index           =   13
         Left            =   2835
         TabIndex        =   43
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
         Height          =   195
         Index           =   14
         Left            =   9675
         TabIndex        =   42
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Type:"
         Height          =   195
         Index           =   2
         Left            =   4185
         TabIndex        =   41
         Top             =   540
         Width           =   705
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNJ_SplitDelete 
      Height          =   900
      Left            =   45
      TabIndex        =   55
      Top             =   8370
      Visible         =   0   'False
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   1588
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483640
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame3 
      Height          =   4740
      Left            =   45
      TabIndex        =   63
      Top             =   2430
      Width           =   15450
      Begin VB.TextBox txtDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13995
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "0.00"
         Top             =   4245
         Width           =   1230
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "0.00"
         Top             =   4230
         Width           =   1095
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   4230
         Width           =   1140
      End
      Begin VB.TextBox txtNominalDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   4230
         Width           =   4785
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNJ_Split 
         Height          =   3510
         Left            =   45
         TabIndex        =   64
         Top             =   585
         Width           =   15360
         _ExtentX        =   27093
         _ExtentY        =   6191
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   3
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Difference:"
         Height          =   195
         Index           =   16
         Left            =   13125
         TabIndex        =   81
         Top             =   4245
         Width           =   780
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   15
         Left            =   8910
         TabIndex        =   80
         Top             =   4275
         Width           =   390
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name:"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   79
         Top             =   4275
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   45
         Top             =   225
         Width           =   15315
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   195
         Index           =   9
         Left            =   13875
         TabIndex        =   73
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   7
         Left            =   11475
         TabIndex        =   72
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Type"
         Height          =   195
         Index           =   4
         Left            =   7455
         TabIndex        =   71
         Top             =   285
         Width           =   675
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   3
         Left            =   5040
         TabIndex        =   70
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblGridCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   1245
         TabIndex        =   69
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal A/C"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   68
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lblGridCaption 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Amount"
         Height          =   195
         Index           =   8
         Left            =   12345
         TabIndex        =   67
         Top             =   285
         Width           =   885
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Code "
         Height          =   195
         Index           =   5
         Left            =   8925
         TabIndex        =   66
         Top             =   285
         Width           =   855
      End
      Begin VB.Label lblGridCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   6
         Left            =   10305
         TabIndex        =   65
         Top             =   285
         Width           =   705
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   74
         Top             =   225
         Width           =   15315
      End
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Loaded for Edit?"
      Height          =   195
      Index           =   9
      Left            =   13590
      TabIndex        =   28
      Top             =   135
      Width           =   1485
   End
End
Attribute VB_Name = "frmNJ_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vatOptionEnabled As Boolean
Public bFunction  As Boolean   'True -> Add New, False -> Edit
Public lHeaderID As Long
Public bEdit As String
Private sNew      As Byte      '1 -> new, 2 -> edit, 3 -> delete
Private szInVat   As String    'Debit vat Code
Private szInVatN  As String    'Debit Vat code name
Private szOutVat  As String
Private szOutVatN As String
Dim sTextBox As String
Dim ispropertyexstAsked As Boolean
Public bEditMode As Boolean


Private Sub cmbClient_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub cmbFund_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

'Private Sub cmbFund_LostFocus()
'         'If cmbNC.ListIndex < 0 Then txtNominal.text = ""
'    'issue 571 Validation
'   'Added by anol 20 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim adoConn    As New ADODB.Connection
'   If txtFund.Text <> "" Then
'            adoConn.Open getConnectionString
'            szSQL = "SELECT FundID, FundName FROM Fund where FundCode='" & txtFund.Text & "';"
'            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If adoRst.EOF Then
'                'It shall give messege if fund code is wrong
'                MsgBox "Please select a valid fund Code to proceed", vbInformation, "select a fund name"
'                cmbFund.SetFocus
'            End If
'            adoConn.Close
'            Set adoConn = Nothing
'    End If
'End Sub



Private Sub cmbVatType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub cmbVatType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub cmdBankClose_Click()
    picVatType.Visible = False
    fraHeader.Enabled = True
    Frame1.Enabled = True
    Frame3.Enabled = True
    Frame2.Enabled = True
    FocusControl cmdVatType
End Sub

Private Sub cmdCancel2_Click()
    cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
    flxNJ_Split_DblClick
   
End Sub

Private Sub cmdFund_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    LoadFund adoconn
    picClient.Left = 7920
    picClient.Top = 1215
    picClient.Visible = True
    sTextBox = "4"
    
    fraHeader.Enabled = False
    Frame1.Enabled = False
    fraHeader.Enabled = False
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdNominal_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call LoadCmbNC(adoconn)     'This is the main loading functions
    picClient.Left = 915
    picClient.Top = 1170
    picClient.Visible = True
    sTextBox = "3"
    
    fraHeader.Enabled = False
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame2.Enabled = False
    fraHeader.Enabled = False
    Frame1.Enabled = False
   
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPropID_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

'Private Sub cmbProperty_LostFocus()
'    'issue 571 Validation
'   'Added by anol 20 May 2015
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim adoConn    As New ADODB.Connection
'   If txtPropID.text <> "" Then
'            adoConn.Open getConnectionString
'            szSQL = "SELECT PropertyID, PropertyNAME " & _
'                    "FROM Property " & _
'                    "where PropertyNAME='" & txtPropID.text & "';"
'            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'            If adoRst.EOF Then
'                'It shall give messege if property is wrong
'                'and cleared down the property ID
'                txtPropID.text = ""
'                MsgBox "Please select a valid Property", vbInformation, "Select a Property"
'                txtPropID.SetFocus
'            End If
'            adoConn.Close
'            Set adoConn = Nothing
'    Else
'        txtPropID.text = ""
'    End If
'End Sub

Private Sub cmbNC_LostFocus()
    'If cmbNC.ListIndex < 0 Then txtNominal.text = ""
    'issue 571 Validation
   'Added by anol 20 May 2015
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim adoconn    As New ADODB.Connection
   If txtNominal.text <> "" Then
            adoconn.Open getConnectionString
            szSQL = "SELECT Code, Name " & _
           "FROM   NominalLedger " & _
           "WHERE ClientID = '" & txtClient.Tag & "' AND Code='" & txtNominal.text & "';"
            adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            If adoRst.EOF Then
                'It shall give messege if Nominal code is wrong
                MsgBox "Please select a valid nominal code to proceed", vbInformation, "select a nominal code"
                cmdNominal.SetFocus
            End If
            adoconn.Close
            Set adoconn = Nothing
    End If
End Sub

Private Sub cmbVatType_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub cmdCancel_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub cmdClient_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 270
    picClient.Visible = True
    LoadflxClient
    fraHeader.Enabled = False
    Frame1.Enabled = False
   fraHeader.Enabled = False
Frame1.Enabled = False
Frame3.Enabled = False
Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClose_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub cmdDelete_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub cmdPropID_Click()
    picClient.Left = 915
    picClient.Top = 470
    picClient.Visible = True
    sTextBox = "2"
    LoadPropertyList
    fraHeader.Enabled = False
    Frame1.Enabled = False
    fraHeader.Enabled = False
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdSave_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Function LoadVatOption(Conn As ADODB.Connection) As Integer
    Dim rsGlobalData As New ADODB.Recordset
    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtPropID.Tag & "'", Conn, adOpenStatic, adLockReadOnly
    If Not rsGlobalData.EOF Then
            LoadVatOption = IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value)
    End If
    rsGlobalData.Close
    Set rsGlobalData = Nothing
End Function

Private Sub Command1_Click()

End Sub

Private Sub cmdVatCode_Click()
    sTextBox = "VatCode"
    picVatType.Top = 1710
    picVatType.Left = 2990
    picVatType.Visible = True
    configflxVatCode
    LoadVatCode
    FocusControl flxVatType
    flxVatType.row = 1
    fraHeader.Enabled = False
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame2.Enabled = False
End Sub
Private Sub LoadVatCode()
   ' Error Handler
   On Error GoTo Error_Handler
   Dim adoconn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoconn = New ADODB.Connection
   adoconn.Open getConnectionString

   szSQL = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM tlbVATCODE where IN_USE ORDER BY VAT_ID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Please setup vat Code."
      picVatType.Visible = False
   Else
      flxVatType.Rows = adoRst.RecordCount + 1
      rRow = 1
      While Not adoRst.EOF
         flxVatType.TextMatrix(rRow, 1) = adoRst.Fields.Item("VAT_CODE").Value
         flxVatType.TextMatrix(rRow, 2) = adoRst.Fields.Item("VAT_RATE").Value
         flxVatType.TextMatrix(rRow, 3) = adoRst.Fields.Item("VAT_ID").Value
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
       picVatType.Visible = True
       flxVatType.row = 1
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoconn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "Prestige Database Error: ", vbExclamation, "Vat code"

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoconn = Nothing
End Sub
Private Sub cmdVatType_Click()
    sTextBox = "VatType"
    picVatType.Top = 1710
    picVatType.Left = 990
    picVatType.Visible = True
    configflxVatType
    LoadCmbVatType
    FocusControl flxVatType
    flxVatType.row = 1
    fraHeader.Enabled = False
Frame1.Enabled = False
Frame3.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub flxClient_Click()
    
    Dim adoconn As New ADODB.Connection
    Dim rstVat As New ADODB.Recordset
    Dim rstRec As New ADODB.Recordset
    fraHeader.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    fraHeader.Enabled = True
    Frame1.Enabled = True
    Frame3.Enabled = True
    Frame2.Enabled = True
    Dim szSQL As String
    adoconn.Open getConnectionString
    If sTextBox = "1" Then 'when you click a client from the grid
    '******
        If Label50(9).Caption = "Loaded" And bEdit = "1" And txtClient.text <> flxClient.TextMatrix(flxClient.row, 1) Then
            flxNJ_Split.Tag = "1"
            cmdSave.Enabled = True
        End If

        txtClient.Tag = flxClient.TextMatrix(flxClient.row, 1)
        txtClient.text = flxClient.TextMatrix(flxClient.row, 2)
        txtVAT.Enabled = True
        
        If txtClient.text = "" Then Exit Sub
        txtFund.text = "" ' You need to clear fund when you change a client (fund assignment by client)
        txtFund.Tag = ""
        
        'adoConn.Open getConnectionString
        LoadCmbProperty adoconn
        
'        LoadCmbFund adoConn
        LoadIO_VAT adoconn
        adoconn.Close
        Set adoconn = Nothing
       
        cmdPropID.SetFocus
        Dim strTemp As String
        txtClient.ForeColor = vbBlack
        strTemp = isControlAccountSet(txtClient.Tag)
        If Len(strTemp) > 0 Then
            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
            strTemp = ""
            picClient.Visible = False
            txtClient.ForeColor = vbRed
            Exit Sub
        End If
        
        '''''
        
    ElseIf sTextBox = "2" Then
        txtPropID.Tag = flxClient.TextMatrix(flxClient.row, 1)
        txtPropID.text = flxClient.TextMatrix(flxClient.row, 2)
         
               
        
        txtDateFrom.SetFocus
    ElseIf sTextBox = "3" Then
        txtNominal.Tag = flxClient.TextMatrix(flxClient.row, 1)
         txtNCCODE.text = flxClient.TextMatrix(flxClient.row, 1)
        txtNominal.text = flxClient.TextMatrix(flxClient.row, 2)
        txtDescription.SetFocus
    ElseIf sTextBox = "4" Then
        txtFund.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
        txtFund.text = flxClient.TextMatrix(flxClient.row, 1)
        txtDr.SetFocus
    End If
'    adoConn.Close
'    Set adoConn = Nothing
    picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            flxClient_Click
        End If
'    Dim adoConn As New ADODB.Connection
'    fraHeader.Enabled = True
'    Frame1.Enabled = True
'    Frame2.Enabled = True
'
'     adoConn.Open getConnectionString
'    If sTextBox = "1" Then
'    '******
'        If Label50(9).Caption = "Loaded" And bEdit = "1" And txtClient.text <> flxClient.TextMatrix(flxClient.row, 1) Then
'            flxNJ_Split.Tag = "1"
'            cmdSave.Enabled = True
'        End If
''        If frmMMain.IsRibbonVersion And IsDate(txtDateFrom.text) = True Then
''
''            Dim szSQL As String
''            'adoConn.Open getConnectionString
''            If IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoConn) = 0 Then
''               ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
''               adoConn.Close
''               Set adoConn = Nothing
''               Exit Sub
''            ElseIf IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoConn) = 9 Then
''               ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
''               adoConn.Close
''               Set adoConn = Nothing
''               Exit Sub
''            End If
''        End If
'        '*******
'        txtClient.Tag = flxClient.TextMatrix(flxClient.row, 0)
'        txtClient.text = flxClient.TextMatrix(flxClient.row, 1)
'
'        If txtClient.text = "" Then Exit Sub
'
'       ' adoConn.Open getConnectionString
'        LoadCmbProperty adoConn
'        LoadCmbNC adoConn
'        LoadCmbFund adoConn
'        LoadIO_VAT adoConn
'        adoConn.Close
'        Set adoConn = Nothing
'
'        cmdPropID.SetFocus
'
'
'        '''''
'
'    ElseIf sTextBox = "2" Then
'        txtPropID.text = flxClient.TextMatrix(flxClient.row, 0)
'        txtPropID.Tag = flxClient.TextMatrix(flxClient.row, 1)
'
'        txtDateFrom.SetFocus
'    End If
''    adoConn.Close
''    Set adoConn = Nothing
'    picClient.Visible = False
End Sub

Private Sub flxNJ_Split_Click()
    txtNominalDesc.Visible = True
    Label50(17).Visible = True
    txtNominalDesc.text = flxNJ_Split.TextMatrix(flxNJ_Split.row, 17)
End Sub

Private Sub flxVatType_Click()
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    Dim rstVat As New ADODB.Recordset
    Dim rstRec As New ADODB.Recordset
    fraHeader.Enabled = True
    Frame1.Enabled = True
    Frame3.Enabled = True
    Frame2.Enabled = True
    adoconn.Open getConnectionString
    If sTextBox = "VatType" Then
        txtVatType.Tag = flxVatType.TextMatrix(flxVatType.row, 1)
        txtVatType.text = flxVatType.TextMatrix(flxVatType.row, 2)
        picVatType.Visible = False
        If txtPropID.text = "" Then
                        'When you do not select a property, take vat code value from the  client
                         'if you open the client form and chage the value you need to instant change the effect. so execute an sql to get the new value.
                         szSQL = "SELECT CLIENTID, CLIENTNAME, CT, V.VAT_CODE, V.VAT_ID, V.VAT_RATE FROM ((CLIENT C INNER JOIN Supplier S ON C.ClientID=S.SupplierID) " & _
                        "LEFT JOIN tlbVatCode V on S.VATCode=cstr(V.vat_ID)) where CLIENTID='" & txtClient.Tag & "' ORDER BY CLIENTID;"
                         rstVat.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                         If Not rstVat.EOF Then
                                txtVatCode.text = IIf(IsNull(rstVat.Fields("VAT_CODE").Value), "", rstVat.Fields("VAT_CODE").Value)
                                If txtVatCode.text <> "" Then
                                    txtRate.text = IIf(IsNull(rstVat.Fields("VAT_RATE").Value), "0.00", rstVat.Fields("VAT_RATE").Value)
                                    txtRate.Tag = txtVatCode.text
                                    txtVatCode.Tag = IIf(IsNull(rstVat.Fields("VAT_ID").Value), "-1", rstVat.Fields("VAT_ID").Value)
                                    txtVatCode.text = txtVatCode.text & " / " & txtRate.text
                                End If
                                 
                         End If
                         cmdVatCode.Enabled = True
                         txtVAT.Enabled = True
                         rstVat.Close
                         Set rstVat = Nothing
         Else
                        vatOptionEnabled = LoadVatOption(adoconn)
                        szSQL = "SELECT P.PropertyID, P.PropertyName,G.VATRate,V.VAT_Rate as RateValue,V.VAT_CODE as VAT_CODE1,G.VATRate  as  Rate,G.vatOptionEnabled " & _
                                         " from ((Property P INNER JOIN globalData G ON P.PropertyID=G.PropertyID) LEFT JOIN tlbVatCode V ON G.VATRate=V.VAT_ID) " & _
                                         "WHERE P.PropertyID = '" & txtPropID.Tag & "' " & _
                                         "ORDER BY P.PropertyID;"
                        rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

                        If vatOptionEnabled = True Then
                            txtVatCode.text = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value)
                            If txtVatCode.text <> "" Then
                                     txtRate.Tag = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value) 'T5
                                     txtRate.text = IIf(IsNull(rstRec.Fields("RateValue").Value), "", rstRec.Fields("RateValue").Value)
                                     txtVatCode.text = IIf(IsNull(rstRec.Fields("VAT_CODE1").Value), "", rstRec.Fields("VAT_CODE1").Value) & " / " & IIf(IsNull(rstRec.Fields("RateValue").Value), "-1", rstRec.Fields("RateValue").Value)
                                     txtVatCode.Tag = IIf(IsNull(rstRec.Fields("VATRate").Value), "", rstRec.Fields("VATRate").Value) ''VAT_ID
                            Else
                                    txtVatCode.text = ""
                                    txtVatCode.Tag = ""
                            End If
                            cmdVatCode.Enabled = True
                            txtVAT.Enabled = True

                        Else
                            txtVatCode.text = ""
                            txtVatCode.Tag = ""
                            txtRate.text = ""
                            txtRate.Tag = ""
                            cmdVatCode.Enabled = False
                            txtVAT.Enabled = False
                        End If
                        
                        rstRec.Close
                        Set rstRec = Nothing
         End If
'         FocusControl cmdVatCode
         If txtVatType.Tag = 0 Or txtVatType.Tag = 3 Or txtVatType.Tag = 4 Then
            cmdVatCode.Enabled = False
            txtRate.text = "0"
            txtRate.Tag = ""
            txtVatCode.Tag = ""
            txtVatCode.text = ""
            FocusControl cmdOk
        Else
            cmdVatCode.Enabled = True
            FocusControl cmdVatCode
        End If
        If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
               txtVAT.text = Format(Val(txtDr.text) * (Val(txtRate.text) / 100), "0.00")
               txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
        End If
        If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
               txtVAT.text = Format(Val(txtCr.text) * (Val(txtRate.text) / 100), "0.00")
               txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
        End If
    End If
    If sTextBox = "VatCode" Then
        txtVatCode.Tag = flxVatType.TextMatrix(flxVatType.row, 3)
        txtVatCode.text = flxVatType.TextMatrix(flxVatType.row, 1) & " / " & flxVatType.TextMatrix(flxVatType.row, 2) 'Vat code name / Rate
        txtRate.text = flxVatType.TextMatrix(flxVatType.row, 2)
        txtRate.Tag = flxVatType.TextMatrix(flxVatType.row, 1)
        If flxVatType.TextMatrix(flxVatType.row, 1) = "" Then
                txtVatCode.text = ""
                txtVatCode.Tag = ""
                txtRate.Tag = ""
                txtRate.text = ""
        End If
        If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
               txtVAT.text = Format(Val(txtDr.text) * (Val(txtRate.text) / 100), "0.00")
               txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
        End If
        If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
               txtVAT.text = Format(Val(txtCr.text) * (Val(txtRate.text) / 100), "0.00")
               txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
        End If
        picVatType.Visible = False
        txtVAT.Enabled = True
        FocusControl cmdOk
    End If
        
      
End Sub

Private Sub flxVatType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxVatType_Click
    End If
End Sub

Private Sub txtClient_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdClient.SetFocus
    End If
End Sub

Private Sub txtCredit_Change()
     txtCredit.text = Format(Val(txtCredit.text), "0.00")
End Sub

Private Sub txtCredit_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtDebit_Change()
    txtDebit.text = Format(Val(txtDebit.text), "0.00")
End Sub

Private Sub txtDebit_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtDescription_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFund.SetFocus
    End If
End Sub

Private Sub txtDiff_Change()
    If Val(txtDebit.text) = 0 And Val(txtCredit.text) = 0 Then
        cmdSave.Enabled = True
    Else
        If Val(txtDiff.text) = 0 And Val(txtDebit.text) > 0 Then
            cmdSave.Enabled = True
        ElseIf Val(txtDiff.text) = 0 And Val(txtCredit.text) > 0 Then
            cmdSave.Enabled = True
        Else
            cmdSave.Enabled = False
        End If
    End If
End Sub

Private Sub txtDiff_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtNC_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtPropID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         cmdPropID.SetFocus
        End If
End Sub

'Private Sub cmdPropID_GotFocus()
'    txtNominalDesc.Visible = False
'    Label50(17).Visible = False
'End Sub

Private Sub txtTitle_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdNominal.SetFocus
    End If
End Sub

Private Sub txtTotal_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
End Sub

'Private adoCA     As ADODB.Recordset
'
'Private Sub chkVatOnly_Click()
'   If chkVatOnly Then
'      optVatOnlyCr.Enabled = True
'      optVatOnlyDr.Enabled = True
'      optVatOnlyDr.Value = True
'      cmbNC.Enabled = False
'      txtDr.Enabled = False
'      txtDr.text = ""
'      txtCr.Enabled = False
'      txtCr.text = ""
'      txtTotal.text = Format(txtVAT.text, "0.00")
'   Else
'      optVatOnlyCr.Enabled = False
'      optVatOnlyDr.Enabled = False
'      optVatOnlyCr.Value = False
'      optVatOnlyDr.Value = False
'      cmbNC.Enabled = True
'      txtDr.Enabled = True
'      txtCr.Enabled = True
'   End If
'End Sub

Private Sub txtVAT_Change()
    If IsNumeric(txtVAT.text) = False Then
        txtVAT.text = "0.00"
    End If
End Sub
Private Sub cmbClient_Click()
  
End Sub

Private Sub LoadIO_VAT(adoconn As ADODB.Connection)

'Resolved by BOSL
'Issue No: 0000482
'Load the Input VAT Account and Output VAT Account from the system.
'Modified By: Asif. 01 Oct 2014
         
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT   S.Code AS NCode, S.CAType AS Type, S.Name AS NName " & _
'           "FROM     NominalLedger AS S " & _
'           "WHERE   (S.CAType = 'I' OR S.CAType = 'O') AND " & _
'                 "NOT S.CAPosting AND " & _
'                 "S.CAFixed AND S.ClientID = '" & txtClient.tag & "'; "
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
'      If adoRst.Fields.Item("Type").Value = "I" Then
'         If (IsNull(adoRst.Fields.Item("Type").Value) Or IsNull(adoRst.Fields.Item("NCode").Value)) Then
'            szInVat = "NONE"
''            szInVatN = ""
'         Else
'            szInVat = adoRst.Fields.Item("NCode").Value
''            szInVatN = adoRst.Fields.Item("NName").Value
'         End If
'      End If
'      If adoRst.Fields.Item("Type").Value = "O" Then
'         If (IsNull(adoRst.Fields.Item("Type").Value) Or IsNull(adoRst.Fields.Item("NCode").Value)) Then
'            szOutVat = "NONE"
''            szOutVatN = ""
'         Else
'            szOutVat = adoRst.Fields.Item("NCode").Value
''            szOutVatN = adoRst.Fields.Item("NName").Value
'         End If
'      End If
'      adoRst.MoveNext
'   Wend
'
'   adoRst.Close
'   Set adoRst = Nothing

    szInVat = GetNominalCodeForControlAccount(adoconn, "Input VAT", txtClient.Tag())
    szInVatN = GetNominalNameForControlAccount(adoconn, "Input VAT", txtClient.Tag())
    szOutVat = GetNominalCodeForControlAccount(adoconn, "Output VAT", txtClient.Tag())
    szOutVatN = GetNominalNameForControlAccount(adoconn, "Output VAT", txtClient.Tag())

End Sub

Private Sub LoadCmbFund(adoconn As ADODB.Connection)
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundCode as FundName FROM Fund;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
   Else
      
         txtFund.Tag = adoRst.Fields.Item("FundID").Value
         txtFund.text = adoRst.Fields.Item("FundName").Value
         
   End If
   adoRst.Close
   Set adoRst = Nothing
   
End Sub
Private Sub configflxVatType()
   flxVatType.Visible = True
   flxVatType.Clear
   flxVatType.Cols = 4
   flxVatType.TextMatrix(0, 0) = ""
   flxVatType.TextMatrix(0, 1) = ""
   flxVatType.ColWidth(0) = 60
   flxVatType.ColWidth(1) = 1400
   flxVatType.ColAlignment(1) = vbLeftJustify
   flxVatType.ColWidth(2) = 2400
   flxVatType.ColAlignment(2) = vbLeftJustify
   flxVatType.ColWidth(3) = 0
   flxVatType.RowHeight(0) = 0
   flxVatType.Rows = 2
   Label9.Caption = "ID"
   Label8.Caption = "Type"
End Sub
Private Sub configflxVatCode()
   flxVatType.Visible = True
   flxVatType.Clear
   flxVatType.Cols = 4
   flxVatType.TextMatrix(0, 0) = ""
   flxVatType.TextMatrix(0, 1) = ""
   flxVatType.ColWidth(0) = 60
   flxVatType.ColWidth(1) = 1200
   flxVatType.ColAlignment(1) = vbLeftJustify
   flxVatType.ColWidth(2) = 2600
   flxVatType.ColWidth(3) = 0
   flxVatType.RowHeight(0) = 0
   flxVatType.Rows = 2
   Label9.Caption = "Vat Code"
   Label8.Caption = "Vat Rate"
End Sub
Private Sub LoadCmbVatType()
'   Dim Data(1, 4)  As String
'   Dim i       As Integer
   Dim rRow As Integer
'   Data(0, 0) = 0
'   Data(1, 0) = "N/A"
'   Data(0, 1) = 1
'   Data(1, 1) = "Input"
'   Data(0, 2) = 2
'   Data(1, 2) = "Output"
'   Data(0, 3) = 3
'   Data(1, 3) = "VAT Only-Input"
'   Data(0, 4) = 4
'   Data(1, 4) = "VAT Only-Output"
'   cmbVatType.Column() = Data()
   
           rRow = 1
           flxVatType.TextMatrix(rRow, 0) = ""
           flxVatType.TextMatrix(rRow, 1) = 0
           flxVatType.TextMatrix(rRow, 2) = "N/A"
           flxVatType.RowHeight(rRow) = 280
           
           
           flxVatType.AddItem ""
           rRow = rRow + 1
           flxVatType.TextMatrix(rRow, 0) = ""
           flxVatType.TextMatrix(rRow, 1) = 1
           flxVatType.TextMatrix(rRow, 2) = "Input"
           flxVatType.RowHeight(rRow) = 280
           
           flxVatType.AddItem ""
           rRow = rRow + 1
           flxVatType.TextMatrix(rRow, 0) = ""
           flxVatType.TextMatrix(rRow, 1) = 2
           flxVatType.TextMatrix(rRow, 2) = "Output"
           flxVatType.RowHeight(rRow) = 280
           
            flxVatType.AddItem ""
           rRow = rRow + 1
           flxVatType.TextMatrix(rRow, 0) = ""
           flxVatType.TextMatrix(rRow, 1) = 3
           flxVatType.TextMatrix(rRow, 2) = "VAT Only-Input"
           flxVatType.RowHeight(rRow) = 280
           
            flxVatType.AddItem ""
           rRow = rRow + 1
           flxVatType.TextMatrix(rRow, 0) = ""
           flxVatType.TextMatrix(rRow, 1) = 4
           flxVatType.TextMatrix(rRow, 2) = "VAT Only-Output"
           flxVatType.RowHeight(rRow) = 280
        
        
   End Sub

Private Sub LoadCmbNC(adoconn As ADODB.Connection)
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer

'   szSQL = "SELECT Code, Name " & _
'           "FROM   NominalLedger " & _
'           "WHERE ClientID = '" & txtClient.text & "' AND Posting AND " & _
'                 "Code NOT IN " & _
'                     "(SELECT NCode FROM SpareTable1 " & _
'                     " WHERE ClientID = '" & txtClient.text & "' AND " & _
'                        "(NCode <> '' OR NOT ISNULL(NCode))) " & _
'           "ORDER BY Code;"
   
'   szSQL = "SELECT Code, Name " & _
'           "FROM   NominalLedger " & _
'           "WHERE ClientID = '" & txtClient.text & "' AND Posting AND " & _
'                 "(ISNULL(CAType) OR CAType = '') " & _
'           "ORDER BY Code;"
'Modified by anol 27 July 2015

'Load normal nominal codes  CAFixed=0 means (in option control account) set allow posting= Yes then show here as NC
Dim rstRec As New ADODB.Recordset
 szSQL = "SELECT N.* " & _
      "FROM NominalLedger AS N " & _
      "WHERE N.ClientID = '" & txtClient.Tag & "' AND " & _
      "Posting AND CAFixed=0 AND CODE NOT IN " & _
      "(SELECT NominalCode FROM tlbClientBanks where Client_ID = '" & txtClient.Tag & "')" & _
      " ORDER BY N.Code;"
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly


 Dim rRow As Integer
   'Dim szSQL As String

  ' Dim adoConn As New ADODB.Connection
   
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 3600
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Nominal Code"
   lblClientName.Caption = "Nominal Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345
  
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
   Set rstRec = Nothing

End Sub
Private Sub LoadFund(adoconn As ADODB.Connection)
    Dim adoRst     As New ADODB.Recordset
    Dim szSQL      As String
    Dim Data()     As String
    Dim TotalRow   As Integer
    Dim TotalCol   As Integer
    Dim i          As Integer
    Dim j          As Integer

    'Modified by anol 27 July 2015
    Dim rstRec As New ADODB.Recordset
    Dim rsFundMatrix As New ADODB.Recordset
    
    rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoconn, adOpenStatic, adLockReadOnly
    If rsFundMatrix("isfundAssign").Value = False Then
         szSQL = "SELECT FundID, FundCode, FundName FROM Fund Order by FundCode;"
    Else
         szSQL = "Select * from fundMatrix where PropertyID='" & txtPropID.Tag & "' and ClientID='" & txtClient.Tag & "' and isDeleted=false"
    End If
    rsFundMatrix.Close
 
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly


    Dim rRow As Integer
   
   
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 3600
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345
  
           rRow = 1
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = " " & rstRec("FundID").Value 'FundID
           flxClient.TextMatrix(rRow, 1) = rstRec("FundCode").Value 'FundCode
           flxClient.TextMatrix(rRow, 2) = rstRec("FundName").Value  'FundName
           flxClient.RowHeight(rRow) = 280 'FundID, FundCode, FundName
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
 
   rstRec.Close
   Set rstRec = Nothing

End Sub
Private Function GetNextNJ_ID() As Long
   Dim adoconn    As New ADODB.Connection
   Dim adoRst     As New ADODB.Recordset
   Dim szSQL      As String

   adoconn.Open getConnectionString

   adoRst.Open "SELECT max(RecordID)+1 FROM NJ_Header;", adoconn, adOpenStatic, adLockReadOnly
   GetNextNJ_ID = IIf(IsNull(adoRst.Fields.Item(0).Value), 1, adoRst.Fields.Item(0).Value)
   adoRst.Close
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Function

Private Sub LoadCmbProperty(adoconn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClient.Tag & "' " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
        txtPropID.Tag = adoRst.Fields("PropertyID").Value
        txtPropID.text = adoRst.Fields("PropertyName").Value
  Else
        txtPropID.Tag = ""
        txtPropID.text = ""
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub



Private Sub cmbNC_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
   If txtClient.Tag = "" Then
      ShowMsgInTaskBar "Please select the client first", "Y", "N"
      cmdClient.SetFocus
      Exit Sub
   End If
End Sub


Private Sub cmbVatRate_Click()
      If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
         txtVAT.text = Format(Val(txtDr.text) * (Val(txtRate.text) / 100), "0.00")
         txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
      End If
      If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
         txtVAT.text = Format(Val(txtCr.text) * (Val(txtRate.text) / 100), "0.00")
         txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
      End If
   
End Sub

Private Sub cmbVatType_Click()
   If txtVatType.Tag = 0 Then
      cmdVatCode.Enabled = False
      txtVatCode.text = ""
      txtVAT.Enabled = False
      txtVAT.text = ""
      If txtDr.text <> "" Then
         txtTotal.text = Format(txtDr.text, "0.00")
      End If
      If txtCr.text <> "" Then
         txtTotal.text = Format(txtCr.text, "0.00")
      End If
   Else
     
      cmdVatCode.Enabled = True
      txtVAT.Enabled = True

      If txtVatType.Tag = 3 Or txtVatType.Tag = 4 Then
         txtTotal.text = Format(txtVAT.text, "0.00")
         txtDr.text = ""
         txtCr.text = ""
         txtDr.SetFocus
         txtVAT.Enabled = False
         txtVAT.text = ""
         cmdVatCode.Enabled = False
      Else
         txtDr.Enabled = True
         txtCr.Enabled = True
         txtVAT.Enabled = True
         
      End If
   End If
End Sub

Private Sub cmbVatType_LostFocus()
   If txtVatType.text = "" Then
      ShowMsgInTaskBar "VAT type has not been selected", "Y", "N"
        FocusControl cmdVatType
      Exit Sub
   End If
End Sub

Private Sub cmdCancel_Click()
   fraHeader.Enabled = True
   cmdClear_Click
   flxNJ_Split.Enabled = True
   cmdEdit.Enabled = True
   'Flag for NLPosting need to clear here
End Sub

Private Sub cmdClear_Click()
'   Dim adoconn As New ADODB.Connection
'   adoconn.Open getConnectionString
   txtVAT.text = ""
   
   txtDescription.text = ""
   txtNominal.text = ""
   txtNCCODE.text = ""
   txtFund.text = ""
   txtFund.Tag = ""
   txtDr.text = ""
   txtCr.text = ""
   txtVatCode.text = ""
   txtVatCode.Tag = ""
   txtRate.text = ""
   txtRate.Tag = ""
   txtVAT.text = ""
   txtVatType.text = ""
   txtVAT.text = ""
   txtTotal.text = ""
   txtDr.Enabled = True
   txtCr.Enabled = True
'   adoconn.Close
   cmdEdit.Enabled = True
   FocusControl cmdNominal
End Sub

Private Sub cmdClose_Click()
   frmNJ.Enabled = True
   Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim inRow As Integer
    Dim iCurEditRow As Integer
    Dim iCol As Integer
    Dim iRow As Integer
    inRow = flxNJ_Split.row
    If flxNJ_Split.row = 0 Then Exit Sub
    flxNJ_Split.Tag = "1"
    If Val(flxNJ_Split.TextMatrix(inRow, 7)) > 0 Then
        txtDebit.text = txtDebit.text - Val(flxNJ_Split.TextMatrix(inRow, 7)) + Val(flxNJ_Split.TextMatrix(inRow, 9))
        txtDiff.text = Val(txtDebit.text) - Val(txtCredit.text)
    End If
    If Val(flxNJ_Split.TextMatrix(inRow, 8)) > 0 Then
        txtCredit.text = txtCredit.text - flxNJ_Split.TextMatrix(inRow, 8) + Val(flxNJ_Split.TextMatrix(inRow, 10))
        txtDiff.text = Val(txtDebit.text) - Val(txtCredit.text)
    End If
    
    'flxNJ_Split.RemoveItem flxNJ_Split.row
    'copy deleted row to the other delete grid
    iCurEditRow = flxNJ_Split.row
    iRow = iCurEditRow
    For iCol = 1 To flxNJ_Split.Cols - 1
       flxNJ_SplitDelete.TextMatrix(flxNJ_SplitDelete.Rows - 1, iCol) = flxNJ_Split.TextMatrix(iRow, iCol)
    Next iCol
    flxNJ_SplitDelete.AddItem ""
     
     'step up from the Nominal grid when delete
     
    For iRow = iCurEditRow To flxNJ_Split.Rows - 2
         For iCol = 1 To flxNJ_Split.Cols - 1
            flxNJ_Split.TextMatrix(iRow, iCol) = flxNJ_Split.TextMatrix(iRow + 1, iCol)
         Next iCol
     Next iRow
     
     'when delete a row just empty the grid when you have 2 or less than 2 Row in the main grid
    If flxNJ_Split.Rows <= 2 Then
            For iCol = 0 To flxNJ_Split.Cols - 1
               flxNJ_Split.TextMatrix(flxNJ_Split.row, iCol) = ""
            Next iCol
    Else
        flxNJ_Split.RemoveItem flxNJ_Split.Rows - 1
    End If
     
     
    sNew = 3
    
    cmdClear_Click
End Sub

Private Sub cmdOK_Click()
   If Val(txtVatType.Tag) > 0 And Val(txtVAT.text) > 0 Then
      If szInVat = "NONE" Or szOutVat = "NONE" Then
         MsgBox "Please setup the input and output VAT control account " & Chr(13) & "in the option module.", vbOKOnly, "Warning"
         Exit Sub
      End If
   End If
   If IsNull(txtClient.Tag) Or txtClient.Tag = "" Then
      MsgBox "Please select the client", vbOKOnly, "Warning"
      FocusControl cmdClient
      Exit Sub
   End If
   
   If txtVatType.text = "Input" Or txtVatType.text = "Output" Then 'you can disable vat from global data so this validation not needed
        If txtVatCode.text = "" Then
            MsgBox "Please select a VAT Code from the list.", vbOKOnly, "Warning"
            FocusControl cmdVatCode
            Exit Sub
        End If
   End If

   If txtVatType.text = "" Then
        MsgBox "Please select a VAT Type", vbOKOnly, "Warning"
        FocusControl cmdVatType
        Exit Sub
   End If
   If txtDateFrom.text = "" Then
      MsgBox "Please enter the date", vbOKOnly, "Warning"
      FocusControl txtDateFrom
      Exit Sub
   End If
   If txtTitle.text = "" Then
      MsgBox "Please enter a title", vbOKOnly, "Warning"
      FocusControl txtTitle
      Exit Sub
   End If
   If txtNominal.text = "" And Val(txtVatType.Tag) < 3 Then
      MsgBox "Please select a nominal account", vbOKOnly, "Warning"
      FocusControl cmdNominal
      Exit Sub
   End If
   'issue 571 Validation
   'Modified by anol
   If txtFund.text = "" Then
      MsgBox "Please select a fund", vbOKOnly, "Warning"
      FocusControl cmdFund
      Exit Sub
   End If
   If Trim(txtPropID.text) = "" And ispropertyexstAsked = False Then
            ispropertyexstAsked = True
            If MsgBox("You have not selected a property. Do you wish to add a property?", vbYesNo, "Select a Property") = vbYes Then
                cmdPropID.SetFocus
                Exit Sub
            End If
   End If
'   If Val(txtTotal.text) <= 0 Then
'      ShowMsgInTaskBar "Please enter a value", "Y", "N"
'      FocusControl txtDr
'      Exit Sub
'   End If
  
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   If frmMMain.IsRibbonVersion And IsDate(txtDateFrom.text) = True Then
        If IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 0 Then
           MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Please correct Posting date"
           txtDateFrom.SetFocus
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 9 Then
           MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Please correct Posting date"
           txtDateFrom.SetFocus
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
   End If
   fraHeader.Enabled = False
   cmdSave.Enabled = True
   Call CreateRowsFlxNJ_Split 'We are doing main adding precess to grid here
   'added by anol 31 Aug 2015
   Dim iRow As Integer
   txtDebit.text = "0.00"
   txtCredit.text = "0.00"
   For iRow = 1 To flxNJ_Split.Rows - 1
        txtDebit.text = txtDebit.text + Val(flxNJ_Split.TextMatrix(iRow, 7)) + IIf(Val(flxNJ_Split.TextMatrix(iRow, 7)) > 0, Val(flxNJ_Split.TextMatrix(iRow, 10)), 0)
        txtCredit.text = txtCredit.text + Val(flxNJ_Split.TextMatrix(iRow, 8)) + IIf(Val(flxNJ_Split.TextMatrix(iRow, 8)) > 0, Val(flxNJ_Split.TextMatrix(iRow, 10)), 0)
   Next
   txtDiff.text = Format(Val(txtDebit.text) - Val(txtCredit.text), "0.00")
   txtDebit.text = Format(Val(txtDebit.text), "0.00")
   txtCredit.text = Format(Val(txtCredit.text), "0.00")
   'End of addition
   If sNew = 2 Then
      flxNJ_Split.Enabled = True
      cmdClear.Enabled = True
      sNew = 1
   End If

   cmdClear_Click
   'MsgBox "ok"
   txtVAT.text = ""
   cmdNominal.SetFocus

End Sub

Private Function returnTypeID(Name As String) As Integer
    If Name = "N/A" Then
        returnTypeID = 0
    ElseIf Name = "Input" Then
        returnTypeID = 1
    ElseIf Name = "Output" Then
        returnTypeID = 2
    ElseIf Name = "VAT Only-Input" Then
        returnTypeID = 3
    ElseIf Name = "VAT Only-Output" Then
        returnTypeID = 4
    End If

End Function
Private Function returnType(id As Integer) As String
    If id = 0 Then
        returnType = "N/A"
    ElseIf id = 1 Then
        returnType = "Input"
    ElseIf id = 2 Then
        returnType = "Output"
    ElseIf id = 3 Then
        returnType = "VAT Only-Input"
    ElseIf id = 4 Then
        returnType = "VAT Only-Output"
    End If

End Function
Private Sub CreateRowsFlxNJ_Split() 'When we click ok we are executing this method , this add new line to a grid
   'anol 2020-10-22 Index has been updated for this sub procedure +2
   Dim inRow As Integer
   flxNJ_Split.Tag = 1
   If sNew = 1 Or sNew = 3 Then                                                  'Add new record
      If flxNJ_Split.TextMatrix(flxNJ_Split.Rows - 1, 0) <> "" Then
         flxNJ_Split.AddItem ""
      End If
      inRow = flxNJ_Split.Rows - 1
   End If
   If sNew = 2 Then                                                  'Edit Grid
      inRow = flxNJ_Split.row
   End If

'   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund|<Type|>Dr|>Cr|>VATRate" & _
'                    0         1            2          3    4    5   6      7
'               "|>VATAmt|>TotalAmt|NC|Fund|VatType|VRate|VatOnly"
'                    8         9    10  11    12      13    14

   If sNew = 1 Or sNew = 3 Then
        flxNJ_Split.TextMatrix(inRow, 0) = UniqueID()
   End If
   flxNJ_Split.TextMatrix(inRow, 2) = IIf(IsNull(txtDescription.text), "", txtDescription.text)
   flxNJ_Split.TextMatrix(inRow, 3) = IIf(IsNull(txtFund.Tag) Or txtFund.Tag = "", "", txtFund.text)
   flxNJ_Split.TextMatrix(inRow, 15) = IIf(IsNull(txtRate.Tag), "", txtRate.Tag) 'VAT CODE NAME'txtNominal.text
   flxNJ_Split.TextMatrix(inRow, 5) = returnType(CInt(txtVatType.Tag))
   flxNJ_Split.TextMatrix(inRow, 6) = txtRate.Tag
   flxNJ_Split.TextMatrix(inRow, 17) = txtNominal.text
   If txtVatType.Tag < 3 Then       'If it is not vat only transaction
            flxNJ_Split.TextMatrix(inRow, 1) = txtNominal.Tag
            flxNJ_Split.TextMatrix(inRow, 4) = IIf(Val(txtDr.text) > 0, "Journal Debit", "Journal Credit")
           
            If Val(txtDr.text) > 0 Then
               flxNJ_Split.TextMatrix(inRow, 7) = Format(Val(txtDr.text), "0.00")
               txtDiff.text = Format(Val(txtDiff.text) + Val(txtTotal.text), "0.00")
               flxNJ_Split.TextMatrix(inRow, 8) = ""
               flxNJ_Split.TextMatrix(inRow, 11) = txtTotal.text                              'Total
            Else
               flxNJ_Split.TextMatrix(inRow, 8) = Format(Val(txtCr.text), "0.00")
               txtDiff.text = Format(Val(txtDiff.text) - Val(txtTotal.text), "0.00")
               flxNJ_Split.TextMatrix(inRow, 7) = ""
               flxNJ_Split.TextMatrix(inRow, 11) = Format((-1) * Val(txtTotal.text), "0.00")   'Total
            End If
            flxNJ_Split.TextMatrix(inRow, 9) = txtRate.text   'we are not using this when we save this split lines
            flxNJ_Split.TextMatrix(inRow, 12) = txtNominal.Tag
            flxNJ_Split.TextMatrix(inRow, 15) = IIf(IsNull(txtRate.Tag), "", txtRate.Tag) 'VAT CODE NAME
            flxNJ_Split.TextMatrix(inRow, 16) = "N"
      
   Else
          If Val(txtDr.text) > 0 Then flxNJ_Split.TextMatrix(inRow, 4) = "Journal Debit"
          If Val(txtCr.text) > 0 Then flxNJ_Split.TextMatrix(inRow, 4) = "Journal Credit"
    
          'Resolved by BOSL
          'Issue No: 0000482
          'Assign the Input VAT Account or Output VAT Account based on the VAT Type selected
          'Modified By: Asif. 01 Oct 2014
          
          If txtVatType.Tag = 3 Then
             flxNJ_Split.TextMatrix(inRow, 1) = szInVat
             flxNJ_Split.TextMatrix(inRow, 12) = szInVat
             flxNJ_Split.TextMatrix(inRow, 17) = szInVatN
          ElseIf txtVatType.Tag = 4 Then
             flxNJ_Split.TextMatrix(inRow, 1) = szOutVat
             flxNJ_Split.TextMatrix(inRow, 12) = szOutVat
             flxNJ_Split.TextMatrix(inRow, 17) = szOutVatN
          End If
    
          If InStr(flxNJ_Split.TextMatrix(inRow, 4), "Debit") > 0 Then
             flxNJ_Split.TextMatrix(inRow, 7) = txtDr.text
             txtDiff.text = Format(Val(txtDiff.text) + Val(txtTotal.text), "0.00")
             flxNJ_Split.TextMatrix(inRow, 11) = txtTotal.text
          Else
             flxNJ_Split.TextMatrix(inRow, 8) = txtCr.text
             txtDiff.text = Format(Val(txtDiff.text) - Val(txtTotal.text), "0.00")
             flxNJ_Split.TextMatrix(inRow, 11) = Format((-1) * Val(txtTotal.text), "0.00")     'Total
          End If
          flxNJ_Split.TextMatrix(inRow, 16) = IIf(InStr(flxNJ_Split.TextMatrix(inRow, 4), "Debit") > 0, "DV", "CV")   'DV--> Debit VAT, CV--> Credit VAT
   End If

   flxNJ_Split.TextMatrix(inRow, 10) = txtVAT.text
   flxNJ_Split.TextMatrix(inRow, 13) = txtFund.Tag
   flxNJ_Split.TextMatrix(inRow, 14) = txtVatType.Tag
 
 
  
   cmdSave.Enabled = Val(txtDiff.text) = 0
End Sub
'
'Private Sub UpdateRowFlxNJ_Split_()
'   Dim inRow As Integer
'
'   inRow = flxNJ_Split.row
'
''   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund|<Type|>Dr|>Cr|>VATRate" & _
''                    0           1           2         3     4    5   6     7
''               "|>VATAmt|>TotalAmt|NC|Fund|VatType|VRate|VatOnly"
''                    8        9     10  11    12      13    14
'   flxNJ_Split.TextMatrix(inRow, 2) = IIf(IsNull(txtDescription.text), "", txtDescription.text)
'   flxNJ_Split.TextMatrix(inRow, 3) = IIf(IsNull(txtFund.Tag) Or txtFund.Tag = "", "", txtFund.Text)
'   If txtVatType.Tag < 3 Then
'      flxNJ_Split.TextMatrix(inRow, 1) = txtNominal.text
'      flxNJ_Split.TextMatrix(inRow, 4) = IIf(Val(txtDr.text) > 0, "Journal Debit", "Journal Credit")
'   Else
'      If txtVatType.Tag = 1 Then flxNJ_Split.TextMatrix(inRow, 4) = "Journal Debit"
'      If txtVatType.Tag = 2 Then flxNJ_Split.TextMatrix(inRow, 4) = "Journal Credit"
'   End If
'   If txtVatType.Tag < 3 Then
'      If Val(txtDr.text) > 0 Then
'         flxNJ_Split.TextMatrix(inRow, 5) = txtDr.text
'         txtDiff.text = Format(Val(txtDiff.text) + Val(txtDr.text), "0.00")
'      Else
'         flxNJ_Split.TextMatrix(inRow, 5) = txtCr.text
'         txtDiff.text = Format(Val(txtDiff.text) - Val(txtCr.text), "0.00")
'      End If
'   Else
'      flxNJ_Split.TextMatrix(inRow, 5) = txtVAT.text
'      If txtVatType.Tag = 1 Then
'         txtDiff.text = Format(Val(txtDiff.text) + Val(txtVAT.text), "0.00")
'      Else
'         txtDiff.text = Format(Val(txtDiff.text) - Val(txtVAT.text), "0.00")
'      End If
'   End If
'   flxNJ_Split.TextMatrix(inRow, 6) = txtRate.Text
'   flxNJ_Split.TextMatrix(inRow, 7) = txtVAT.text
'   flxNJ_Split.TextMatrix(inRow, 8) = txtTotal.text         'Total
'   flxNJ_Split.TextMatrix(inRow, 9) = txtNominal.tag
'   flxNJ_Split.TextMatrix(inRow, 10) = txtFund.Tag
'   flxNJ_Split.TextMatrix(inRow, 11) = txtVatType.Tag
'   flxNJ_Split.TextMatrix(inRow, 12) = cmbVatRate.Value
'
'   cmdSave.Enabled = Val(txtDiff.text) = 0
'End Sub
'
'Private Sub CreateVATLine(iType As Integer)
'   Dim inRow As Integer
'
'   flxNJ_Split.AddItem ""
'   inRow = flxNJ_Split.Rows - 1
'
'   flxNJ_Split.TextMatrix(inRow, 0) = UniqueID()
'   flxNJ_Split.TextMatrix(inRow, 1) = szi
'   flxNJ_Split.TextMatrix(inRow, 2) = IIf(IsNull(txtDescription.text), "", txtDescription.text)
'   flxNJ_Split.TextMatrix(inRow, 3) = IIf(IsNull(txtFund.Tag) Or txtFund.Tag = "", "", txtFund.Text)
'
'   If iType = 1 Then       'DR
'      flxNJ_Split.TextMatrix(inRow, 4) = "Journal Debit"
'   Else                    'Cr
'
'   End If
'   flxNJ_Split.TextMatrix(inRow, 5) = Format(txtVAT.text, "0.00")
'End Sub

Private Sub cmdSave_Click()
    'this sub procedure is uptodate with +2 grid row position
   On Error GoTo ERR_HANDLER
   
   If Not cmdOk.Enabled Then
      cmdClient.SetFocus
      Exit Sub
   End If
   If txtClient.ForeColor = vbRed Then
         MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & _
         vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
         Exit Sub
   End If
   If flxNJ_Split.Rows = 2 Or flxNJ_Split.Rows = 1 Then
         MsgBox "No data has been entered for saving.... ", vbCritical + vbOKOnly, "Warning!"
         Exit Sub
   End If
   If txtClient.Tag = "" Then
        MsgBox "Please select a valid client.", vbInformation, "Select a valid client."
        cmdClient.SetFocus
        Exit Sub
   End If
'   If txtPropID.Tag = "" Then
'        MsgBox "Please select a valid Property.", vbInformation, "Select a valid Property."
'        cmdPropID.SetFocus
'        Exit Sub
'   End If
   
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer
   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString
    If frmMMain.IsRibbonVersion And IsDate(txtDateFrom.text) = True Then
        If IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 0 Then
           MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Please correct Posting date"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 9 Then
           MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Please correct Posting date"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
        If DateDiff("d", lblPostingDate.ToolTipText, txtDateFrom.text) > 0 Then
            MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
            Exit Sub
        End If

    End If
    cmdSave.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    FocusControl cmdClose
    ispropertyexstAsked = False

'  Save the header first
   adoconn.BeginTrans
   If bFunction Then                               'ADD NEW
      szSQL = "SELECT * FROM NJ_Header;"
      adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

      lHeaderID = GetNextNJ_ID()
      adoRst.AddNew
      adoRst.Fields.Item("RecordID").Value = lHeaderID
      adoRst.Fields.Item("CreatedBy").Value = User
      adoRst.Fields.Item("CreatedDate").Value = Now
   Else
      szSQL = "SELECT * FROM NJ_Header WHERE RecordID = " & lHeaderID & ";"
      adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
   End If

   adoRst.Fields.Item("ClientID").Value = txtClient.Tag
   adoRst.Fields.Item("PropertyID").Value = IIf(IsNull(txtPropID.Tag), "", txtPropID.Tag)
   adoRst.Fields.Item("NJDate").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
   adoRst.Fields.Item("NJTitle").Value = txtTitle.text
   adoRst.Fields.Item("PostingDate").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
         adoRst.Fields.Item("LastModifiedBy").Value = User
      adoRst.Fields.Item("LastModifiedDate").Value = Now
   adoRst.Update

   adoRst.Close
   Set adoRst = Nothing

'  Save the splits
   If Not bFunction Then                               'EDIT
      adoconn.Execute "UPDATE NLPOSTING SET DeleteFlag=1 where  TRANS_ID ='" & lHeaderID & "' AND PARENT_RECORD IN (SELECT RecordID from  NJ_Split where ParentID= " & lHeaderID & " )"
      adoconn.Execute "DELETE * FROM NJ_Split WHERE ParentID = " & lHeaderID & ";"
   End If

   adoRst.Open "SELECT * FROM NJ_Split;", adoconn, adOpenDynamic, adLockOptimistic
'Exit Sub
   
   For iRow = 1 To flxNJ_Split.Rows - 1
      If flxNJ_Split.TextMatrix(iRow, 13) <> "" Then
      adoRst.AddNew
      adoRst.Fields.Item("ParentID").Value = lHeaderID
      adoRst.Fields.Item("RecordID").Value = flxNJ_Split.TextMatrix(iRow, 0)
      adoRst.Fields.Item("NC").Value = flxNJ_Split.TextMatrix(iRow, 12)
      adoRst.Fields.Item("SpLineDes").Value = flxNJ_Split.TextMatrix(iRow, 2)
      adoRst.Fields.Item("FundID").Value = Val(flxNJ_Split.TextMatrix(iRow, 13))
      adoRst.Fields.Item("TYPE_ID").Value = flxNJ_Split.TextMatrix(iRow, 14)
      adoRst.Fields.Item("NetAmt").Value = IIf(flxNJ_Split.TextMatrix(iRow, 7) = "", _
                     IIf(flxNJ_Split.TextMatrix(iRow, 8) = "", 0, flxNJ_Split.TextMatrix(iRow, 8)), _
                     flxNJ_Split.TextMatrix(iRow, 7))
      adoRst.Fields.Item("VAT_CODE").Value = flxNJ_Split.TextMatrix(iRow, 15)
      adoRst.Fields.Item("VATAmt").Value = IIf(flxNJ_Split.TextMatrix(iRow, 10) = "", 0, flxNJ_Split.TextMatrix(iRow, 10))
      adoRst.Fields.Item("TotalAmt").Value = IIf(Val(flxNJ_Split.TextMatrix(iRow, 11)) < 0, _
                                                 Val(flxNJ_Split.TextMatrix(iRow, 11)) * (-1), _
                                                 Val(flxNJ_Split.TextMatrix(iRow, 11)))

      
      adoRst.Update
      End If
   Next iRow
   adoRst.Close

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'##########################         NOMINAL LEDGER POSTING : NLPosting        ########################################################################
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'   szHeader$ = "HeaderID(0)|<NominalCode(1)|<Description(2)|<Fund(3)|<Type(4)" & _
'               "|>NetDrAmount(5)|>NetCrAmount(6)|<VATRate(7)|>VATAmt(8)" & _
'               "|>TotalAmt(9)|NC(10)|Fund(11)|Type(12)|VRate(13)"
'  Update Nominal Posting Table
   bFunction = True
   If bFunction Then             'Create New NJ
                  With adoRst
                     szSQL = "SELECT * FROM NLPosting;"
                     .Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
            
                     For iRow = 1 To flxNJ_Split.Rows - 1
                        If flxNJ_Split.TextMatrix(iRow, 16) = "N" Then
                           .AddNew
                           .Fields.Item("THIS_RECORD").Value = UniqueID()
                           .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                           .Fields.Item("TRANS_ID").Value = lHeaderID
                           .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                           .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                           .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
                           If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                              .Fields.Item("TRANSACTION_TYPE").Value = 15
                              .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 7))
                           Else
                              .Fields.Item("TRANSACTION_TYPE").Value = 16
                              .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 8)) * -1
                           End If
                           'Updated upto this print anol 2020-10-23
                           .Fields.Item("ACCOUNT_NUMBER").Value = flxNJ_Split.TextMatrix(iRow, 1)
                           .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
                           .Fields.Item("UNIT_ID").Value = ""
                           .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
                           .Fields.Item("REFERENCE").Value = txtTitle.text
                           .Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 1)
                           .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
                           .Fields.Item("AMOUNT_TYPE").Value = "A"
                           .Fields.Item("USER_NUMBER").Value = "N"   'I dont know what is this field has been used for. just for the time being, i put N; means the transaction has been posted from NJ
                           .Fields.Item("ClientID").Value = txtClient.Tag
                           .Fields.Item("Deleteflag").Value = False
                           .Update
            
                           If Val(flxNJ_Split.TextMatrix(iRow, 10)) > 0 Then
                              .AddNew
                              .Fields.Item("THIS_RECORD").Value = UniqueID()
                              .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                              .Fields.Item("TRANS_ID").Value = lHeaderID
                              .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                              .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                              .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
                              
                              'Resolved by BOSL
                              'Issue No: 0000482
                              'Assign the Input VAT Account or Output VAT Account based on the VAT Type selected
                              'Modified By: Asif. 01 Oct 2014
            
                              If flxNJ_Split.TextMatrix(iRow, 14) = "1" Or flxNJ_Split.TextMatrix(iRow, 14) = "3" Then
                                 .Fields.Item("NOMINAL_CODE").Value = szInVat
                                 
                              ElseIf flxNJ_Split.TextMatrix(iRow, 14) = "2" Or flxNJ_Split.TextMatrix(iRow, 14) = "4" Then
                                 .Fields.Item("NOMINAL_CODE").Value = szOutVat
                                 
                              End If
                              
                              If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                                 .Fields.Item("TRANSACTION_TYPE").Value = 15   'input vat
            '                     .Fields.Item("NOMINAL_CODE").Value = szInVat
                                 .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10))
                              Else
                                 .Fields.Item("TRANSACTION_TYPE").Value = 16
            '                     .Fields.Item("NOMINAL_CODE").Value = szOutVat
                                 .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10)) * -1
                              End If
                              
                              '''''''''''''''''''''
                              
                              .Fields.Item("ACCOUNT_NUMBER").Value = .Fields.Item("NOMINAL_CODE").Value 'flxNJ_Split.TextMatrix(iRow, 1) updated 2020-10-27 by anol
                              .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
                              .Fields.Item("UNIT_ID").Value = ""
                              .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
                              
                              .Fields.Item("REFERENCE").Value = txtTitle.text
                              .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
                              .Fields.Item("AMOUNT_TYPE").Value = "V"
                              .Fields.Item("USER_NUMBER").Value = "N"
                              .Fields.Item("ClientID").Value = txtClient.Tag
                               .Fields.Item("Deleteflag").Value = False
                              .Update
                           End If
                        Else
                           .AddNew
                           .Fields.Item("THIS_RECORD").Value = UniqueID()
                           .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                           .Fields.Item("TRANS_ID").Value = lHeaderID
                           .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                           .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                           .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
                           If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                              .Fields.Item("TRANSACTION_TYPE").Value = 15
                              .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 11))
                           Else
                              .Fields.Item("TRANSACTION_TYPE").Value = 16
                              .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 11)) * -1
                           End If
                           .Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 1)
                           
                              'Resolved by BOSL
                              'Issue No: 0000482
                              'Assign the Input VAT Account or Output VAT Account based on the VAT Type selected
                              'Modified By: Asif. 01 Oct 2014
            
                            If flxNJ_Split.TextMatrix(iRow, 14) = "1" Or flxNJ_Split.TextMatrix(iRow, 14) = "3" Then
                              .Fields.Item("NOMINAL_CODE").Value = szInVat
                              
                            ElseIf flxNJ_Split.TextMatrix(iRow, 14) = "2" Or flxNJ_Split.TextMatrix(iRow, 14) = "4" Then
                              .Fields.Item("NOMINAL_CODE").Value = szOutVat
                              
                            End If
            '               .Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 10)
            
            ''''''''''''''''''''''''''''
                           
                           .Fields.Item("ACCOUNT_NUMBER").Value = flxNJ_Split.TextMatrix(iRow, 1)
                           .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
                           .Fields.Item("UNIT_ID").Value = ""
                           .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
                           '.Fields.Item("Amount").Value = flxNJ_Split.TextMatrix(iRow, 8)
                           
                           .Fields.Item("REFERENCE").Value = txtTitle.text
                           .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
                           .Fields.Item("AMOUNT_TYPE").Value = "V"
                           .Fields.Item("USER_NUMBER").Value = "N"
                           .Fields.Item("ClientID").Value = txtClient.Tag
                           .Update
                        End If
                     Next iRow
                     .Close
                  End With
   Else         '  Edit New NJ
      With adoRst
         For iRow = 1 To flxNJ_Split.Rows - 1
            If flxNJ_Split.TextMatrix(iRow, 16) = "N" Then 'I dont know what does this N means
               szSQL = "SELECT * FROM NLPosting " & _
                       "WHERE PARENT_RECORD = '" & flxNJ_Split.TextMatrix(iRow, 0) & "' AND " & _
                             "AMOUNT_TYPE = 'A' AND DeleteFlag=false;"
               .Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

               If .EOF Then
                  .AddNew
                  .Fields.Item("THIS_RECORD").Value = UniqueID()
                  .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                  .Fields.Item("TRANS_ID").Value = lHeaderID
                  .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                  .Fields.Item("AMOUNT_TYPE").Value = "A"
                  .Fields.Item("USER_NUMBER").Value = "N"   'I dont know what is this field has been used for. just for the time being, i put N; means the transaction has been posted from NJ
               End If
               .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
               
               If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                  .Fields.Item("TRANSACTION_TYPE").Value = 15
                  .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 7))
               Else
                  .Fields.Item("TRANSACTION_TYPE").Value = 16
                  .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 8)) * -1
               End If
               .Fields.Item("ACCOUNT_NUMBER").Value = flxNJ_Split.TextMatrix(iRow, 1)
               .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
               .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
               .Fields.Item("REFERENCE").Value = txtTitle.text
               .Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 1)
               .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
               .Fields.Item("ClientID").Value = txtClient.Tag
               .Update
               .Close

               If Val(flxNJ_Split.TextMatrix(iRow, 10)) > 0 Then
                  szSQL = "SELECT * FROM NLPosting " & _
                          "WHERE PARENT_RECORD = '" & flxNJ_Split.TextMatrix(iRow, 0) & "' AND " & _
                                "AMOUNT_TYPE = 'V';"
                  .Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

                  If .EOF Then
                     .AddNew
                     .Fields.Item("THIS_RECORD").Value = UniqueID()
                     .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                     .Fields.Item("TRANS_ID").Value = lHeaderID
                     .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                     .Fields.Item("AMOUNT_TYPE").Value = "V"
                     .Fields.Item("USER_NUMBER").Value = "N"   'I dont know what is this field has been used for. just for the time being, i put N; means the transaction has been posted from NJ
                  End If
                  .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
                  .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
                  
                  'Resolved by BOSL
                  'Issue No: 0000482
                  'Assign the Input VAT Account or Output VAT Account based on the VAT Type selected
                  'Modified By: Asif. 01 Oct 2014

                  If flxNJ_Split.TextMatrix(iRow, 14) = "1" Or flxNJ_Split.TextMatrix(iRow, 14) = "3" Then
                     .Fields.Item("NOMINAL_CODE").Value = szInVat
                     
                  ElseIf flxNJ_Split.TextMatrix(iRow, 14) = "2" Or flxNJ_Split.TextMatrix(iRow, 14) = "4" Then
                     .Fields.Item("NOMINAL_CODE").Value = szOutVat
                     
                  End If
                  
                  If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                     .Fields.Item("TRANSACTION_TYPE").Value = 15   'input vat
                     '.Fields.Item("NOMINAL_CODE").Value = szInVat
                     .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10))
                  Else
                     .Fields.Item("TRANSACTION_TYPE").Value = 16
                     '.Fields.Item("NOMINAL_CODE").Value = szOutVat
                     .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10)) * -1
                  End If
                  ''''''''''''''''''''''''
                  
                  .Fields.Item("ACCOUNT_NUMBER").Value = flxNJ_Split.TextMatrix(iRow, 1)
                  .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
                  .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
                  
                  .Fields.Item("REFERENCE").Value = txtTitle.text
                  .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
                  .Fields.Item("ClientID").Value = txtClient.Tag
                  .Fields.Item("DeleteFlag").Value = False
                  .Update
                  .Close
               End If
            Else
               szSQL = "SELECT * FROM NLPosting " & _
                       "WHERE PARENT_RECORD = '" & flxNJ_Split.TextMatrix(iRow, 0) & "' AND " & _
                             "AMOUNT_TYPE = 'V';"
               .Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

               If .EOF Then
                  .AddNew
                  .Fields.Item("THIS_RECORD").Value = UniqueID()
                  .Fields.Item("PARENT_RECORD").Value = flxNJ_Split.TextMatrix(iRow, 0)
                  .Fields.Item("TRANS_ID").Value = lHeaderID
                  .Fields.Item("TRANSACTION_REF").Value = lHeaderID
                  .Fields.Item("AMOUNT_TYPE").Value = "V"
                  .Fields.Item("USER_NUMBER").Value = "N"   'I dont know what is this field has been used for. just for the time being, i put N; means the transaction has been posted from NJ
               End If
               .Fields.Item("POSTED_DATE").Value = Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
               .Fields.Item("TRANSACTION_DATE").Value = Format(txtDateFrom.text, "dd mmmm yyyy")
               
               If flxNJ_Split.TextMatrix(iRow, 4) = "Journal Debit" Then
                  .Fields.Item("TRANSACTION_TYPE").Value = 15
                  .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10))
               Else
                  .Fields.Item("TRANSACTION_TYPE").Value = 16
                  .Fields.Item("Amount").Value = Val(flxNJ_Split.TextMatrix(iRow, 10)) * -1
               End If
               
               'Resolved by BOSL
               'Issue No: 0000482
               'Assign the Input VAT Account or Output VAT Account based on the VAT Type selected
               'Modified By: Asif. 01 Oct 2014
               .Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 1)
               If flxNJ_Split.TextMatrix(iRow, 14) = "1" Or flxNJ_Split.TextMatrix(iRow, 14) = "3" Then
                  .Fields.Item("NOMINAL_CODE").Value = szInVat
                  
               ElseIf flxNJ_Split.TextMatrix(iRow, 14) = "2" Or flxNJ_Split.TextMatrix(iRow, 14) = "4" Then
                  .Fields.Item("NOMINAL_CODE").Value = szOutVat
                  
               End If
                  
               '.Fields.Item("NOMINAL_CODE").Value = flxNJ_Split.TextMatrix(iRow, 10)
               
               .Fields.Item("ACCOUNT_NUMBER").Value = flxNJ_Split.TextMatrix(iRow, 1)
               .Fields.Item("PROPERTY_ID").Value = txtPropID.Tag
               .Fields.Item("FUND_ID").Value = flxNJ_Split.TextMatrix(iRow, 13)
               
               .Fields.Item("REFERENCE").Value = txtTitle.text
               .Fields.Item("TRANSACTION_DESCRIPTION").Value = flxNJ_Split.TextMatrix(iRow, 2)
               .Fields.Item("ClientID").Value = txtClient.Tag
               .Fields.Item("Deleteflag").Value = False
               .Update
               .Close
            End If
         Next iRow
      End With
   End If

   adoconn.CommitTrans
   'added by anol 17 Nov 2015
   'Call ConvertNLPostingAmountToDRCR(adoConn)
  
   flxNJ_Split.Tag = ""
   ShowMsgInTaskBar "Nominal Ledger has been updated", "Y", "P"
   frmNJ.txtClient.text = "ALL"
   frmNJ.txtProperty.text = "ALL"
   Call frmNJ.LoadFlxNJ(adoconn, "")
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
   ConfigflxNJ_SplitDelete
   cmdClose_Click
   'cmdSave.Enabled = True
   Exit Sub
ERR_HANDLER:
   MsgBox ERR.Number & ": " & ERR.description, vbCritical + vbOKOnly, "Updating Nominal Journal"
   adoconn.RollbackTrans
   MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
   
End Sub

Private Function ReturnVatRate(strCode As String) As String
    Dim adoconn As New ADODB.Connection
    Dim rsVat As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsVat.Open "Select vat_Rate from tlbVatCode where vat_code='" & strCode & "'", adoconn, adOpenStatic, adLockReadOnly
    If Not rsVat.EOF Then
        ReturnVatRate = rsVat("vat_Rate").Value
    End If
    rsVat.Close
    Set rsVat = Nothing
    adoconn.Close
End Function
Private Sub flxNJ_Split_DblClick()
'I am adding index +2 for the row
   If flxNJ_Split.TextMatrix(flxNJ_Split.row, 0) = "" Then Exit Sub
   'If txtNominal.text = "" Then Exit Sub
   cmdEdit.Enabled = False
   cmdClear.Enabled = True
   cmdVatCode.Enabled = True
   sNew = 2
   flxNJ_Split.Enabled = False
   cmdClear.Enabled = True

   Dim inRow As Integer

   inRow = flxNJ_Split.row
   txtDescription.text = flxNJ_Split.TextMatrix(inRow, 2)
   txtFund.Tag = flxNJ_Split.TextMatrix(inRow, 13)
   txtVatType.Tag = flxNJ_Split.TextMatrix(inRow, 14)
   'mark 1 here I need to put back name and Id after reading it from the grid
    txtVatType.text = flxNJ_Split.TextMatrix(inRow, 5)
    txtVatType.Tag = returnTypeID(flxNJ_Split.TextMatrix(inRow, 5))
    txtRate.Tag = flxNJ_Split.TextMatrix(inRow, 6) 'like T1
    If flxNJ_Split.TextMatrix(inRow, 6) <> "" Then
         txtRate.text = ReturnVatRate(flxNJ_Split.TextMatrix(inRow, 6))
         txtVatCode.text = txtRate.Tag & " / " & ReturnVatRate(flxNJ_Split.TextMatrix(inRow, 6))
         
    Else
         txtVatCode.text = ""
    End If
    
    txtNominal.Tag = flxNJ_Split.TextMatrix(inRow, 12)
    txtNCCODE.text = flxNJ_Split.TextMatrix(inRow, 12)
    txtVAT.text = flxNJ_Split.TextMatrix(inRow, 10)
    txtTotal.text = IIf(Val(flxNJ_Split.TextMatrix(inRow, 11)) < 0, _
                       Val(flxNJ_Split.TextMatrix(inRow, 11)) * (-1), _
                       Val(flxNJ_Split.TextMatrix(inRow, 11)))                 'Total
   If flxNJ_Split.TextMatrix(inRow, 4) = "Journal Debit" Then
      txtDr.text = flxNJ_Split.TextMatrix(inRow, 7)
      txtDiff.text = Format(Val(txtDiff.text) - Val(txtTotal.text), "0.00")
   Else
      txtCr.text = flxNJ_Split.TextMatrix(inRow, 8)
      txtDiff.text = Format(Val(txtDiff.text) + Val(txtTotal.text), "0.00")
   End If
   If Val(txtDr.text) = 0 And Val(txtCr.text) = 0 Then                        'VAT only transactions
      If flxNJ_Split.TextMatrix(inRow, 4) = "Journal Debit" Then
         txtDr.text = flxNJ_Split.TextMatrix(inRow, 10)
      Else
         txtCr.text = flxNJ_Split.TextMatrix(inRow, 10)
      End If
   End If
   'Mark 1
   'cmbVatRate.Value = flxNJ_Split.TextMatrix(inRow, 13)
   If txtVatType.Tag = "" Then
        txtVatCode.text = ""
        txtVatCode.Tag = ""
        txtRate.text = ""
   Else
   End If
   'added by anol 12 JUN 2016
   txtNominal.text = flxNJ_Split.TextMatrix(inRow, 17)
   txtFund.text = flxNJ_Split.TextMatrix(inRow, 3)
   'End of addition
   cmdSave.Enabled = False
   'added by anol 19 Nov 2015
   txtDiff.text = Format(Val(txtDebit.text) - Val(txtCredit.text), "0.00")
   If txtVatType.text = "N/A" Then
         cmdVatCode.Enabled = False
   End If
   'End of addition
End Sub

Private Sub Form_Activate()
    Dim strTemp As String
    txtClient.ForeColor = vbBlack
'    cmdVATCode.Enabled = False
    strTemp = isControlAccountSet(txtClient.Tag)
    If Len(strTemp) > 0 Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClient.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        strTemp = ""
        txtClient.ForeColor = vbRed
        Exit Sub
    End If
        
   If Label50(9).Caption = "Not Loaded" And lHeaderID > 0 Then
      Dim adoconn As New ADODB.Connection
      Dim adoRst  As New ADODB.Recordset
      Dim adoRst1  As New ADODB.Recordset
      Dim szSQL   As String
      Dim iRow    As Integer
        
        
        
        
      adoconn.Open getConnectionString

      szSQL = "SELECT N.*,Client.ClientName FROM NJ_Header N,Client WHERE N.ClientID=Client.clientID " & _
              "AND RecordID = " & lHeaderID & ";"
      adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
      
      szSQL = "SELECT N.*,Client.ClientName, Property.PropertyName FROM NJ_Header N,Client,Property WHERE N.ClientID=Client.clientID " & _
              "AND N.PropertyID = Property.PropertyID AND RecordID = " & lHeaderID & ";"
      adoRst1.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
      txtClient.Tag = adoRst.Fields.Item("ClientID").Value
      txtClient.text = adoRst.Fields.Item("ClientName").Value
      If Not adoRst1.EOF Then
         txtPropID.Tag = IIf(IsNull(adoRst1.Fields.Item("PropertyID").Value) Or adoRst1.Fields.Item("PropertyID").Value = "", "", adoRst1.Fields.Item("PropertyID").Value)
         txtPropID.text = adoRst1.Fields.Item("PropertyName").Value
      Else
         txtPropID.Tag = ""
         txtPropID.text = ""
      End If
      adoRst1.Close
      Set adoRst1 = Nothing
      txtDateFrom.text = Format(adoRst.Fields.Item("NJDate").Value, "dd/mm/yyyy")
      lblPostingDate.ToolTipText = Format(adoRst.Fields.Item("PostingDate").Value, "dd/mm/yyyy")
      txtTitle.text = adoRst.Fields.Item("NJTitle").Value
      adoRst.Close

'Debug.Print "SELECT DISTINCT S.*, F.FundName, T.DESCRIPTION " & _
            "FROM NJ_Split AS S, Fund AS F, tlbTransactionTypes AS T, NLPosting AS N " & _
            "WHERE S.ParentID = " & lHeaderID & " AND " & _
                  "S.FundID = F.FundID AND " & _
                  "S.RecordID = N.PARENT_RECORD AND " & _
                  "N.TRANSACTION_TYPE = T.TYPE_ID;"

'   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund|<Type|>Dr|>Cr|>VATRate" & _
'                    0           1           2         3     4    5   6     7
'               "|>VATAmt|>TotalAmt|NC|Fund|VatType|VRate|VatOnly"
'                    8        9     10  11    12      13    14
'issue 619 ,there he using nlposting table, this is the main record source for amount..2018/07/19 by anol and also type ID is not found in the split table .

'      adoRST.Open "SELECT DISTINCT S.* ,F.FundName, T.DESCRIPTION " & _
'                  "FROM NJ_Split AS S, Fund AS F, tlbTransactionTypes AS T, NLPosting AS N " & _
'                  "WHERE S.ParentID = " & lHeaderID & " AND " & _
'                        "S.FundID = F.FundID AND " & _
'                        "S.RecordID = N.PARENT_RECORD AND " & _
'                        "N.TRANSACTION_TYPE = T.TYPE_ID;"
'    Debug.Print time
'    Debug.Print lHeaderID
    Debug.Print "SELECT DISTINCT S.* ,F.FundName, T.DESCRIPTION " & _
                  "FROM NJ_Split AS S, Fund AS F, tlbTransactionTypes AS T, NLPosting AS N " & _
                  "WHERE S.ParentID = " & lHeaderID & " AND " & _
                        "S.FundID = F.FundID AND " & _
                        "S.RecordID = N.PARENT_RECORD AND " & _
                        "N.TRANSACTION_TYPE = T.TYPE_ID AND N.DeleteFlag=false ;"
                        
                        
      adoRst.Open "SELECT DISTINCT S.* ,F.FundName, T.DESCRIPTION " & _
                  "FROM NJ_Split AS S, Fund AS F, tlbTransactionTypes AS T, NLPosting AS N " & _
                  "WHERE S.ParentID = " & lHeaderID & " AND " & _
                        "S.FundID = F.FundID AND " & _
                        "S.RecordID = N.PARENT_RECORD AND " & _
                        "N.TRANSACTION_TYPE = T.TYPE_ID AND N.DeleteFlag=false ;"

'        Debug.Print time
                        
'       adoRst.Open "SELECT DISTINCT S.* ,F.FundName, T.DESCRIPTION,V.Name " & _
'                  "FROM NJ_Split AS S, Fund AS F, tlbTransactionTypes AS T, NLPosting AS N,NominalLedger V " & _
'                  "WHERE S.ParentID = " & lHeaderID & " AND " & _
'                        "S.FundID = F.FundID AND v.Code=S.NC AND V.ClientID='" & txtClient.Tag & "'" & _
'                        "S.RecordID = N.PARENT_RECORD AND " & _
'                        "N.TRANSACTION_TYPE = T.TYPE_ID;"
                        '"N.ClientID = NJ.ClientID AND NJ.CODE=N.Nominal_code AND "
      iRow = 1
      While Not adoRst.EOF
         flxNJ_Split.TextMatrix(iRow, 0) = adoRst.Fields.Item("RecordID").Value
         flxNJ_Split.TextMatrix(iRow, 1) = adoRst.Fields.Item("NC").Value
         flxNJ_Split.TextMatrix(iRow, 2) = adoRst.Fields.Item("SpLineDes").Value
         flxNJ_Split.TextMatrix(iRow, 3) = adoRst.Fields.Item("FundName").Value
         flxNJ_Split.TextMatrix(iRow, 4) = adoRst.Fields.Item("Description").Value
         flxNJ_Split.TextMatrix(iRow, 5) = returnType(adoRst.Fields.Item("TYPE_ID").Value)
         flxNJ_Split.TextMatrix(iRow, 6) = adoRst.Fields.Item("VAT_CODE").Value 'vat code and number 5 is for vat type
         
         'flxNJ_Split.TextMatrix(inRow, 4) = IIf(Val(adoRST.Fields.Item("NetAmt").Value) > 0, "Journal Debit", "Journal Credit")
         If InStr(flxNJ_Split.TextMatrix(iRow, 4), "Debit") > 0 Then
            flxNJ_Split.TextMatrix(iRow, 7) = IIf(Val(adoRst.Fields.Item("NetAmt").Value) = 0, "", Format(Val(adoRst.Fields.Item("NetAmt").Value), "0.00"))
            flxNJ_Split.TextMatrix(iRow, 11) = Format(adoRst.Fields.Item("TotalAmt").Value, "0.00")
         Else
            flxNJ_Split.TextMatrix(iRow, 8) = IIf(Val(adoRst.Fields.Item("NetAmt").Value) = 0, "", Format(Val(adoRst.Fields.Item("NetAmt").Value), "0.00"))
            flxNJ_Split.TextMatrix(iRow, 11) = Format(Val(adoRst.Fields.Item("TotalAmt").Value) * (-1), "0.00")
         End If
         flxNJ_Split.TextMatrix(iRow, 10) = Format(adoRst.Fields.Item("VATAmt").Value, "0.00")
         flxNJ_Split.TextMatrix(iRow, 12) = adoRst.Fields.Item("NC").Value
         flxNJ_Split.TextMatrix(iRow, 13) = adoRst.Fields.Item("FundID").Value
         flxNJ_Split.TextMatrix(iRow, 14) = adoRst.Fields.Item("TYPE_ID").Value 'this is vat type ID
         flxNJ_Split.TextMatrix(iRow, 15) = adoRst.Fields.Item("VAT_CODE").Value
         Debug.Print adoRst.Fields.Item("VAT_CODE").Value
         If Val(flxNJ_Split.TextMatrix(iRow, 7)) > 0 Or Val(flxNJ_Split.TextMatrix(iRow, 8)) > 0 Then
            flxNJ_Split.TextMatrix(iRow, 16) = "N"
         Else
            If InStr(flxNJ_Split.TextMatrix(iRow, 4), "Debit") > 0 Then
               flxNJ_Split.TextMatrix(iRow, 16) = "DV"
            Else
               flxNJ_Split.TextMatrix(iRow, 16) = "CV"
            End If
         End If
         Dim rsCheck As New ADODB.Recordset
         rsCheck.Open "Select Name from NominalLedger where NominalLedger.Code='" & adoRst.Fields.Item("NC").Value & "' AND NominalLedger.clientID='" & txtClient.Tag & "'", adoconn, adOpenStatic, adLockReadOnly
         flxNJ_Split.TextMatrix(iRow, 17) = rsCheck.Fields.Item("Name").Value
         rsCheck.Close
         adoRst.MoveNext
         iRow = iRow + 1
         If Not adoRst.EOF Then flxNJ_Split.AddItem ""
      Wend
      adoRst.Close
      Set adoRst = Nothing
      adoconn.Close
      Set adoconn = Nothing
      cmdVatCode.Enabled = False
   'Dim iRow As Integer
   frmNJ_Entry.txtDebit.text = "0.00"
   frmNJ_Entry.txtCredit.text = "0.00"
   For iRow = 1 To frmNJ_Entry.flxNJ_Split.Rows - 1
        frmNJ_Entry.txtDebit.text = frmNJ_Entry.txtDebit.text + Val(frmNJ_Entry.flxNJ_Split.TextMatrix(iRow, 7)) + IIf(Val(flxNJ_Split.TextMatrix(iRow, 7)) > 0, Val(flxNJ_Split.TextMatrix(iRow, 10)), 0)
        frmNJ_Entry.txtCredit.text = frmNJ_Entry.txtCredit.text + Val(frmNJ_Entry.flxNJ_Split.TextMatrix(iRow, 8)) + IIf(Val(flxNJ_Split.TextMatrix(iRow, 8)) > 0, Val(flxNJ_Split.TextMatrix(iRow, 10)), 0)
   Next
   'added by anol 18 Nov 2015
    frmNJ_Entry.txtDiff.text = Val(frmNJ_Entry.txtDebit.text) - Val(frmNJ_Entry.txtCredit.text)
'    cmdSave.Enabled = False
   'End of addition
      txtVatType.Tag = ""
      txtVatType.text = ""
      Label50(9).Caption = "Loaded"
   End If
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 8655 '6840
   Me.Width = 15655
   txtRate.text = "0.00"
   txtVatType.Tag = 0
   txtVatType.text = "N/A"
   cmdVatCode.Enabled = True

   sNew = 1
   ConfigflxNJ_SplitDelete
   Me.BackColor = MODULEBACKCOLOR
   lblNJ_Id.BackColor = MODULEBACKCOLOR
   fraHeader.BackColor = MODULEBACKCOLOR
   ConfigFlxNJ_Split

   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   
   'LoadCmbVatRate adoConn
   'We are not loadiing vat rate at the form load, this shall be empty when you set vat type N/A
   LoadCmbClient adoconn 'Load the initial value of Client
 '  LoadCmbProperty adoConn 'Load the initial value of property
   If txtClient.text <> "" Then
'        loadNCInit adoConn
'        LoadCmbFund adoConn
        txtNominal.text = ""
        txtNCCODE.text = ""
        txtFund.text = ""
        txtFund.Tag = ""
        LoadIO_VAT adoconn
        configflxVatType
        LoadCmbVatType
    End If
   
   adoconn.Close
   Set adoconn = Nothing

  
   Label50(9).Caption = "Not Loaded"
   If Len(txtDateFrom.text) < 10 Then txtDateFrom.text = Format(Date, "dd/mm/yyyy")
   If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
        txtRate.Visible = False
   End If
   Call WheelHook(Me.hWnd)
End Sub
Private Sub loadNCInit(adoconn As ADODB.Connection)
     Dim rstRec As New ADODB.Recordset
     Dim szSQL As String
     szSQL = "SELECT N.* " & _
      "FROM NominalLedger AS N " & _
      "WHERE N.ClientID = '" & txtClient.Tag & "' AND " & _
      "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClient.Tag & "')" & _
      " ORDER BY N.Code;"
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not rstRec.EOF Then
            txtNominal.Tag = rstRec.Fields.Item(0).Value
            txtNCCODE.text = rstRec.Fields.Item(0).Value
            txtNominal.text = rstRec.Fields.Item(1).Value
   End If
End Sub
Private Sub ConfigFlxNJ_Split()
   Dim szHeader As String, iCol As Integer

   flxNJ_Split.Clear
   flxNJ_Split.Cols = 18
   flxNJ_Split.Rows = 2
   flxNJ_Split.RowHeight(0) = 0

   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund|<Type|>Dr|>Cr|>VATRate" & _
               "|>VATAmt|>TotalAmt|NC|Fund|VatType|VRate|VatOnly|NominalName"

'   lblGridCaption(7).Width = 0
'   lblGridCaption(8).Left = lblGridCaption(7).Left

   flxNJ_Split.FormatString = szHeader$
   flxNJ_Split.ColWidth(0) = 100
   
'   For iCol = 1 To flxNJ_Split.Cols - 8
'      flxNJ_Split.ColWidth(iCol) = lblGridCaption(iCol + 1).Left - lblGridCaption(iCol).Left
'      Debug.Print "flxNJ_Split.ColWidth(" & iCol & ") =" & lblGridCaption(iCol + 1).Left - lblGridCaption(iCol).Left
'      'Debug.Print "=" & lblGridCaption(iCol + 1).Left - lblGridCaption(iCol).Left
'      'lblGridCaption(iCol).Width = flxNJ_Split.ColWidth(iCol)
'   Next iCol
   
    flxNJ_Split.ColWidth(1) = 1040
    flxNJ_Split.ColWidth(2) = 3840
    flxNJ_Split.ColWidth(3) = 2360
    flxNJ_Split.ColWidth(4) = 0
    flxNJ_Split.ColWidth(5) = 1360
    flxNJ_Split.ColWidth(6) = 960
    flxNJ_Split.ColWidth(7) = 1220
    flxNJ_Split.ColWidth(8) = 1220
    flxNJ_Split.ColWidth(9) = 0
    flxNJ_Split.ColWidth(10) = 1200
   'flxNJ_Split.ColWidth(iCol) = flxNJ_Split.Width + flxNJ_Split.Left - lblGridCaption(9).Left - 300  'TotalAmount
   flxNJ_Split.ColWidth(11) = flxNJ_Split.Width + flxNJ_Split.Left - lblGridCaption(9).Left   'TotalAmount
'   lblGridCaption(iCol).Width = flxNJ_Split.ColWidth(iCol)
   flxNJ_Split.ColWidth(12) = 0
   flxNJ_Split.ColWidth(13) = 0
   flxNJ_Split.ColWidth(14) = 0
   flxNJ_Split.ColWidth(15) = 0
   flxNJ_Split.ColWidth(16) = 0
   flxNJ_Split.ColWidth(17) = 0

   'txtDiff.Width = flxNJ_Split.ColWidth(iCol)
   txtDiff.Left = lblGridCaption(9).Left
   Label50(16).Left = txtDiff.Left - Label50(16).Width - 100
End Sub
Private Sub ConfigflxNJ_SplitDelete()
   Dim szHeader As String, iCol As Integer

   flxNJ_SplitDelete.Clear
   flxNJ_SplitDelete.Cols = 18
   flxNJ_SplitDelete.Rows = 2
   flxNJ_SplitDelete.RowHeight(0) = 0

   szHeader$ = "HeaderID|<NominalCode|<Description|<Fund|<Type|>Dr|>Cr|>VATRate" & _
               "|>VATAmt|>TotalAmt|NC|Fund|VatType|VRate|VatOnly|NominalName"

'   lblGridCaption(7).Width = 0
'   lblGridCaption(8).Left = lblGridCaption(7).Left

   flxNJ_SplitDelete.FormatString = szHeader$
   flxNJ_SplitDelete.ColWidth(0) = 0
'   For iCol = 1 To flxNJ_SplitDelete.Cols - 6
'      flxNJ_SplitDelete.ColWidth(iCol) = lblGridCaption(iCol + 1).Left - lblGridCaption(iCol).Left
'      lblGridCaption(iCol).Width = flxNJ_SplitDelete.ColWidth(iCol)
'   Next iCol
   flxNJ_SplitDelete.ColWidth(iCol) = flxNJ_SplitDelete.Width + flxNJ_SplitDelete.Left - lblGridCaption(9).Left - 300  'TotalAmount
   lblGridCaption(iCol).Width = flxNJ_SplitDelete.ColWidth(iCol)
   flxNJ_SplitDelete.ColWidth(iCol + 1) = 0
   flxNJ_SplitDelete.ColWidth(iCol + 2) = 0
   flxNJ_SplitDelete.ColWidth(iCol + 3) = 0
   flxNJ_SplitDelete.ColWidth(iCol + 4) = 0
   flxNJ_SplitDelete.ColWidth(iCol + 5) = 0
   flxNJ_SplitDelete.ColWidth(iCol + 6) = 0

'   txtDiff.Width = flxNJ_SplitDelete.ColWidth(iCol)
'   txtDiff.Left = lblGridCaption(9).Left
'   Label50(16).Left = txtDiff.Left - Label50(16).Width - 100
End Sub
Private Sub LoadCmbClient(adoconn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
        txtClient.text = adoRst.Fields("CLIENTNAME").Value
        txtClient.Tag = adoRst.Fields("CLIENTID").Value
        txtPropID.Tag = ""
        txtPropID.text = ""
   End If
End Sub

'Private Sub LoadCmbVatRate(adoConn As adodb.Connection)
'   Dim Data() As String
'
'   Dim szSQL As String
'   Dim adoRst As New adodb.Recordset
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   szSQL = "SELECT VAT_CODE, VAT_RATE FROM tlbVatCode WHERE IN_USE;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
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
'   cmbVatRate.Column() = Data()
'   cmbVatRate.ListIndex = 0
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
    UnLoadForm Me
   Label50(9).Caption = "Not Loaded"

   frmNJ.Enabled = True

   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer

   
'**At the end --> Release the NLPosting table******************************************************
'   adoRst.Open "SELECT Field1, Field2 FROM ShoppingCentre WHERE Field1 = 'SYSTEMUSER:" & SystemUser & "' AND Field2 = 'SYSTEMNAME:" & WS_Name & "';", adoConn, adOpenStatic, adLockReadOnly
'
'   If Not adoRst.EOF Then
'      adoConn.Execute "UPDATE ShoppingCentre SET Field1 = '', Field2 = '';"
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
    frmNJ.flxNJ.row = frmNJ.iSelRow
    frmNJ.flxNJ.col = 0
    frmNJ.flxNJ.CellBackColor = vbWhite
   If bEditMode = True Then
        adoconn.Open getConnectionString
        adoconn.Execute "Update NJ_Header Set  DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName=''," & _
                "PrestigeUserName='',ServerIPaddress='' where UserSessionID='" & frmNJ.UserSessionID & "'"
                 frmNJ.flxNJ.row = frmNJ.iSelRow
                 frmNJ.flxNJ.col = 0
                 frmNJ.flxNJ.CellBackColor = vbWhite
'       adoConn.Execute "Update NJ_Header Set  DateTimeStamp='" & Now & "',Module='Nominal Journal',UserSessionID='" & UserSessionID & "',WindowsUserName='" & SystemUser & "',MachineName='" & WS_Name & "'," & _
'                "PrestigeUserName='" & User & "',ServerIPaddress='" & GetIPaddress & "' where RecordID = " & szRecordID & ""
                
             bEditMode = False
        adoconn.Close
        Set adoconn = Nothing
   End If
  
   frmNJ_Entry.lblNJ_Id.Visible = False
   'lHeaderID = 0
                If Trim(txtClient.text) = "" And flxNJ_Split.Tag = "1" And Val(txtDiff.text) = 0 Then
                    MsgBox "Please select a valid client", vbInformation, "Please select a property"
                    Cancel = True
                    Exit Sub
                End If
   If flxNJ_Split.Tag = "1" And Val(txtDiff.text) = 0 Then
      If MsgBox("Do you wish to save this journal?", vbYesNo, "Save Journal? ") = vbNo Then
             
      Else
            If cmdSave.Enabled = True Then
                cmdSave_Click
            End If
      End If
      frmNJ.ReloadFlxNJ
   ElseIf flxNJ_Split.Tag = "1" And Val(txtDiff.text) <> 0 Then
        If MsgBox("There is difference on this journal. Do you wish to complete this journal?", vbYesNo, "Save Journal? ") = vbNo Then
             
      Else
            
                Cancel = True
                
      End If
   End If
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClient.Tag = "" Then
          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
          Exit Sub
    End If
    If IsDate(lblPostingDate.ToolTipText) = False Then Exit Sub
    If frmMMain.IsRibbonVersion Then
        Dim adoconn As New ADODB.Connection
        Dim szSQL As String
        adoconn.Open getConnectionString
        If IsPeriodStatus(lblPostingDate.ToolTipText, txtClient.Tag, adoconn) = 0 Then
           ShowMsgInTaskBar "The posting date of this transaction falls within a closed financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(lblPostingDate.ToolTipText, txtClient.Tag, adoconn) = 9 Then
           ShowMsgInTaskBar "The posting date of this transaction does not fall in any existing financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
    End If
    DispayCalendar Me, lblPostingDate.ToolTipText, txtDateFrom.text, txtClient.Tag
End Sub

Private Sub txtCr_Change()
    If Val(txtCr.text) > 0 Then txtDr.text = ""
    If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
       txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
    End If
    If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
       txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
    End If
    If Val(txtCr.text) = 0 And Val(txtDr.text) = 0 Then
         txtTotal.text = "0.00"
    End If
End Sub

Private Sub txtCr_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
   SelTxtInCtrl txtCr
End Sub

Private Sub txtCr_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'        KeyAscii = 0
'   End If
   If KeyAscii = 13 Then
        cmdVatType.Enabled = True
        cmdVatType.SetFocus
   End If
   DigitTextKeyPress txtCr, KeyAscii
End Sub

Private Sub txtCr_LostFocus()
    If Val(txtCr.text) > 0 And txtVatType.text = "" Then
        txtVatType.text = "N/A"
        txtVatType.Tag = "0"
    End If
    If Val(txtCr.text) < 0 Then txtCr.text = Val(txtCr.text) * (-1)
    txtCr.text = IIf(txtCr.text = "", "0.00", Format(txtCr.text, "0.00"))
    'Modified By BOSL
    'issue 463
    'Anol 27 Aug 2014
    txtVAT.text = Format(txtVAT.text, "0.00")
    If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
        txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
    End If
    If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
        txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
    End If
      

End Sub

Private Sub txtDr_Change()
   If Val(txtDr.text) > 0 Then txtCr.text = "0.00"
End Sub

Private Sub txtDr_GotFocus()
   txtNominalDesc.Visible = False
   Label50(17).Visible = False
   SelTxtInCtrl txtDr
End Sub

Private Sub txtDr_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'       txtCr.SetFocus
'   End If
   If KeyAscii = 13 Then
        txtCr.Enabled = True
        txtCr.SetFocus
   End If
   
   DigitTextKeyPress txtDr, KeyAscii
End Sub

Private Sub txtDr_LostFocus()
    If Val(txtDr.text) > 0 And txtVatType.text = "" Then
        txtVatType.text = "N/A"
        txtVatType.Tag = "0"
    End If
   If Val(txtDr.text) < 0 Then txtDr.text = Val(txtDr.text) * (-1)

   txtDr.text = IIf(txtDr.text = "", "0.00", Format(txtDr.text, "0.00"))

    'Modified By BOSL
    'issue 463
    'Anol 27 Aug 2014
    If Val(txtRate.text) > 0 Then
          If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
                txtVAT.text = Format(Val(txtDr.text) * (Val(txtRate.text) / 100), "0.00")
          End If
          If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
                txtVAT.text = Format(Val(txtCr.text) * (Val(txtRate.text) / 100), "0.00")
          End If
    End If
    txtVAT.text = Format(txtVAT.text, "0.00")
    If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
         txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
    End If
    If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
         txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
    End If

End Sub

Private Sub txtDateFrom_Change()
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
   TextBoxChangeDate txtDateFrom
   lblPostingDate.ToolTipText = txtDateFrom.text
End Sub

Private Sub txtDateFrom_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
   If txtDateFrom.text = "dd/mm/yyyy" Then
      txtDateFrom.text = ""
      Exit Sub
   End If
   If Len(txtDateFrom.text) < 10 Then txtDateFrom.text = Format(Date, "dd/mm/yyyy")
   If lblPostingDate.ToolTipText = "" Then lblPostingDate.ToolTipText = txtDateFrom.text
   SelTxtInCtrl txtDateFrom
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTitle.SetFocus
    End If
   TextBoxKeyPrsDate txtDateFrom, KeyAscii
End Sub

Private Sub txtDateFrom_LostFocus()
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClient.Tag = "" Then
          ShowMsgInTaskBar "Please select a valid client to proceed.", "Y", "N"
          Exit Sub
    End If
    If frmMMain.IsRibbonVersion And IsDate(txtDateFrom.text) = True Then
        Dim adoconn As New ADODB.Connection
        Dim szSQL As String
        adoconn.Open getConnectionString
        If IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 0 Then
           ShowMsgInTaskBar "The transaction date cannot fall within a closed financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDateFrom.text, txtClient.Tag, adoconn) = 9 Then
           ShowMsgInTaskBar "The transaction date does not fall in any existing financial period", "Y", "N"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
    End If
   If txtDateFrom.text <> "" Then TextBoxFormatDate txtDateFrom
   'lblPostingDate.ToolTipText = txtDateFrom.text
End Sub

Private Sub txtVAT_GotFocus()
    txtNominalDesc.Visible = False
    Label50(17).Visible = False
   SelTxtInCtrl txtVAT
End Sub

Private Sub txtVAT_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtVAT, KeyAscii
End Sub

Private Sub txtVAT_LostFocus()
  'Resolved by BOSL
'issue 463
'Modified by anol 20 Aug 2014

      txtVAT.text = Format(txtVAT.text, "0.00")
      If Val(txtDr.text) > 0 And Val(txtCr.text) = 0 Then
         txtTotal.text = Format(Val(txtDr.text) + Val(txtVAT.text), "0.00")
      End If
      If Val(txtCr.text) > 0 And Val(txtDr.text) = 0 Then
         txtTotal.text = Format(Val(txtCr.text) + Val(txtVAT.text), "0.00")
      End If
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
Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
      flxClient.RowHeight(i) = 240
     ' If sTextBox = "1" Then
            If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
                flxClient.RowHeight(i) = 0
            End If
       'End If
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
          '  If sTextBox = "11" Or sTextBox = "14" Then
                
                'flxClient.SetFocus
'                flxClient.row = 1
'                flxClient.RowSel = 1
'                flxClient.ColSel = 1
'                flxClient.CellBackColor = RGB(174, 179, 233)
                
'                Dim iRow As Integer
'                flxClient.row = 1
'                For iRow = 1 To flxClient.Cols - 1
'                   flxClient.col = iRow
'                   flxClient.CellBackColor = RGB(174, 179, 233)
'                Next iRow
               ' SelectOnly1RowFlxGrid flxClient, 1 'flxClient.row
           ' Else
                txtSearchClientName.SetFocus
           
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'         txtSearchClientName.SetFocus
'    End If
    If KeyAscii = 27 Then
          flxClient.Clear
          flxClient.Cols = 2
          flxClient.Rows = 2
          picClient.Visible = False
          fraHeader.Enabled = True
          Frame1.Enabled = True
          If sTextBox = "1" Then
                cmdClient.SetFocus
          ElseIf sTextBox = "2" Then
                cmdPropID.SetFocus
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
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 3600
   
   
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
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
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
               flxClient.TextMatrix(rRow, 0) = "" 'rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value 'IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
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
Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    fraHeader.Enabled = True
    Frame1.Enabled = True
    fraHeader.Enabled = True
    Frame1.Enabled = True
    Frame3.Enabled = True
    Frame2.Enabled = True
    cmdPropID.SetFocus
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
   flxClient.ColWidth(2) = 3600
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   picClient.Height = 4095
   flxClient.Height = 3345
   'lblJobName.Visible = False
   
      adoconn.Open getConnectionString
           
     If sTextBox = "2" Then
             If txtClient.Tag = "ALL" Then
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "ORDER BY PropertyID;"
             Else
                 szSQL = "SELECT PropertyID, PropertyName " & _
                 "FROM Property " & _
                 "WHERE ClientID = '" & txtClient.Tag & "' " & _
                 "ORDER BY PropertyID;"

             End If
    End If

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
  
   If sTextBox = "2" Then
           
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
            flxClient.AddItem ""
            flxClient.TextMatrix(rRow, 0) = ""
            flxClient.TextMatrix(rRow, 1) = ""
           
   End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

