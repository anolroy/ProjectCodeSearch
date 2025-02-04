VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUnits2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unit Maintenance"
   ClientHeight    =   13035
   ClientLeft      =   2745
   ClientTop       =   1470
   ClientWidth     =   14145
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnits2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13383.53
   ScaleMode       =   0  'User
   ScaleWidth      =   14145
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   2484
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   287
      Top             =   4212
      Visible         =   0   'False
      Width           =   6285
      Begin VB.CommandButton Command1 
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
         TabIndex        =   288
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxUnitType 
         Height          =   3480
         Left            =   45
         TabIndex        =   289
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6138
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
      Begin MSForms.Label lblName 
         Height          =   195
         Left            =   1665
         TabIndex        =   295
         Top             =   120
         Width           =   1590
         VariousPropertyBits=   8388627
         Caption         =   "Unit Type"
         Size            =   "2805;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   294
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   293
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.TextBox txtSearchUnitTypeCode 
         Height          =   252
         Left            =   72
         TabIndex        =   292
         Top             =   360
         Width           =   1536
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2699;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtUnitTypeName 
         Height          =   252
         Left            =   1620
         TabIndex        =   291
         Top             =   360
         Width           =   4548
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblCode 
         Height          =   192
         Left            =   72
         TabIndex        =   290
         Top             =   108
         Width           =   1344
         VariousPropertyBits=   8388627
         Caption         =   "Unit Type"
         Size            =   "2371;339"
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
         Index           =   0
         Left            =   0
         Top             =   72
         Width           =   5856
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   6930
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   277
      Top             =   3510
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
         TabIndex        =   278
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3480
         Left            =   45
         TabIndex        =   279
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6138
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
      Begin MSForms.Label Label3 
         Height          =   195
         Left            =   1665
         TabIndex        =   276
         Top             =   120
         Width           =   1590
         VariousPropertyBits=   8388627
         Caption         =   "Property Name"
         Size            =   "2805;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   282
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   284
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   283
         Top             =   120
         Width           =   1410
         VariousPropertyBits=   8388627
         Caption         =   "Property ID"
         Size            =   "2487;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   281
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
         TabIndex        =   280
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
   Begin TabDlg.SSTab tabUnits 
      Height          =   4935
      Left            =   75
      TabIndex        =   70
      Top             =   3480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   9
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Unit Details"
      TabPicture(0)   =   "frmUnits2.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Tenancy"
      TabPicture(1)   =   "frmUnits2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLeaseHeading"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame1(0)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Maintenance History"
      TabPicture(2)   =   "frmUnits2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraJS_PO"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Occupancy History"
      TabPicture(3)   =   "frmUnits2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblMainUnit(14)"
      Tab(3).Control(1)=   "lblMainUnit(15)"
      Tab(3).Control(2)=   "lblMainUnit(16)"
      Tab(3).Control(3)=   "lblMainUnit(17)"
      Tab(3).Control(4)=   "lblMainUnit(18)"
      Tab(3).Control(5)=   "gridACHistory"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "&Health && Safety"
      TabPicture(4)   =   "frmUnits2.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label43"
      Tab(4).Control(1)=   "Label41(2)"
      Tab(4).Control(2)=   "Label41(0)"
      Tab(4).Control(3)=   "Label41(4)"
      Tab(4).Control(4)=   "Label41(5)"
      Tab(4).Control(5)=   "Label41(7)"
      Tab(4).Control(6)=   "Label41(1)"
      Tab(4).Control(7)=   "Label41(3)"
      Tab(4).Control(8)=   "Label41(6)"
      Tab(4).Control(9)=   "VerticalLabel(0)"
      Tab(4).Control(10)=   "VerticalLabel(1)"
      Tab(4).Control(11)=   "VerticalLabel(2)"
      Tab(4).Control(12)=   "cboInspectedBy"
      Tab(4).Control(13)=   "cboSchedule"
      Tab(4).Control(14)=   "cboSafetyType"
      Tab(4).Control(15)=   "gridSafety"
      Tab(4).Control(16)=   "txtUnitSafetyID"
      Tab(4).Control(17)=   "chkCertificate"
      Tab(4).Control(18)=   "cmdSafetyNew"
      Tab(4).Control(19)=   "cmdSafetyEdit"
      Tab(4).Control(20)=   "cmdSafetyCancel"
      Tab(4).Control(21)=   "cmdSafetySave"
      Tab(4).Control(22)=   "cmdSafety"
      Tab(4).Control(23)=   "txtRef"
      Tab(4).Control(24)=   "txtNextDueDate"
      Tab(4).Control(25)=   "txtSafetyTelephone"
      Tab(4).Control(26)=   "txtDateChk"
      Tab(4).Control(27)=   "txtComment"
      Tab(4).Control(28)=   "chkAlarm"
      Tab(4).Control(29)=   "cmdInspectedBy"
      Tab(4).Control(30)=   "cmdAttachment"
      Tab(4).ControlCount=   31
      TabCaption(5)   =   "U&tilities"
      TabPicture(5)   =   "frmUnits2.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label82(7)"
      Tab(5).Control(1)=   "Label82(2)"
      Tab(5).Control(2)=   "Label83"
      Tab(5).Control(3)=   "Label82(1)"
      Tab(5).Control(4)=   "Label82(8)"
      Tab(5).Control(5)=   "Label82(3)"
      Tab(5).Control(6)=   "Label82(9)"
      Tab(5).Control(7)=   "Label82(6)"
      Tab(5).Control(8)=   "Label82(0)"
      Tab(5).Control(9)=   "Label82(4)"
      Tab(5).Control(10)=   "Label82(5)"
      Tab(5).Control(11)=   "Label82(10)"
      Tab(5).Control(12)=   "cboUnitUtilityStatus"
      Tab(5).Control(13)=   "cboAuthority_Supplier"
      Tab(5).Control(14)=   "cmdSetUtilitiesType"
      Tab(5).Control(15)=   "cboUtilitiesType"
      Tab(5).Control(16)=   "gridUtilities"
      Tab(5).Control(17)=   "txtChargeRate"
      Tab(5).Control(18)=   "txtUtilitiesReference"
      Tab(5).Control(19)=   "cmdUtilitiesSave"
      Tab(5).Control(20)=   "cmdUtilitiesCancel"
      Tab(5).Control(21)=   "cmdUtilitiesEdit"
      Tab(5).Control(22)=   "cmdUtilitiesNew"
      Tab(5).Control(23)=   "txtUnitUtilitiesID"
      Tab(5).Control(24)=   "txtUnitUtilityIniReading"
      Tab(5).Control(25)=   "txtFinalReading"
      Tab(5).Control(26)=   "txtDateVacated"
      Tab(5).Control(27)=   "cmdUStatus"
      Tab(5).Control(28)=   "txtUnitUtilityStDt"
      Tab(5).Control(29)=   "txtUnitUtilityCom"
      Tab(5).ControlCount=   30
      TabCaption(6)   =   "&Insurance"
      TabPicture(6)   =   "frmUnits2.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label46"
      Tab(6).Control(1)=   "gridInsurance"
      Tab(6).Control(2)=   "txtPropertyInsuranceID"
      Tab(6).Control(3)=   "cmdInsuranceSave"
      Tab(6).Control(4)=   "cmdInsuranceCancel"
      Tab(6).Control(5)=   "cmdInsuranceEdit"
      Tab(6).Control(6)=   "cmdInsuranceNew"
      Tab(6).Control(7)=   "fraInsurance"
      Tab(6).ControlCount=   8
      TabCaption(7)   =   "M&emo"
      TabPicture(7)   =   "frmUnits2.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label81"
      Tab(7).Control(1)=   "cmdUnitMemoCancel"
      Tab(7).Control(2)=   "cmdUnitMemoSave"
      Tab(7).Control(3)=   "cmdUnitMemoEdit"
      Tab(7).Control(4)=   "txtUnitMemo"
      Tab(7).Control(5)=   "Frame17"
      Tab(7).ControlCount=   6
      Begin VB.Frame fraJS_PO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   -67800
         TabIndex        =   263
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
         Begin VB.CommandButton cmdQuoteReq 
            Caption         =   "Quote Request"
            Height          =   300
            Left            =   60
            TabIndex        =   266
            Top             =   780
            Width           =   1335
         End
         Begin VB.CommandButton cmdAsJS 
            Caption         =   "Job Sheet"
            Height          =   300
            Left            =   60
            TabIndex        =   264
            Top             =   60
            Width           =   1335
         End
         Begin VB.CommandButton cmdAsPO 
            Caption         =   "Purchase Order"
            Height          =   300
            Left            =   60
            TabIndex        =   265
            Top             =   420
            Width           =   1335
         End
      End
      Begin VB.Frame fraInsurance 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   975
         Left            =   -74920
         TabIndex        =   229
         Top             =   360
         Width           =   12255
         Begin VB.CommandButton cmdUtilitiesAttach 
            Caption         =   "::"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11620
            TabIndex        =   243
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtExpiryDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7065
            ScrollBars      =   2  'Vertical
            TabIndex        =   238
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtStartDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6165
            ScrollBars      =   2  'Vertical
            TabIndex        =   237
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtAnnualPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5265
            ScrollBars      =   2  'Vertical
            TabIndex        =   236
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtSumInsured 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4290
            ScrollBars      =   2  'Vertical
            TabIndex        =   235
            Top             =   555
            Width           =   990
         End
         Begin VB.TextBox txtPolicyNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3150
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   234
            Top             =   555
            Width           =   1140
         End
         Begin VB.CommandButton cmdSetInsuranceType 
            Caption         =   "..."
            Height          =   315
            Left            =   2900
            TabIndex        =   233
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   10305
            ScrollBars      =   2  'Vertical
            TabIndex        =   242
            Top             =   555
            Width           =   1300
         End
         Begin VB.CommandButton cmdSetInsurer 
            Caption         =   "..."
            Height          =   315
            Left            =   1320
            TabIndex        =   231
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtTelephone 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7965
            ScrollBars      =   2  'Vertical
            TabIndex        =   239
            Top             =   555
            Width           =   1000
         End
         Begin VB.CommandButton cmdUsage 
            Caption         =   "..."
            Height          =   315
            Left            =   10080
            TabIndex        =   241
            Top             =   555
            Width           =   255
         End
         Begin MSDataListLib.DataCombo cboInsurer 
            Bindings        =   "frmUnits2.frx":09AA
            DataSource      =   "adoInsurer"
            Height          =   315
            Left            =   120
            TabIndex        =   230
            Top             =   555
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboInsuranceType 
            Bindings        =   "frmUnits2.frx":09C3
            DataSource      =   "adoInsuranceType"
            Height          =   315
            Left            =   1575
            TabIndex        =   232
            Top             =   555
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboUsage 
            Bindings        =   "frmUnits2.frx":09E2
            DataSource      =   "adoInsUsage"
            Height          =   315
            Left            =   8970
            TabIndex        =   240
            Top             =   555
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Expiry Date"
            Height          =   495
            Index           =   6
            Left            =   7065
            TabIndex        =   254
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Start Date"
            Height          =   435
            Index           =   5
            Left            =   6165
            TabIndex        =   253
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Annual PR"
            Height          =   495
            Index           =   4
            Left            =   5265
            TabIndex        =   252
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Sum Insured"
            Height          =   495
            Index           =   3
            Left            =   4290
            TabIndex        =   251
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Policy No"
            Height          =   195
            Index           =   2
            Left            =   3150
            TabIndex        =   250
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "Insurer"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   249
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Comment"
            Height          =   195
            Index           =   9
            Left            =   10305
            TabIndex        =   248
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Attach."
            Height          =   195
            Index           =   10
            Left            =   11620
            TabIndex        =   247
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Insurance Type"
            Height          =   195
            Index           =   1
            Left            =   1575
            TabIndex        =   246
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Tel"
            Height          =   195
            Index           =   7
            Left            =   7965
            TabIndex        =   245
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Usage"
            Height          =   195
            Index           =   8
            Left            =   8970
            TabIndex        =   244
            Top             =   315
            Width           =   735
         End
      End
      Begin VB.TextBox txtUnitUtilityCom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -64240
         TabIndex        =   54
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtUnitUtilityStDt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68640
         TabIndex        =   49
         Top             =   720
         Width           =   945
      End
      Begin VB.CommandButton cmdUStatus 
         Caption         =   "..."
         Height          =   315
         Left            =   -68840
         TabIndex        =   48
         Top             =   720
         Width           =   220
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   -74760
         TabIndex        =   220
         Top             =   3720
         Width           =   11835
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   315
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   223
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   315
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   222
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   315
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   221
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   224
            Top             =   240
            Width           =   4890
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "8625;503"
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.CommandButton cmdAttachment 
         Caption         =   "::"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -63660
         TabIndex        =   40
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdInspectedBy 
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
         Height          =   315
         Left            =   -66840
         TabIndex        =   35
         Top             =   840
         Width           =   215
      End
      Begin VB.CheckBox chkAlarm 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -64340
         TabIndex        =   38
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -65500
         TabIndex        =   37
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txtDateChk 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70125
         TabIndex        =   32
         Top             =   840
         Width           =   990
      End
      Begin VB.Frame Frame1 
         Caption         =   "Landlord Contact Details"
         Height          =   2175
         Index           =   0
         Left            =   -69120
         TabIndex        =   205
         Top             =   2520
         Width           =   5895
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Landlord"
            Height          =   195
            Left            =   180
            TabIndex        =   216
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Mobile:"
            Height          =   195
            Left            =   180
            TabIndex        =   215
            Top             =   1500
            Width           =   525
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Telephone:"
            Height          =   195
            Left            =   180
            TabIndex        =   214
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label lblClientName 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1080
            TabIndex        =   213
            Top             =   300
            Width           =   4005
         End
         Begin VB.Label lblClientTelephone 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1080
            TabIndex        =   212
            Top             =   1200
            Width           =   4005
         End
         Begin VB.Label lblClientMobile 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   211
            Top             =   1500
            Width           =   4005
         End
         Begin VB.Label lblClientEmail 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1080
            TabIndex        =   210
            Top             =   1800
            Width           =   4005
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            Height          =   195
            Left            =   180
            TabIndex        =   209
            Top             =   1800
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Address:"
            Height          =   195
            Left            =   180
            TabIndex        =   208
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblClientAddress1 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1080
            TabIndex        =   207
            Top             =   600
            Width           =   4005
         End
         Begin VB.Label lblClientAddress2 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1080
            TabIndex        =   206
            Top             =   900
            Width           =   4005
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tenant Contact Details"
         Height          =   2055
         Left            =   -69120
         TabIndex        =   192
         Top             =   420
         Width           =   5895
         Begin VB.Label Label5 
            Caption         =   "Email:"
            Height          =   255
            Left            =   180
            TabIndex        =   204
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label lblTenantEmail 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1740
            TabIndex        =   203
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label lblTenantMobile 
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1740
            TabIndex        =   202
            Top             =   1380
            Width           =   3375
         End
         Begin VB.Label lblTenantDirectLine 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1740
            TabIndex        =   201
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label lblTenantContact 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1740
            TabIndex        =   200
            Top             =   780
            Width           =   3375
         End
         Begin VB.Label lblTenantSageAcc 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1740
            TabIndex        =   199
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label lblTenantCompany 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1740
            TabIndex        =   198
            Top             =   180
            Width           =   3375
         End
         Begin VB.Label Label18 
            Caption         =   "Direct Line:"
            Height          =   255
            Left            =   180
            TabIndex        =   197
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label Label17 
            Caption         =   "Mobile:"
            Height          =   255
            Left            =   180
            TabIndex        =   196
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Contact:"
            Height          =   255
            Left            =   180
            TabIndex        =   195
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label15 
            Caption         =   "Lessee:"
            Height          =   255
            Left            =   180
            TabIndex        =   194
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label14 
            Caption         =   "Company Name:"
            Height          =   255
            Left            =   180
            TabIndex        =   193
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Unit Maintenance"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   174
         Top             =   360
         Width           =   12225
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   268
            Top             =   3900
            Width           =   2775
            Begin VB.OptionButton optDiary 
               Caption         =   "Diary Entries"
               Height          =   255
               Left            =   1440
               TabIndex        =   269
               Top             =   160
               Width           =   1215
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "Jobs"
               Height          =   255
               Left            =   720
               TabIndex        =   270
               Top             =   160
               Width           =   735
            End
            Begin VB.OptionButton optAll 
               Caption         =   "All"
               Height          =   255
               Left            =   120
               TabIndex        =   271
               Top             =   160
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdEmailJS_PO 
            Caption         =   "Email"
            Height          =   355
            Left            =   9000
            TabIndex        =   178
            Top             =   4035
            Width           =   1395
         End
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "View &Diary Entry"
            Height          =   355
            Left            =   4680
            TabIndex        =   176
            Top             =   4035
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   10680
            TabIndex        =   179
            Top             =   4035
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   6840
            TabIndex        =   177
            Top             =   4035
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "View &Job"
            Height          =   355
            Left            =   3120
            TabIndex        =   175
            Top             =   4035
            Width           =   1395
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   3200
            Left            =   120
            TabIndex        =   180
            Top             =   690
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   5636
            _Version        =   393216
            Cols            =   10
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
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
            _Band(0).Cols   =   10
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   10800
            TabIndex        =   191
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   9840
            TabIndex        =   190
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   3240
            TabIndex        =   189
            Top             =   255
            Width           =   1275
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Reported"
            Height          =   435
            Index           =   2
            Left            =   2385
            TabIndex        =   188
            Top             =   255
            Width           =   810
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   187
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   186
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9240
            TabIndex        =   185
            Top             =   255
            Width           =   435
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   8400
            TabIndex        =   184
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Name / Diary Entry"
            Height          =   495
            Index           =   4
            Left            =   4800
            TabIndex        =   183
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   6000
            TabIndex        =   182
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   7200
            TabIndex        =   181
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.TextBox txtDateVacated 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67705
         TabIndex        =   50
         Top             =   720
         Width           =   945
      End
      Begin VB.TextBox txtFinalReading 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -65120
         TabIndex        =   53
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtUnitMemo 
         Height          =   3135
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   143
         Top             =   480
         Width           =   11835
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit"
         Height          =   355
         Left            =   -65940
         TabIndex        =   142
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save"
         Height          =   355
         Left            =   -64920
         TabIndex        =   141
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
         Height          =   355
         Left            =   -63900
         TabIndex        =   140
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtUnitUtilityIniReading 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66020
         TabIndex        =   52
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtUnitUtilitiesID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -63240
         TabIndex        =   127
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUtilitiesNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -67200
         TabIndex        =   126
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -66165
         TabIndex        =   125
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -64095
         TabIndex        =   124
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -65160
         TabIndex        =   55
         Top             =   4500
         Width           =   975
      End
      Begin VB.TextBox txtUtilitiesReference 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71080
         TabIndex        =   46
         Top             =   720
         Width           =   1050
      End
      Begin VB.TextBox txtChargeRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66800
         TabIndex        =   51
         Top             =   720
         Width           =   780
      End
      Begin VB.Frame Frame4 
         Caption         =   "Payables"
         Height          =   1455
         Left            =   -74640
         TabIndex        =   115
         Top             =   3120
         Width           =   4935
         Begin VB.Label lblLeaseSCPayable 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   123
            Top             =   525
            Width           =   2535
         End
         Begin VB.Label Label69 
            Caption         =   "SC Payable:"
            Height          =   255
            Left            =   240
            TabIndex        =   122
            Top             =   525
            Width           =   1275
         End
         Begin VB.Label lblLeaseRentPayable 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   121
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label71 
            Caption         =   "Rent Payable:"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblLeaseRentReview 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   119
            Top             =   795
            Width           =   2535
         End
         Begin VB.Label Label75 
            Caption         =   "Rent Review Date:"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   795
            Width           =   1455
         End
         Begin VB.Label lblLeaseReviewFreq 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   117
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label77 
            Caption         =   "Review Frequency:"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Lease Information"
         Height          =   915
         Left            =   -74640
         TabIndex        =   99
         Top             =   540
         Width           =   4935
         Begin VB.Label Label20 
            Caption         =   "Lease Reference:"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   103
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblLeaseReference 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1740
            TabIndex        =   102
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label58 
            Caption         =   "Property:"
            Height          =   255
            Left            =   180
            TabIndex        =   101
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblLeaseProperty 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1500
            TabIndex        =   100
            Top             =   600
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lease Details"
         Height          =   1635
         Left            =   -74640
         TabIndex        =   104
         Top             =   1440
         Width           =   4935
         Begin VB.Label lblLeaseType 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   114
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label60 
            Caption         =   "Tenancy Type:"
            Height          =   255
            Left            =   180
            TabIndex        =   113
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label73 
            Caption         =   "Holding Over:"
            Height          =   255
            Left            =   180
            TabIndex        =   112
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label lblLeaseHoldingOver 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   111
            Top             =   780
            Width           =   2535
         End
         Begin VB.Label Label67 
            Caption         =   "Insurance:"
            Height          =   255
            Left            =   180
            TabIndex        =   110
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label lblLeaseInsurance 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   109
            Top             =   510
            Width           =   2535
         End
         Begin VB.Label Label65 
            Caption         =   "Start Date:"
            Height          =   255
            Left            =   180
            TabIndex        =   108
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label lblLeaseStartDate 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   107
            Top             =   1050
            Width           =   2535
         End
         Begin VB.Label Label63 
            Caption         =   "Expiry Date:"
            Height          =   255
            Left            =   180
            TabIndex        =   106
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label lblLeaseExpiryDate 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   105
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdInsuranceNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -66420
         TabIndex        =   58
         Top             =   4515
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -65505
         TabIndex        =   59
         Top             =   4515
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -63675
         TabIndex        =   57
         Top             =   4515
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -64590
         TabIndex        =   56
         Top             =   4515
         Width           =   855
      End
      Begin VB.TextBox txtPropertyInsuranceID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67920
         TabIndex        =   97
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSafetyTelephone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66625
         TabIndex        =   36
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txtNextDueDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -69155
         TabIndex        =   33
         Top             =   840
         Width           =   990
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71655
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   840
         Width           =   1545
      End
      Begin VB.CommandButton cmdSafety 
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
         Height          =   315
         Left            =   -73080
         TabIndex        =   29
         Top             =   840
         Width           =   215
      End
      Begin VB.CommandButton cmdSafetySave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -65130
         TabIndex        =   41
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton cmdSafetyCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -63975
         TabIndex        =   42
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton cmdSafetyEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -66285
         TabIndex        =   89
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton cmdSafetyNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -67440
         TabIndex        =   88
         Top             =   4545
         Width           =   975
      End
      Begin VB.CheckBox chkCertificate 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -64000
         TabIndex        =   39
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox txtUnitSafetyID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66120
         TabIndex        =   87
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridACHistory 
         Height          =   4035
         Left            =   -74820
         TabIndex        =   86
         Top             =   780
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7117
         _Version        =   393216
         ForeColor       =   -2147483641
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridSafety 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   90
         Top             =   1185
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSDataListLib.DataCombo cboSafetyType 
         Bindings        =   "frmUnits2.frx":09FC
         DataSource      =   "adoSafetyType"
         Height          =   315
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUtilities 
         Height          =   3315
         Left            =   -74955
         TabIndex        =   128
         Top             =   1140
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSDataListLib.DataCombo cboUtilitiesType 
         Bindings        =   "frmUnits2.frx":0A18
         DataSource      =   "adoUtilitiesType"
         Height          =   315
         Left            =   -74160
         TabIndex        =   43
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdSetUtilitiesType 
         Caption         =   "..."
         Height          =   315
         Left            =   -72940
         TabIndex        =   44
         Top             =   720
         Width           =   220
      End
      Begin VB.Frame Frame8 
         Caption         =   "Unit Analysis Information"
         Height          =   4140
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   11895
         Begin VB.TextBox txtUnitAnalysisID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4800
            TabIndex        =   84
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdAnalysisSave 
            Caption         =   "&Save"
            Height          =   315
            Left            =   9690
            TabIndex        =   83
            Top             =   3780
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisCancel 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10725
            TabIndex        =   82
            Top             =   3780
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   8655
            TabIndex        =   81
            Top             =   3780
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   7620
            TabIndex        =   80
            Top             =   3780
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysis 
            Caption         =   "..."
            Height          =   315
            Left            =   2340
            TabIndex        =   22
            Top             =   420
            Width           =   255
         End
         Begin VB.TextBox txtAnalysisPercentage 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   10440
            TabIndex        =   27
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox txtAnalysisQuantity 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   9180
            TabIndex        =   26
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox txtAnalysisValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   7920
            TabIndex        =   25
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox txtAnalysisDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2580
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   420
            Width           =   4035
         End
         Begin VB.TextBox txtAnalysisTotalArea 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   72
            Top             =   3780
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo cboAnalysisType 
            Bindings        =   "frmUnits2.frx":0A37
            DataSource      =   "adoAnalysisType"
            Height          =   315
            Left            =   180
            TabIndex        =   21
            Top             =   420
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboAnalysisOption 
            Bindings        =   "frmUnits2.frx":0A55
            DataSource      =   "adoSelectOption"
            Height          =   315
            Left            =   6600
            TabIndex        =   24
            Top             =   420
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUnitAnalysis 
            Height          =   2955
            Left            =   180
            TabIndex        =   85
            Top             =   780
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   5212
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label36 
            Caption         =   "Percentage: (%)"
            Height          =   255
            Left            =   10440
            TabIndex        =   79
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label35 
            Caption         =   "Quantity:"
            Height          =   255
            Left            =   9180
            TabIndex        =   78
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "Value(sq.ft/sq.m):"
            Height          =   195
            Left            =   7920
            TabIndex        =   77
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Select Option:"
            Height          =   255
            Left            =   6600
            TabIndex        =   76
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Analysis Type:"
            Height          =   255
            Left            =   180
            TabIndex        =   75
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Description:"
            Height          =   255
            Left            =   2580
            TabIndex        =   74
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Total Area:"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   3840
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin MSDataListLib.DataCombo cboAuthority_Supplier 
         Bindings        =   "frmUnits2.frx":0A73
         DataSource      =   "adoSupplier"
         Height          =   315
         Left            =   -72745
         TabIndex        =   45
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "SupplierName"
         BoundColumn     =   "SupplierID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboSchedule 
         Bindings        =   "frmUnits2.frx":0A8D
         DataSource      =   "adoSafetyStatus"
         Height          =   315
         Left            =   -72855
         TabIndex        =   30
         Top             =   840
         Width           =   1235
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboInspectedBy 
         Bindings        =   "frmUnits2.frx":0AAB
         DataSource      =   "adoInspector"
         Height          =   315
         Left            =   -68190
         TabIndex        =   34
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboUnitUtilityStatus 
         Bindings        =   "frmUnits2.frx":0AC6
         DataSource      =   "adoUStatus"
         Height          =   315
         Left            =   -70045
         TabIndex        =   47
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridInsurance 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   255
         Top             =   1440
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483640
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label82 
         Caption         =   "Comments"
         Height          =   195
         Index           =   10
         Left            =   -64240
         TabIndex        =   228
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label82 
         Caption         =   "Start Date"
         Height          =   435
         Index           =   5
         Left            =   -68600
         TabIndex        =   227
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Index           =   4
         Left            =   -70005
         TabIndex        =   226
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label82 
         Caption         =   "Occupier ID"
         Height          =   435
         Index           =   0
         Left            =   -74955
         TabIndex        =   225
         Top             =   360
         Width           =   645
      End
      Begin VB.Image VerticalLabel 
         Height          =   600
         Index           =   2
         Left            =   -63660
         Picture         =   "frmUnits2.frx":0ADF
         Top             =   210
         Width           =   210
      End
      Begin VB.Image VerticalLabel 
         Height          =   435
         Index           =   1
         Left            =   -64005
         Picture         =   "frmUnits2.frx":0EE9
         Top             =   360
         Width           =   225
      End
      Begin VB.Image VerticalLabel 
         Height          =   585
         Index           =   0
         Left            =   -64350
         Picture         =   "frmUnits2.frx":127D
         Top             =   225
         Width           =   195
      End
      Begin VB.Label Label41 
         Caption         =   "Contact / Telephone"
         Height          =   375
         Index           =   6
         Left            =   -66625
         TabIndex        =   219
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Date Checked"
         Height          =   345
         Index           =   3
         Left            =   -70125
         TabIndex        =   218
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label41 
         Caption         =   "Schedule"
         Height          =   255
         Index           =   1
         Left            =   -72855
         TabIndex        =   217
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblMainUnit 
         Caption         =   "Usage"
         Height          =   255
         Index           =   18
         Left            =   -67560
         TabIndex        =   173
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblMainUnit 
         Caption         =   "End Date"
         Height          =   255
         Index           =   17
         Left            =   -68880
         TabIndex        =   172
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         Caption         =   "Start Date"
         Height          =   255
         Index           =   16
         Left            =   -70320
         TabIndex        =   171
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         Caption         =   "Tenant Name"
         Height          =   255
         Index           =   15
         Left            =   -73320
         TabIndex        =   170
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         Caption         =   "A/C Code"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   169
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label82 
         Caption         =   "End  Date"
         Height          =   435
         Index           =   6
         Left            =   -67665
         TabIndex        =   146
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label82 
         Caption         =   "Final    Reading"
         Height          =   435
         Index           =   9
         Left            =   -65120
         TabIndex        =   145
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label81 
         BackColor       =   &H00B3C0C6&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   -74820
         TabIndex        =   144
         Top             =   5700
         Width           =   11895
      End
      Begin VB.Label lblLeaseHeading 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   -72120
         TabIndex        =   139
         Top             =   360
         Width           =   5940
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Reference"
         Height          =   195
         Index           =   3
         Left            =   -71040
         TabIndex        =   134
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label82 
         Caption         =   "Initial Reading"
         Height          =   435
         Index           =   8
         Left            =   -66000
         TabIndex        =   133
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   1
         Left            =   -74200
         TabIndex        =   132
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label83 
         BackColor       =   &H00B3C0C6&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   -74820
         TabIndex        =   131
         Top             =   5580
         Width           =   11715
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   -72705
         TabIndex        =   130
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Index           =   7
         Left            =   -66740
         TabIndex        =   129
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label46 
         BackColor       =   &H00B3C0C6&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   -74880
         TabIndex        =   98
         Top             =   5580
         Width           =   12075
      End
      Begin VB.Label Label41 
         Caption         =   "Comment"
         Height          =   255
         Index           =   7
         Left            =   -65500
         TabIndex        =   96
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label41 
         Caption         =   "Inspected By"
         Height          =   255
         Index           =   5
         Left            =   -68160
         TabIndex        =   95
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label41 
         Caption         =   "Next Inspection"
         Height          =   375
         Index           =   4
         Left            =   -69155
         TabIndex        =   94
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label41 
         Caption         =   "Type"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   93
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label41 
         Caption         =   "Reference"
         Height          =   255
         Index           =   2
         Left            =   -71655
         TabIndex        =   92
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label43 
         BackColor       =   &H00B3C0C6&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   -74820
         TabIndex        =   91
         Top             =   5580
         Width           =   11955
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   100
      Left            =   -120
      ScaleHeight     =   45
      ScaleWidth      =   12675
      TabIndex        =   155
      Top             =   3405
      Width           =   12735
   End
   Begin VB.Frame Frame5 
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   108
      TabIndex        =   20
      Top             =   0
      Width           =   8625
      Begin VB.CommandButton cmdTenancyType 
         Caption         =   "..."
         Height          =   252
         Left            =   7596
         TabIndex        =   286
         Top             =   1224
         Width           =   324
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   252
         Left            =   7596
         TabIndex        =   285
         Top             =   180
         Width           =   324
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Unit"
         Height          =   315
         Left            =   4596
         TabIndex        =   261
         Top             =   3000
         Width           =   1215
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
         Left            =   3400
         TabIndex        =   0
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton cmdNewUnit 
         Caption         =   "&New Unit"
         Height          =   315
         Left            =   816
         TabIndex        =   262
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelUnit 
         Caption         =   "&Cancel Unit"
         Height          =   315
         Left            =   3372
         TabIndex        =   260
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveUnit 
         Caption         =   "&Save Unit"
         Height          =   315
         Left            =   5832
         TabIndex        =   259
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditUnit 
         Caption         =   "&Edit Unit"
         Height          =   315
         Left            =   2136
         TabIndex        =   258
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseUnit 
         Caption         =   "C&lose"
         Height          =   315
         Left            =   7056
         TabIndex        =   257
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopyUnit 
         Caption         =   "C&opy Unit"
         Height          =   315
         Left            =   0
         TabIndex        =   256
         Top             =   1908
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCurrentUsage 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   167
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox txtRentalPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   165
         Top             =   2595
         Width           =   2415
      End
      Begin VB.TextBox txtUnitAddress4 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         MaxLength       =   70
         TabIndex        =   7
         Top             =   2295
         Width           =   2415
      End
      Begin VB.TextBox txtUnitAddress2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         MaxLength       =   70
         TabIndex        =   5
         Top             =   1620
         Width           =   2415
      End
      Begin VB.TextBox txtUnitAddress3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         MaxLength       =   70
         TabIndex        =   6
         Top             =   1965
         Width           =   2415
      End
      Begin VB.TextBox txtUnitPostCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   2640
         Width           =   915
      End
      Begin VB.TextBox txtUnitAddress1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         MaxLength       =   70
         TabIndex        =   4
         Top             =   1290
         Width           =   2415
      End
      Begin VB.CommandButton cmdUnitLookup 
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
         Height          =   315
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   555
         Width           =   345
      End
      Begin VB.TextBox txtUnitNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   2
         Top             =   555
         Width           =   2430
      End
      Begin VB.TextBox txtTotalArea 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtUnitName 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFE&
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   945
         Width           =   2415
      End
      Begin VB.CommandButton cmdUnitStatus 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7935
         TabIndex        =   12
         Top             =   855
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CommandButton cmdManagement 
         Caption         =   "..."
         Height          =   315
         Left            =   7935
         TabIndex        =   14
         Top             =   1215
         Width           =   400
      End
      Begin VB.CommandButton cmdCurrentTenant 
         Caption         =   "..."
         Height          =   315
         Left            =   7935
         TabIndex        =   16
         Top             =   1575
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CommandButton cmdLandlord 
         Caption         =   "..."
         Height          =   315
         Left            =   7935
         TabIndex        =   18
         Top             =   1920
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.CommandButton cmdUnitType 
         Caption         =   "..."
         Height          =   315
         Left            =   7935
         TabIndex        =   10
         Top             =   150
         Width           =   400
      End
      Begin MSDataListLib.DataCombo cboLandLord 
         Bindings        =   "frmUnits2.frx":1667
         DataSource      =   "adoLandLord"
         Height          =   315
         Left            =   5520
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   15466238
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboCurrentTenant 
         Bindings        =   "frmUnits2.frx":1681
         DataSource      =   "adoCurrentTenant"
         Height          =   315
         Left            =   5520
         TabIndex        =   15
         Top             =   1575
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   15466238
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboStatus 
         Bindings        =   "frmUnits2.frx":16A0
         DataSource      =   "adoStatus"
         Height          =   315
         Left            =   5520
         TabIndex        =   11
         Top             =   855
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   15466238
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboManagement 
         Bindings        =   "frmUnits2.frx":16B8
         DataSource      =   "adoManagement"
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   1215
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   15466238
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboUnitType 
         Bindings        =   "frmUnits2.frx":16D4
         DataSource      =   "adoUnitType"
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Top             =   150
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         BackColor       =   15466238
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   1305
         TabIndex        =   275
         Top             =   180
         Width           =   2430
         VariousPropertyBits=   746604571
         Size            =   "4286;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Usage:"
         Height          =   255
         Index           =   13
         Left            =   4320
         TabIndex        =   168
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Price:"
         Height          =   255
         Index           =   12
         Left            =   4320
         TabIndex        =   166
         Top             =   2595
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area:"
         Height          =   255
         Index           =   10
         Left            =   4320
         TabIndex        =   154
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   136
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Address:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   135
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Tenancy type:"
         Height          =   255
         Index           =   7
         Left            =   4320
         TabIndex        =   67
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name:"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   66
         Top             =   1935
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Tenant:"
         Height          =   255
         Index           =   8
         Left            =   4320
         TabIndex        =   65
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   64
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit No:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   62
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type:"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   61
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label lblMainUnit 
         BackStyle       =   0  'Transparent
         Caption         =   "Post Code:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   2700
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdImgDelete 
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11880
      TabIndex        =   151
      ToolTipText     =   "Delete current image"
      Top             =   2352
      Width           =   555
   End
   Begin MSAdodcLib.Adodc adoLandLord 
      Height          =   330
      Left            =   -45
      Top             =   10680
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Landlord"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoCurrentTenant 
      Height          =   330
      Left            =   10680
      Top             =   9780
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Current Tenant"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoUnitType 
      Height          =   330
      Left            =   -45
      Top             =   10920
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Unit Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Height          =   390
      Left            =   4800
      ScaleHeight     =   390
      ScaleWidth      =   2655
      TabIndex        =   137
      Top             =   4200
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
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
         Height          =   195
         Left            =   120
         TabIndex        =   138
         Top             =   80
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdUploadImageAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11880
      TabIndex        =   69
      ToolTipText     =   "Add new image"
      Top             =   2710
      Width           =   555
   End
   Begin MSAdodcLib.Adodc adoSelectOption 
      Height          =   330
      Left            =   10680
      Top             =   9120
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Select Option"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoAnalysisType 
      Height          =   330
      Left            =   10680
      Top             =   9480
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AnalysisType"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSafetyType 
      Height          =   330
      Left            =   10680
      Top             =   10440
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Safety Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSafetyStatus 
      Height          =   330
      Left            =   10680
      Top             =   10740
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Safety Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsuranceType 
      Height          =   330
      Left            =   6000
      Top             =   12240
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Insurance Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInspector 
      Height          =   330
      Left            =   -45
      Top             =   9360
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Inspector"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoUtilitiesType 
      Height          =   330
      Left            =   -45
      Top             =   11220
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Unitlities Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoProperty 
      Height          =   330
      Left            =   -45
      Top             =   9720
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Property"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoManagement 
      Height          =   330
      Left            =   -45
      Top             =   10380
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Management"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoMType 
      Height          =   330
      Left            =   -45
      Top             =   10080
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Maintenance Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox fmeUnitLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   3825
      Left            =   2640
      ScaleHeight     =   3795
      ScaleWidth      =   7905
      TabIndex        =   147
      Top             =   8475
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6615
         TabIndex        =   274
         Top             =   225
         Width           =   930
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   273
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4905
         TabIndex        =   272
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txtSearchAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2700
         TabIndex        =   163
         Top             =   240
         Width           =   2190
      End
      Begin VB.TextBox txtSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   162
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox txtSearchUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   50
         TabIndex        =   150
         Top             =   240
         Width           =   1020
      End
      Begin VB.CommandButton cmdGridUnitLookup 
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUnitLookup 
         Height          =   3195
         Left            =   45
         TabIndex        =   148
         Top             =   555
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   13553358
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   161
         Top             =   30
         Width           =   285
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1140
         TabIndex        =   160
         Top             =   30
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2745
         TabIndex        =   159
         Top             =   30
         Width           =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PostCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   4890
         TabIndex        =   158
         Top             =   30
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   5775
         TabIndex        =   157
         Top             =   30
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6660
         TabIndex        =   156
         Top             =   30
         Width           =   450
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   50
         Top             =   30
         Width           =   7575
      End
   End
   Begin MSAdodcLib.Adodc adoStatus 
      Height          =   330
      Left            =   10680
      Top             =   10080
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSupplier 
      Height          =   330
      Left            =   -45
      Top             =   11520
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Suppliers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsurer 
      Height          =   330
      Left            =   6000
      Top             =   11760
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Insurer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsUsage 
      Height          =   330
      Left            =   6000
      Top             =   12000
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Usage"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoUStatus 
      Height          =   330
      Left            =   10800
      Top             =   11520
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   255
      Left            =   8880
      TabIndex        =   267
      ToolTipText     =   "Next image"
      Top             =   3120
      Width           =   555
      PicturePosition =   393216
      Size            =   "979;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblMainUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Unit Code)"
      Height          =   255
      Index           =   11
      Left            =   10800
      TabIndex        =   164
      Top             =   11280
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSForms.CommandButton cmdImgLeftMove 
      Height          =   255
      Left            =   11880
      TabIndex        =   153
      ToolTipText     =   "Next image"
      Top             =   2040
      Width           =   555
      PicturePosition =   393216
      Size            =   "979;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label lblImageName 
      Height          =   195
      Left            =   8760
      TabIndex        =   152
      Top             =   0
      Width           =   3120
      VariousPropertyBits=   8388627
      Caption         =   "Image Name:"
      Size            =   "5503;344"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image imgUnitPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2771
      Left            =   8760
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3045
   End
   Begin VB.Label Label31 
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   68
      Top             =   2445
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmUnits2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LOAD_UNIT_UNITID As String
Public HEALTH_N_SAFETY_ATTACH As Boolean

Dim NEWMODE_ As Boolean
Dim SEARCHUNITMODE_ As Boolean
'' UNIT DETAILS ENTRY FLAG
Dim M_HISTORY_NEW_ENTRY_ As Boolean
Dim UNIT_ANALYSIS_NEW_ENTRY As Boolean
Dim UNIT_INSURANCE_NEW_ENTRY As Boolean
Private INSURANCE_ID As String
Dim UNIT_UTILITIES_NEW_ENTRY As Boolean
Dim IMAGE_FILE_NAME_ As String
'Private HEALTH_SAFETY_ID As String
Dim HEALTH_SAFETY_NEW_ENTRY As Boolean
Dim bSortingCol1 As Boolean, bSortingCol2 As Boolean, bSortingCol3 As Boolean

Dim DSN_ALARM_ As String
Dim lblAsJS_PO As String
Dim sTextBox As String
'''''''''''''''''''Modified by Mahboob Change ID 3/Work Item 2 :- Declare variable to separate unit type and tenancy type
Dim isUnitType As Integer
''''''''''''''''''''''''''End of modification







Private Sub cboLandLord_KeyPress(KeyAscii As Integer)
    
     If KeyAscii = 13 Then
        FocusControl txtTotalArea
    End If
End Sub

Private Sub cboManagement_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cboLandLord
    End If
End Sub

Private Sub cboUnitType_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        FocusControl txtCurrentUsage
    End If
End Sub

Private Sub cmdDelete_Click()
    If txtUnitNo.text = "" Then
    
    ''''''''''''''''''Modified By Mahboob 13/03/2023 Change ID:1/Work Item:6 :- set message box while unit number is empty
       MsgBox "Please select a unit to delete.", vbInformation, "Select Unit Number"
       cmdUnitLookup.SetFocus
       Exit Sub
       ''''''''''''''End of modification
    
      'ShowMsgInTaskBar "Please select a Unit No to continue."
      
   End If
   'Check the existence of Unit under it or any transaction is made
   Dim adoConn As New ADODB.Connection
   Dim rsTransaction As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsTransaction.Open "Select UnitNumber from Units where UnitNumber='" & txtUnitNo.text & "'", adoConn, adOpenKeyset
   If rsTransaction.EOF Then
        MsgBox "This Unit number was not found in the database", vbInformation, "Not found"
        FocusControl cmdUnitLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   rsTransaction.Open "Select UnitID from tlbReceipt where unitID='" & txtUnitNo.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Unit cannot be deleted. Because there is some receipt reference with this Unit ID", vbInformation, "Cannot Delete"
        cmdCancelUnit_Click
        FocusControl cmdUnitLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   rsTransaction.Open "Select UNIT_ID from tlbBankPayment where UNIT_ID='" & txtUnitNo.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Unit cannot be deleted. Because there is some bank receipt reference with this Unit ID", vbInformation, "Cannot Delete"
        cmdCancelUnit_Click
        FocusControl cmdUnitLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   'Delete Unit
    If MsgBox("  Are you sure you wish to delete this unit information?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Delete unit information") = vbYes Then
            adoConn.Execute "DELETE FROM Units where Unitnumber='" & txtUnitNo.text & "'"
            MsgBox "Delete Successful", vbInformation
   End If
   adoConn.Close
   cmdCancelUnit_Click
   FocusControl cmdUnitLookup
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
End Sub



Private Sub cmdTenancyType_Click()
'''''''''''''''''Modification by Md. Mahboob 20230208 -:Change ID 6/Work Item 1  To loading picture box for Unit type
 ''''''''''''''''' isUnitType = 2 for search tenacy type
 isUnitType = 2
    Picture2.Left = 4000
    Picture2.Top = 400
    lblCode.Caption = "Tenancy type Code"
    lblName.Caption = "Tenancy type Name"
    Picture2.Visible = True
    LoadUniTypeListList
    Picture2.Enabled = True
    txtSearchUnitTypeCode.SetFocus
End Sub

Private Sub Command1_Click()
'''''''''''''''''Modification by Md. Mahboob 20230208 Change ID 3/Work Item 11 -:  To close the picture box
Picture2.Visible = False
End Sub

Private Sub Command2_Click()
'''''''''''''''''Modification by Md. Mahboob 20230208 -:Change ID 3/Work Item 4  To loading picture box for Unit type
 ''''''''''''''''' isUnitType = 1 for search unit type
 isUnitType = 1
    Picture2.Left = 4000
    Picture2.Top = 200
    lblCode.Caption = "Unit Type Code"
    lblName.Caption = "Unit Type Name"
    Picture2.Visible = True
    LoadUniTypeListList
    Picture2.Enabled = True
    txtSearchUnitTypeCode.SetFocus
    ''''''''''''''''''''End of modification
End Sub
'''''''''''''''''Modification by Md. Mahboob 20230208 -:Change ID 3/Work Item 5  To load data in gridview of unit type
Private Sub LoadUniTypeListList()
    
   Dim rRow As Integer
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchUnitTypeCode.text = ""
   txtUnitTypeName.text = ""
   flxUnitType.RowHeight(0) = 0
   flxUnitType.Cols = 3
   flxUnitType.ColWidth(0) = 80
   flxUnitType.ColWidth(1) = 1500
   flxUnitType.ColWidth(2) = 4500
   flxUnitType.Clear
   'flxUnitType.Rows = 5
   flxUnitType.ColAlignment(0) = vbLeftJustify
   flxUnitType.ColAlignment(1) = vbLeftJustify
   flxUnitType.ColAlignment(2) = vbLeftJustify
   txtSearchUnitTypeCode.Width = 1530
   txtUnitTypeName.Visible = True
   txtSearchUnitTypeCode.Left = 45
   txtUnitTypeName.Left = 1620
   txtUnitTypeName.text = ""
   txtSearchClientID.text = ""
   txtSearchUnitTypeCode.Left = 45
   adoConn.Open getConnectionString
   If isUnitType = 1 Then
   szSQL = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTYP'"
                 Else
    szSQL = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'TNTYPE'"
                 End If
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
            flxUnitType.Rows = rstRec.RecordCount + 1
        While Not rstRec.EOF
           flxUnitType.row = 1
           flxUnitType.RowSel = 1
           flxUnitType.ColSel = 1
           flxUnitType.TextMatrix(rRow, 0) = ""
           flxUnitType.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxUnitType.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxUnitType.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxUnitType.AddItem ""
           rRow = rRow + 1
        Wend
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub flxUnitType_Click()
'''''''''''''''''Modification by Md. Mahboob 20230208 -: Change ID 3/Work Item 6 To loading picture box for Unit type
 ''''''''''''''''' isUnitType = 1 for search unit type
If isUnitType = 1 Then
 cboUnitType.BoundText = flxUnitType.TextMatrix(flxUnitType.row, 1)
    cboUnitType.text = flxUnitType.TextMatrix(flxUnitType.row, 2)
    Else
    cboManagement.BoundText = flxUnitType.TextMatrix(flxUnitType.row, 1)
    cboManagement.text = flxUnitType.TextMatrix(flxUnitType.row, 2)
    End If
    Picture2.Visible = False
End Sub

''''''''''''''''''''''End of modification
Private Sub txtSearchUnitTypeCode_Change()
'''''''''''''''''Modification by Md. Mahboob 20230218 -:Change ID 3/Work Item 7  To searching unit code by typing
   Dim i As Integer

   If Len(txtSearchUnitTypeCode.text) > 0 Then
        txtUnitTypeName.text = ""
   End If

   For i = flxUnitType.Rows - 1 To 1 Step -1
        flxUnitType.RowHeight(i) = 240
        If InStr(1, UCase(flxUnitType.TextMatrix(i, 1)), UCase(txtSearchUnitTypeCode.text), vbTextCompare) = 0 Then
              flxUnitType.RowHeight(i) = 0
        End If
        If flxUnitType.RowHeight(i) = 240 Then
              flxUnitType.row = i
        End If
   Next i
End Sub
'''''''''''''End of modification
Private Sub txtSearchUnitTypeCode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'''''''''''''''''Modification by Md. Mahboob 20230218 -: Change ID 3/Work Item 8 To moving cursor unit type code to unit name on enter and flix grid on down arrow
If KeyCode = vbKeyDown Then
           flxUnitType.SetFocus
    End If
    If KeyCode = 13 Then
           txtUnitTypeName.SetFocus
    End If
End Sub
'''''''''''''End of modification
Private Sub txtSearchUnitTypeCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
''''''''''''''''''''''''Modified by Mahboob 20230208 Change ID 3/Work Item 9 visuable false picture2 and command2 set focuse
If KeyAscii = 27 Then
            Picture2.Visible = False
           Command2.SetFocus
    End If
End Sub
'''''''''''''End of modification
Private Sub txtUnitTypeName_Change()
'''''''''''''''''Modification by Md. Mahboob 20230218 -: Change ID 3/Work Item 10 To searching unit name by typing
   Dim i As Integer

   If Len(txtUnitTypeName.text) > 0 Then
        txtSearchUnitTypeCode.text = ""
   End If

   For i = flxUnitType.Rows - 1 To 1 Step -1
        flxUnitType.RowHeight(i) = 240
        If InStr(1, UCase(flxUnitType.TextMatrix(i, 2)), UCase(txtUnitTypeName.text), vbTextCompare) = 0 Then
              flxUnitType.RowHeight(i) = 0
        End If
        If flxUnitType.RowHeight(i) = 240 Then
              flxUnitType.row = i
        End If
   Next i

End Sub
'''''''''''''End of modification

Private Sub txtUnitTypeName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'''''''''''''''''Modification by Md. Mahboob 20230218 -: Change ID 3/Work Item 11 To focus flexgrid on enter press
If KeyCode = 13 Then
         flxUnitType.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxUnitType.Visible Then
            flxUnitType.SetFocus
        End If
    End If
End Sub
'''''''''''''End of modification
Private Sub txtCurrentUsage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cboManagement
    End If
End Sub

Private Sub txtRentalPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSaveUnit
    End If
     DigitTextKeyPress txtRentalPrice, KeyAscii
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
          
         
          'If sTextBox = "1" Then
           cmdProperty.SetFocus
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
            If sTextBox = "2" Then
                    txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                    If cmdUnitLookup.Enabled And NEWMODE_ = False Then
                        cmdUnitLookup.SetFocus
                    Else
                       If txtUnitNo.Enabled Then txtUnitNo.SetFocus
                    End If
                    cboProperty_Change
            End If
            picClient.Visible = False
        
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
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
   'lblClientName.Caption = "Property Name"
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
           
        szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
                "FROM PROPERTY " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = "ALL"
'           flxClient.TextMatrix(rRow, 2) = "ALL"
'           flxClient.RowHeight(rRow) = 240
'           flxClient.AddItem ""
'           rRow = 2
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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Function LoadGridUnitLookup(ByVal strFilter_ As String)
  'cmdClientID.Default = True
   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
   conUnit_.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE, UNITTYPE, " & _
           "Occupied,P.PropertyName,P.PropertyID " & _
           "FROM UNITS,Property P where P.PropertyID=Units.PropertyID " & strFilter_

'Debug.Print sSQLQuery_
   rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   gridUnitLookup.Clear
   
   'Resolved by BOSL
   'Issue No: 0000445, 00000442
   'Modified By: Asif. 02 Aug 2014
   'gridUnitLookup.Rows = 2
   gridUnitLookup.Rows = rstUnit_.RecordCount + 1
   
  
   ConfigGridUnitLookup
   While Not rstUnit_.EOF
        gridUnitLookup.TextMatrix(iRow, 0) = rstUnit_!UnitNumber
        gridUnitLookup.TextMatrix(iRow, 1) = rstUnit_!UnitName
        gridUnitLookup.TextMatrix(iRow, 2) = IIf(IsNull(rstUnit_!Address), "", rstUnit_!Address)
        gridUnitLookup.TextMatrix(iRow, 3) = IIf(IsNull(rstUnit_!UnitPostCode), "", rstUnit_!UnitPostCode)
        gridUnitLookup.TextMatrix(iRow, 4) = IIf(IsNull(rstUnit_!UNITTYPE), "", rstUnit_!UNITTYPE)
        gridUnitLookup.TextMatrix(iRow, 5) = IIf(rstUnit_!OCCUPIED = "N", "Vacant", "Occupied")
        'added by anol 23 Aug 2016
        gridUnitLookup.TextMatrix(iRow, 6) = IIf(IsNull(rstUnit_!PropertyName), "", rstUnit_!PropertyName)
        gridUnitLookup.TextMatrix(iRow, 7) = IIf(IsNull(rstUnit_!propertyID), "", rstUnit_!propertyID)
        
        rstUnit_.MoveNext
'      If Not rstUnit_.EOF Then gridUnitLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstUnit_.Close
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing

   'cmdSelected.Enabled = True
End Function


'Resolved by BOSL
'Issue No: 0000445, 00000442
'The function generates the expression of matching string pattern by using SQL LIKE operation and
'uses the in-built Filter function of the ADODB recordset to filter the records that match with the
'expression and finally bind the filtered records to the grid.
'Modified By: Asif. 02 Aug 2014
Private Function FilterUnitsList() As String
   
   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the RDO Connections to the dataset
   conUnit_.Open getConnectionString

   Dim strFilterByProperty As String
   
   If txtPropertyName.text = "" Then
      strFilterByProperty = ""
   Else
      strFilterByProperty = "AND (((UNITS.PROPERTYID) = '" & txtPropertyName.Tag & "'));"
   End If
      
   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'   sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
'           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE, UNITTYPE, " & _
'           "Occupied " & _
'           "FROM UNITS " & strFilterByProperty
' Modified by anol 23 aug 2016
           
sSQLQuery_ = "SELECT UNITNUMBER, UNITNAME, UNITADDRESSLINE1 + ' ' + UNITADDRESSLINE2 + ' ' +  " & _
           "UNITADDRESSLINE3 + ' ' + UNITADDRESSLINE4 as Address, UNITPOSTCODE, UNITTYPE, " & _
           "Occupied,P.PropertyName,P.PropertyID " & _
           "FROM UNITS,Property P where P.PropertyID=Units.PropertyID " & strFilterByProperty

 Dim Filter As String
   
   If Len(txtSearchUnit.text) > 0 Then
      txtSearchName.text = ""
      txtSearchAddress.text = ""
      Filter = " UNITNUMBER LIKE '*" + UCase(txtSearchUnit.text) + "*'"
      
   End If
   
   If Len(txtSearchName.text) > 0 Then
      txtSearchUnit.text = ""
      txtSearchAddress.text = ""
      Filter = " UNITNAME LIKE '%" + UCase(txtSearchName.text) + "*'"
   End If

   If Len(txtSearchAddress.text) > 0 Then
      txtSearchUnit.text = ""
      txtSearchName.text = ""
      Filter = " Address LIKE '*" + UCase(txtSearchAddress.text) + "*'"
   End If
   
'Debug.Print sSQLQuery_
   rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

   rstUnit_.Filter = Filter
   
'   MsgBox adoRst.RecordCount
      
   gridUnitLookup.Clear
   
   gridUnitLookup.Rows = rstUnit_.RecordCount + 1
   Dim iRow As Integer
   iRow = 1

'   gridUnitLookup.Cols = 6
   ConfigGridUnitLookup
   
   While Not rstUnit_.EOF
      gridUnitLookup.TextMatrix(iRow, 0) = rstUnit_!UnitNumber
      gridUnitLookup.TextMatrix(iRow, 1) = rstUnit_!UnitName
      gridUnitLookup.TextMatrix(iRow, 2) = IIf(IsNull(rstUnit_!Address), "", rstUnit_!Address)
      gridUnitLookup.TextMatrix(iRow, 3) = IIf(IsNull(rstUnit_!UnitPostCode), "", rstUnit_!UnitPostCode)
      gridUnitLookup.TextMatrix(iRow, 4) = IIf(IsNull(rstUnit_!UNITTYPE), "", rstUnit_!UNITTYPE)
      gridUnitLookup.TextMatrix(iRow, 5) = IIf(rstUnit_!OCCUPIED = "N", "Vacant", "Occupied")
      
      'added by anol 23 Aug 2016
        gridUnitLookup.TextMatrix(iRow, 6) = IIf(IsNull(rstUnit_!PropertyName), "", rstUnit_!PropertyName)
        gridUnitLookup.TextMatrix(iRow, 7) = IIf(IsNull(rstUnit_!propertyID), "", rstUnit_!propertyID)
      
      rstUnit_.MoveNext
'      If Not rstUnit_.EOF Then gridUnitLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstUnit_.Close
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing

   'cmdSelected.Enabled = True
End Function

Private Sub ConfigGridUnitLookup()
   fmeUnitLookup.Visible = True
   gridUnitLookup.Visible = True
   gridUnitLookup.RowHeight(0) = 0
   gridUnitLookup.Cols = 8
   
   gridUnitLookup.ColWidth(0) = 1100
   gridUnitLookup.TextMatrix(0, 0) = "Unit Number"
   gridUnitLookup.ColAlignment(0) = vbLeftJustify

   gridUnitLookup.ColWidth(1) = 1500
   gridUnitLookup.TextMatrix(0, 1) = "Name"
   gridUnitLookup.ColAlignment(1) = vbLeftJustify

   gridUnitLookup.ColWidth(2) = 2200
   gridUnitLookup.TextMatrix(0, 2) = "Address"
   gridUnitLookup.ColAlignment(2) = vbLeftJustify

   gridUnitLookup.ColWidth(3) = 800
   gridUnitLookup.TextMatrix(0, 3) = "PostCode"
   gridUnitLookup.ColAlignment(3) = vbLeftJustify

   gridUnitLookup.ColWidth(4) = 800
   gridUnitLookup.TextMatrix(0, 4) = "Unit Type"
   gridUnitLookup.ColAlignment(4) = vbLeftJustify

   gridUnitLookup.ColWidth(5) = 800
   gridUnitLookup.TextMatrix(0, 5) = "Status"
   gridUnitLookup.ColAlignment(5) = vbLeftJustify
   
   gridUnitLookup.ColWidth(6) = 0 'Property ID
   gridUnitLookup.ColWidth(7) = 0 'Property Name
End Sub

Private Sub cboProperty_Change()
'   ''''''''''''''''''''Md.Mahboob 2023/03/17 Change ID 2/Work Item 1 Stop generate auto number
'   If NEWMODE_ Then
'      txtUnitNo.text = GenerateUnitNumber
'   End If
   ''''''''''''''''''''End of Modification
'Resolved by BOSL
'Issue No: 0000467
'Modified By: Asif. 04 Sep 2014
   If txtUnitAddress1.Enabled = False And txtUnitName.text <> "" Then
        Dim selectedProperty As String
        selectedProperty = txtPropertyName.Tag
   
        cmdCancelUnit_Click
        txtPropertyName.Tag = selectedProperty
   End If
   
End Sub

Private Sub cmdAddDiary_Click()
   If txtUnitNo.text = "" Then Exit Sub

   With frmMaintananceDairy
      .CallingForm = "U"          'Calling from lessee form
      .isEdit = False
      .RecordType = "D"
      .lblJobName.Caption = "Diary Name"
      .Label1.Caption = "Diary Entry No."
      Load frmMaintananceDairy
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With

   Me.Enabled = False
End Sub

Private Sub cmdAnalysis_Click()
   Dim sSQLQuery_  As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "ATYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoAnalysisType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'ATYP'"

   adoAnalysisType.RecordSource = sSQLQuery_
   adoAnalysisType.CommandType = adCmdText
   adoAnalysisType.Refresh
End Sub

Private Sub cmdAnalysisCancel_Click()
   UnitAnalysisButtonMode DefaultMode
   Frame5.Enabled = True
End Sub

Private Sub cmdAnalysisEdit_Click()
   If txtUnitAnalysisID.text = "" Then
       Exit Sub
   End If
   UnitAnalysisButtonMode EditMode
   UNIT_ANALYSIS_NEW_ENTRY = False
   Frame5.Enabled = False
End Sub

Private Sub cmdAnalysisNew_Click()
   UnitAnalysisButtonMode NewEntryMode
   UNIT_ANALYSIS_NEW_ENTRY = True
   Frame5.Enabled = False
End Sub

Private Sub cmdAnalysisSave_Click()
   Dim rdoConn As New ADODB.Connection
   rdoConn.Open getConnectionString
   Dim rsTransaction As New ADODB.Recordset
   rsTransaction.Open "Select UnitNumber from Units where UnitNumber='" & txtUnitNo.text & "'", rdoConn, adOpenKeyset
   If rsTransaction.EOF Then
        MsgBox "This Unit number was not found in the database", vbInformation, "Not found"
        FocusControl cmdUnitLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   If SaveUnitAnalysis(rdoConn) Then
      ShowMsgInTaskBar "The unit analysis have been saved successfully."
      PopulateGridUnitAnalysis rdoConn
      SetTotalArea
   Else
       ShowMsgInTaskBar "Could not save unit analysis", , "N"
   End If
   UnitAnalysisButtonMode DefaultMode
   rdoConn.Close
   Set rdoConn = Nothing
   Frame5.Enabled = True
End Sub

Private Sub cmdCancelMHistory_Click()
MaintenanceHistoryButtonMode DefaultMode
End Sub

Private Sub cmdAsJS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame7.Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdAsJS_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If lblAsJS_PO = "Print as..." Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3)

      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "JOB SHEET"

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   If lblAsJS_PO = "Email as..." Then
      
   End If
   cmdAsJS_KeyPress 27
End Sub

Private Sub cmdAsPO_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If lblAsJS_PO = "Print as..." Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3)

      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "PURCHASE ORDER"

      Load frmReport
      frmReport.LoadReportViewer Report
   End If

   cmdAsJS_KeyPress 27
End Sub

Private Sub cmdAsPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame7.Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdAttachment_Click()
   Me.Enabled = False
   Load frmAttachment

   If HEALTH_SAFETY_NEW_ENTRY Then
      If txtUnitSafetyID.text = "" Then txtUnitSafetyID.text = UniqueID()
   Else
      txtUnitSafetyID.text = gridSafety.TextMatrix(gridSafety.row, 0)
   End If

   HEALTH_N_SAFETY_ATTACH = False

   frmAttachment.OwnerID = txtUnitSafetyID.text
   frmAttachment.CallerForm = "Unit"
   frmAttachment.Show
End Sub

Private Sub cmdCancelUnit_Click()
   'ComponentEnableModeUnit frmUnits2, DefaultMode
''         Dim ctrl As Control
''         For Each ctrl In Me.Controls
''            Select Case TypeName(ctrl)
'''               Case "TextBox"
'''                  ctrl.Enabled = False
'''                  ctrl.text = ""
''               Case "CheckBox"
''                  ctrl.Enabled = False
''               Case "DataCombo"
''                  ctrl.Enabled = False
''                  ctrl.text = ""
''            End Select
''         Next ctrl
'Control State mode by anol 23 Aug 2016
   txtUnitName.text = ""
   txtUnitNo.text = ""
   txtUnitName.Locked = True
   txtUnitName.text = ""
   txtUnitAddress1.Locked = True
   txtUnitAddress1.text = ""
   txtUnitAddress2.Locked = True
   txtUnitAddress2.text = ""
   txtUnitAddress3.Locked = True
   txtUnitAddress3.text = ""
   txtUnitAddress4.Locked = True
   txtUnitAddress4.text = ""
   txtUnitPostCode.Locked = True
   txtUnitPostCode.text = ""
   cboUnitType.Locked = True
   cboUnitType.text = ""
   txtCurrentUsage.Locked = True
   txtCurrentUsage.text = ""
   cboStatus.Locked = True
   cboStatus.text = ""
   cboManagement.Locked = True
   cboManagement.text = ""
   cboCurrentTenant.Locked = True
   cboCurrentTenant.text = ""
   cboLandLord.Locked = True
   cboLandLord.text = ""
   txtTotalArea.Locked = True
   txtTotalArea.text = ""
   txtRentalPrice.Locked = True
   txtRentalPrice.text = ""
   'End of addition
   
         Me.Controls("gridUnitLookup").Visible = False

         Me.cmdNewUnit.Enabled = True
         Me.cmdEditUnit.Enabled = False
         Me.cmdCopyUnit.Enabled = False
         Me.cmdSaveUnit.Enabled = False
         Me.cmdCancelUnit.Enabled = False
         Me.cmdCloseUnit.Enabled = True
         Me.cmdUploadImageAdd.Enabled = False
    'End of default mode
         
   NEWMODE_ = False
   SEARCHUNITMODE_ = True
   txtUnitNo.Enabled = True
   txtUnitNo.Locked = True
   txtUnitName.Enabled = True
   cmdUnitLookup.Enabled = True
'   cboProperty.Enabled = True

'''''''''''''''''''''''''''added by mahboob 13/03/2023 Change ID:1/Work Item:7 :- enabling and disabling the following button controls
cmdNewUnit.Enabled = True
   cmdEditUnit.Enabled = True
   cmdDelete.Enabled = True
   cmdCancelUnit.Enabled = False
   txtPropertyName.text = ""
   '''''''''''''''''''''''''''''End of modification

   If txtUnitNo.text = "" Then
      Exit Sub
   End If
   tabUnits.Enabled = True
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub

   AddNewAttachmentInCombo cmbFiles, "Unit", txtUnitNo.text

   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdCloseUnit_Click()
   Unload Me
End Sub

Private Sub cmdCopyUnit_Click()
'   cboProperty.Enabled = True
   If txtPropertyName.text = "" Then
       ShowMsgInTaskBar "Unable to generate a unit number. Please select a property to continue.", , "N"
       cmdProperty.SetFocus
       Exit Sub
   End If

   NEWMODE_ = True
   SEARCHUNITMODE_ = False

   cboCurrentTenant.Enabled = False
   cboStatus.Enabled = False

   cmdProperty.SetFocus
   tabUnits.Enabled = False
   cmdUnitLookup.Enabled = False
   txtUnitNo.text = GenerateUnitNumber
   txtCurrentUsage.text = ""

   txtUnitAddress1.Enabled = True
   txtUnitAddress2.Enabled = True
   txtUnitAddress3.Enabled = True
   txtUnitAddress4.Enabled = True
   txtUnitName.Enabled = True
   cboUnitType.Enabled = True
   cboStatus.Enabled = True
   cboManagement.Enabled = True
   cboCurrentTenant.Enabled = True
   cboLandLord.Enabled = True
   txtTotalArea.Enabled = True

   cmdUnitType.Enabled = True
   cmdManagement.Enabled = True
   cmdCurrentTenant.Enabled = True
   cmdLandlord.Enabled = True
   cmdSaveUnit.Enabled = True
   cmdCopyUnit.Enabled = False
   cmdNewUnit.Enabled = False
   cmdEditUnit.Enabled = False
   cmdCancelUnit.Enabled = True

   lblTenantCompany.Caption = ""
   lblTenantContact.Caption = ""
   lblTenantDirectLine.Caption = ""
   lblTenantEmail.Caption = ""
   lblTenantMobile.Caption = ""
   lblTenantSageAcc.Caption = ""

   lblClientAddress1.Caption = ""
   lblClientAddress2.Caption = ""
   lblClientEmail.Caption = ""
   lblClientMobile.Caption = ""
   lblClientTelephone.Caption = ""
   lblClientName.Caption = ""
   txtUnitAddress1.SetFocus
   SelTxtInCtrl txtUnitAddress1
   txtUnitNo.Enabled = True
   txtUnitNo.Locked = False
End Sub

Private Sub cmdCurrentTenant_Click()
   Load frmLeasee1
   frmLeasee1.Show
End Sub

Private Sub cmdEditMHistory_Click()
   If gridMaintenanceHistory.TextMatrix(1, 0) = "" Then Exit Sub

   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      frmMaintenanceJob.isEdit = True
      frmMaintenanceJob.CallingForm = "U"                         'Unit
      frmMaintenanceJob.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintenanceJob
      frmMaintenanceJob.ZOrder 0
      frmMaintenanceJob.Show
   Else
      frmMaintananceDairy.isEdit = True
      frmMaintananceDairy.CallingForm = "U"                       'Unit
      frmMaintananceDairy.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintananceDairy
      frmMaintananceDairy.ZOrder 0
      frmMaintananceDairy.Show
   End If
   Me.Enabled = False
End Sub

Private Sub cmdEditUnit_Click()
   If txtUnitNo.text = "" Then
   
   ''''''''''''''''''Modified By Mahboob 13/03/2023 Change ID 1/Work Item 4:- set message box while unit number is empty
       MsgBox "Please select a unit to edit.", vbInformation, "Select Unit "
       cmdUnitLookup.SetFocus
       '''''''''''''''End of change
       
       'ShowMsgInTaskBar "Please select a unit to continue.", , "N"
       Exit Sub
   End If
    Dim adoConn As New ADODB.Connection
   Dim rsTransaction As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsTransaction.Open "Select UnitNumber from Units where UnitNumber='" & txtUnitNo.text & "'", adoConn, adOpenKeyset
   If rsTransaction.EOF Then
        MsgBox "This Unit number was not found in the database", vbInformation, "Not found"
        FocusControl cmdUnitLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   adoConn.Close
   lblMainUnit(11).Caption = txtUnitNo.text
   NEWMODE_ = False
   SEARCHUNITMODE_ = False
   ComponentEnableModeUnit frmUnits2, EditMode
   txtUnitNo.Locked = False
   cboCurrentTenant.Enabled = False
   cmdProperty.SetFocus
   tabUnits.Enabled = False
   cmdUnitLookup.Enabled = False
   
   '''''''''''''''''Modification by Md. Mahboob 20230305 Change ID 1/Work Item 5 -:  To disable the following controls
   cboLandLord.Enabled = False
   cmdDelete.Enabled = False
   cmdNewUnit.Enabled = False
   txtCurrentUsage.Enabled = False
'''''''''''''''''''''End of modification

End Sub
Private Function ComponentEnableModeUnit(ByVal frmCurrent As Form, ByVal mode As ComponentMode)
   Dim ctrl As Control

   Select Case mode
   
      Case ComponentMode.DefaultMode
'         For Each ctrl In frmCurrent.Controls
'            Select Case TypeName(ctrl)
'               Case "TextBox"
'                  ctrl.Locked = True
'                  ctrl.text = ""
'               Case "CheckBox"
'                  ctrl.Enabled = False
'               Case "DataCombo"
'                  ctrl.Enabled = False
'                  ctrl.text = ""
'            End Select
'         Next ctrl
            txtUnitNo.Locked = False
            txtUnitNo.text = ""
            txtUnitName.Locked = False
            txtUnitName.text = ""
            txtUnitAddress1.Locked = False
            txtUnitAddress1.text = ""
            txtUnitAddress2.Locked = False
            txtUnitAddress2.text = ""
            txtUnitAddress3.Locked = False
            txtUnitAddress3.text = ""
            txtUnitAddress4.Locked = False
            txtUnitAddress4.text = ""
            txtUnitPostCode.Locked = False
            txtUnitPostCode.text = ""
            cboUnitType.Locked = False
            cboUnitType.text = ""
            txtCurrentUsage.Locked = False
            txtCurrentUsage.text = ""
            cboStatus.Locked = False
            cboStatus.text = ""
            cboManagement.Locked = False
            cboManagement.text = ""
            cboCurrentTenant.Locked = False
            cboCurrentTenant.text = ""
            cboLandLord.Locked = False
            cboLandLord.text = ""
            txtTotalArea.Locked = False
            txtTotalArea.text = ""
            txtRentalPrice.Locked = False
         frmCurrent.Controls("gridUnitLookup").Visible = False

         frmCurrent.cmdNewUnit.Enabled = True
         frmCurrent.cmdEditUnit.Enabled = False
         frmCurrent.cmdCopyUnit.Enabled = False
         frmCurrent.cmdSaveUnit.Enabled = False
         frmCurrent.cmdCancelUnit.Enabled = False
         frmCurrent.cmdCloseUnit.Enabled = True
         frmCurrent.cmdUploadImageAdd.Enabled = False

      Case ComponentMode.GridRowOnSelection
         For Each ctrl In frmCurrent.Controls
             Select Case TypeName(ctrl)
                 Case "TextBox"
                     ctrl.Locked = True
                     ctrl.text = ""
                 Case "CheckBox", "DataCombo"
                     ctrl.Enabled = False
             End Select
             
         Next ctrl
         frmCurrent.Controls("gridUnitLookup").Visible = False
         
         frmCurrent.cmdNewUnit.Enabled = True
         frmCurrent.cmdEditUnit.Enabled = True
         frmCurrent.cmdSaveUnit.Enabled = False
         frmCurrent.cmdCancelUnit.Enabled = False
         frmCurrent.cmdCloseUnit.Enabled = True
         frmCurrent.cmdUploadImageAdd.Enabled = False
          
      Case ComponentMode.NewEntryMode
'         For Each ctrl In frmCurrent.Controls
'             Select Case TypeName(ctrl)
'                 Case "TextBox"
'                     ctrl.Locked = False
'                     ctrl.Enabled = True
'                     ctrl.text = ""
'                 Case "DataCombo"
'                     ctrl.Enabled = True
'                     ctrl.Locked = False
'                     If ctrl.Name <> "cboProperty" Then ctrl.text = ""
'             End Select
'         Next ctrl

            txtUnitNo.Locked = False
            txtUnitNo.text = ""
           
            
            txtUnitName.Locked = False
            txtUnitName.text = ""
            txtUnitAddress1.Locked = False
            txtUnitAddress1.text = ""
            txtUnitAddress2.Locked = False
            txtUnitAddress2.text = ""
            txtUnitAddress3.Locked = False
            txtUnitAddress3.text = ""
            txtUnitAddress4.Locked = False
            txtUnitAddress4.text = ""
            txtUnitPostCode.Locked = False
            txtUnitPostCode.text = ""
            cboUnitType.Locked = False
            cboUnitType.text = ""
            txtCurrentUsage.Locked = False
            txtCurrentUsage.text = ""
            cboStatus.Locked = False
            cboStatus.text = ""
            cboManagement.Locked = False
            cboManagement.text = ""
            cboCurrentTenant.Locked = False
            cboCurrentTenant.text = ""
            cboLandLord.Locked = False
            cboLandLord.text = ""
            txtTotalArea.Locked = False
            txtTotalArea.text = ""
            txtRentalPrice.Locked = False
            
         frmCurrent.Controls("gridUnitLookup").Visible = False

         frmCurrent.cmdNewUnit.Enabled = False
         frmCurrent.cmdEditUnit.Enabled = False
         frmCurrent.cmdSaveUnit.Enabled = True
         frmCurrent.cmdCancelUnit.Enabled = True
         frmCurrent.cmdCloseUnit.Enabled = False
         frmCurrent.cmdUploadImageAdd.Enabled = True
         frmCurrent.cmdImgDelete.Enabled = True
              
      Case ComponentMode.EditMode
         For Each ctrl In frmCurrent.Controls
             Select Case TypeName(ctrl)
                Case "TextBox"
                     ctrl.Locked = False
                 Case "CheckBox", "DataCombo", "ComboBox"
                     ctrl.Enabled = True
                                                             
             End Select
         Next ctrl
         frmCurrent.Controls("gridUnitLookup").Visible = False

         frmCurrent.cmdNewUnit.Enabled = False
         frmCurrent.cmdEditUnit.Enabled = False
         frmCurrent.cmdCopyUnit.Enabled = False
         frmCurrent.cmdSaveUnit.Enabled = True
         frmCurrent.cmdCancelUnit.Enabled = True
         frmCurrent.cmdCloseUnit.Enabled = False
         frmCurrent.cmdUploadImageAdd.Enabled = True
         frmCurrent.cmdImgDelete.Enabled = True
      
      Case ComponentMode.GridLostFocus
         For Each ctrl In frmCurrent.Controls
             Select Case TypeName(ctrl)
                  Case "TextBox"
                     ctrl.Locked = True
                  Case "CheckBox", "DataCombo"
                     ctrl.Enabled = False
             End Select
         Next ctrl
         frmCurrent.Controls("gridUnitLookup").Visible = False

         frmCurrent.cmdNewUnit.Enabled = True
         frmCurrent.cmdEditUnit.Enabled = False
         frmCurrent.cmdSaveUnit.Enabled = False
         frmCurrent.cmdCancelUnit.Enabled = False
         frmCurrent.cmdCloseUnit.Enabled = True
         frmCurrent.cmdUploadImageAdd.Enabled = False
       
   End Select
End Function
Private Sub cmdEmailJS_PO_Click()
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      fraJS_PO.Top = Frame7.Top + cmdEmailJS_PO.Top + cmdEmailJS_PO.Height - fraJS_PO.Height
      fraJS_PO.Left = cmdEmailJS_PO.Left + Frame7.Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
      Frame7.Enabled = False
      lblAsJS_PO = "Email as..."
   End If
End Sub

Private Sub cmdGridUnitLookup_Click()
   fmeUnitLookup.Visible = False
End Sub

Private Sub cmdImgDelete_Click()
   If imgUnitPicture.Picture = 0 Then Exit Sub
   If MsgBox("Are you sure to delete the image?", vbQuestion + vbYesNo, "Delete Image") = vbNo Then Exit Sub
   DeleteImage imgUnitPicture, IMAGE_FILE_NAME_, txtUnitNo.text, "Units"
   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdImgLeftMove_Click()
   IMAGE_FILE_NAME_ = MoveNextImage(imgUnitPicture, txtUnitNo.text, "Units", IMAGE_FILE_NAME_, lblImageName)
End Sub

Private Sub cmdInspectedBy_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "IPT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInspector.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IPT'"

   adoInspector.RecordSource = sSQLQuery_
   adoInspector.CommandType = adCmdText
   adoInspector.Refresh
End Sub

Private Sub cmdInsuranceCancel_Click()
   InsuranceButtonMode DefaultMode
   Frame5.Enabled = True
End Sub

Private Sub cmdInsuranceEdit_Click()
   If txtPropertyInsuranceID.text = "" Then
      Exit Sub
   End If

   InsuranceButtonMode EditMode
   UNIT_INSURANCE_NEW_ENTRY = False
   Frame5.Enabled = False
End Sub

Private Sub cmdInsuranceNew_Click()
   InsuranceButtonMode NewEntryMode
   UNIT_INSURANCE_NEW_ENTRY = True
   INSURANCE_ID = ""
   Frame5.Enabled = False
End Sub

Private Sub cmdInsuranceSave_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If SaveUnitInsurance(adoConn) Then
      ShowMsgInTaskBar "The Insurance Information have been saved successfully."
      PopulateInsurance adoConn
   Else
       ShowMsgInTaskBar "Could not save the Insurance Information", , "N"
   End If

   InsuranceButtonMode DefaultMode
   adoConn.Close
   Set adoConn = Nothing
   Frame5.Enabled = True
End Sub

Private Sub cmdLandlord_Click()
   Load frmClientNew4
   frmClientNew4.Show
End Sub

Private Sub cmdManagement_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "TNTYPE"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoManagement.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'TNTYPE'"

   adoManagement.RecordSource = sSQLQuery_
   adoManagement.CommandType = adCmdText
   adoManagement.Refresh
End Sub
'
'Private Sub cmdMType_Click()
'   Dim sSQLQuery_ As String
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "MTYP"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoMType.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'MTYP'"
'
'   adoMType.RecordSource = sSQLQuery_
'   adoMType.CommandType = adCmdText
'   adoMType.Refresh
'End Sub

Private Sub cmdNewMHistory_Click()
   If txtUnitNo.text = "" Then Exit Sub

'   Load frmMaintenanceJob
'   With frmMaintenanceJob
'      .CallingForm = "U"          'Calling from lessee form
'      .RecordType = "J"
'      .lblJobName.Caption = "Job Name"
'      .Label1.Caption = "Job No."
'      .txtRef.Enabled = True
'      .isEdit = False
'      .Show
'      .ZOrder 0
'   End With
'
'   Me.Enabled = False
    If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
        With frmMaintenanceJob
          .isEdit = True
          .CallingForm = "U"          'Calling from Unit form
          .RecordType = "J"
          .lblJobName.Caption = "Job Name"
          .Label1.Caption = "Job No."
          .txtRef.Enabled = True
          .UpdateRow = gridMaintenanceHistory.row
          .Frame1.Enabled = False
          .Show
          .ZOrder 0
        End With
    Else
        ShowMsgInTaskBar "Please select a job."
    End If
End Sub

Private Sub cmdNewUnit_Click()
   cmdProperty.Enabled = True
   
   '''''''''''''''''Change by Md. Mahboob Change ID 2/Work Item 2 for disable auto generate unit number cha
'   If txtPropertyName.text = "" Then
       ShowMsgInTaskBar "Please select a Property to continue.", , "N"
       cmdProperty.SetFocus
       'GoTo Disable_Button
'       Exit Sub
'   End If
   
'   If txtPropertyName.text = "" Then
'       ShowMsgInTaskBar "Unable to generate a unit number. Please select a property to continue.", "Y", "N"
'       cmdProperty.SetFocus
'       Exit Sub
'   End If
'''''''''''''''''''''''''End of modification
   NEWMODE_ = True
   SEARCHUNITMODE_ = False
'   ComponentEnableModeUnit frmUnits2, NewEntryMode
   '
   txtUnitName.text = ""
   txtUnitNo.text = ""
   txtUnitName.Locked = False
   txtUnitName.text = ""
   txtUnitAddress1.Locked = False
   txtUnitAddress1.text = ""
   txtUnitAddress2.Locked = False
   txtUnitAddress2.text = ""
   txtUnitAddress3.Locked = False
   txtUnitAddress3.text = ""
   txtUnitAddress4.Locked = False
   txtUnitAddress4.text = ""
   txtUnitPostCode.Locked = False
   txtUnitPostCode.text = ""
   cboUnitType.Locked = False
   cboUnitType.Enabled = True
   cboUnitType.text = ""
   
   txtCurrentUsage.Locked = False
   txtCurrentUsage.Enabled = True
   txtCurrentUsage.text = ""
   
   cboStatus.Locked = False
   cboStatus.Enabled = True
   cboStatus.text = ""
   
   cboManagement.Locked = False
   cboManagement.text = ""
   cboManagement.Enabled = True
   
   cboCurrentTenant.Locked = False
   cboCurrentTenant.text = ""
   cboCurrentTenant.Enabled = True
   
   cboLandLord.Locked = False
   cboLandLord.text = ""
   cboLandLord.Enabled = True
   
   txtTotalArea.Locked = False
   txtTotalArea.text = ""
   txtTotalArea.Enabled = True
   
   txtRentalPrice.Locked = False
   txtRentalPrice.Enabled = True
   
   
   cboCurrentTenant.Enabled = False
   cboStatus.Enabled = False
   tabUnits.Enabled = False
   cmdUnitLookup.Enabled = False
   txtCurrentUsage.text = ""
   '''''''''''''''''Modification by Md. Mahboob 20230208 Change ID 2/Work Item 2-:  To allow text enter in unit no text box
   '''''''''''''''''on Add New Unit
   txtUnitNo.Locked = False
''''''''''''''''''End of modification
   lblTenantCompany.Caption = ""
   lblTenantContact.Caption = ""
   lblTenantDirectLine.Caption = ""
   lblTenantEmail.Caption = ""
   lblTenantMobile.Caption = ""
   lblTenantSageAcc.Caption = ""

   lblClientAddress1.Caption = ""
   lblClientAddress2.Caption = ""
   lblClientEmail.Caption = ""
   lblClientMobile.Caption = ""
   lblClientTelephone.Caption = ""
   lblClientName.Caption = ""
   cmdProperty.SetFocus
   cmdSaveUnit.Enabled = True
   
   '''''''''''''''''Modification by Md. Mahboob 20230208 Change ID 1/Work Item 3-:  To disable/enable the following buttons and controls
   '''''''''''''''''on Add New Unit

   cmdEditUnit.Enabled = False
    cmdDelete.Enabled = False
    cmdNewUnit.Enabled = False
    cmdCancelUnit.Enabled = True
    cboLandLord.Enabled = False
    txtCurrentUsage.Enabled = False
   
   '''''''''''''''''Modification by Md. Mahboob 20230208 Change ID 3/Work Item 3-:  To disable/enable the following buttons and controls
   '''''''''''''''''on Add New Unit

    cboUnitType.Enabled = False
'    txtPropertyName.text = ""

   '''''''''''''''''Modification by Md. Mahboob 20230208 Change ID 6/Work Item 2-:  To disable/enable the following buttons and controls
   '''''''''''''''''on Add New Unit

   cboManagement.Enabled = False
   ''''''''''End of modification
   
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
'   MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      ShowMsgInTaskBar "File has been moved from original location."

'   MousePointer = vbDefault
End Sub

Private Sub cmdPrintJobSheet_Click()
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      fraJS_PO.Top = Frame7.Top + cmdEmailJS_PO.Top + cmdEmailJS_PO.Height - fraJS_PO.Height
      fraJS_PO.Left = cmdPrintJobSheet.Left + Frame7.Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
      Frame7.Enabled = False
      lblAsJS_PO = "Print as..."
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3)

   Report.ParameterFields(2).AddCurrentValue "Diary Entry"
   Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdproperty_Click()
    sTextBox = "2"
    picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadPropertyList
    'fraGrid.Enabled = False
    picClient.Enabled = True
    
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdQuoteReq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame7.Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdSafety_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "STYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoSafetyType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'STYP'"

   adoSafetyType.RecordSource = sSQLQuery_
   adoSafetyType.CommandType = adCmdText
   adoSafetyType.Refresh
End Sub

Private Sub cmdSafetyCancel_Click()
   HealthSafetyButtonMode DefaultMode
   Frame5.Enabled = True
End Sub

Private Sub cmdSafetyEdit_Click()
   If txtUnitSafetyID.text = "" Then Exit Sub

   HealthSafetyButtonMode EditMode
   HEALTH_SAFETY_NEW_ENTRY = False
   If gridSafety.TextMatrix(gridSafety.row, 11) = "Yes" Then
      HEALTH_N_SAFETY_ATTACH = True
   Else
      HEALTH_N_SAFETY_ATTACH = False
   End If
   Frame5.Enabled = False
End Sub

Private Sub cmdSafetyNew_Click()
   HealthSafetyButtonMode NewEntryMode
   HEALTH_SAFETY_NEW_ENTRY = True
   txtUnitSafetyID.text = ""
   Frame5.Enabled = False
End Sub

Private Sub cmdSafetySave_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If SaveHealthSafety(adoConn) Then
      ShowMsgInTaskBar "The Health and Safety Information have been saved successfully."
      PopulateHealthSafety adoConn
   Else
       ShowMsgInTaskBar "Could not save The Health and Safety Information", , "N"
   End If

   HealthSafetyButtonMode DefaultMode
   adoConn.Close
   Set adoConn = Nothing
   Frame5.Enabled = True
End Sub

Private Sub cmdSaveUnit_Click()
   If InStr(frmMMain.rtxtMessageDisplay.text, "Unit ID already exits") > 0 Then Exit Sub

   Dim conUnit_ As New ADODB.Connection
   
 '''''''''''''''''Mahboob 12/03/2023 Change ID 2/Work Item 4 Restrict user if property is empty
   If txtPropertyName.text = "" Then
       ShowMsgInTaskBar "Please select a property to continue.", "Y", "N"
       cmdProperty.SetFocus
       Exit Sub
   End If
   ''''''End of modification
   
   If Trim(txtUnitNo.text) = "" Then
      ShowMsgInTaskBar "Please enter a Unit ID to continue.", , "N"
     ' txtUnitNo.text = ""
      FocusControl txtUnitNo
      Exit Sub
   End If
  If Trim(txtUnitName.text) = "" Then
      ShowMsgInTaskBar "Please enter a Unit Name to continue.", , "N"
     ' txtUnitNo.text = ""
      FocusControl txtUnitName
      Exit Sub
   End If
   conUnit_.Open getConnectionString

   SaveUnitInformation conUnit_

   If Not NEWMODE_ And txtUnitNo.text <> lblMainUnit(11).Caption Then     'Edit mode - has user changed the Unit Code?
      conUnit_.Execute "UPDATE Units " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE UnitMaintHistory " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE PropertyInsurance " & _
                       "SET PropertyID = '" & txtUnitNo.text & "' " & _
                       "WHERE PropertyID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE UnitAnalysis " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbRechargePre " & _
                       "SET UNIT_ID = '" & txtUnitNo.text & "' " & _
                       "WHERE UNIT_ID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE UnitUtilities " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE UnitSafety " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "' AND Module = 'U';"
      conUnit_.Execute "UPDATE LeaseDetails " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbRecharged " & _
                       "SET UNIT_ID = '" & txtUnitNo.text & "' " & _
                       "WHERE UNIT_ID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE DemandRecords " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbReceipt " & _
                       "SET UnitID = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tblPoA " & _
                       "SET UnitID = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tblPrevGLU " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbCreditNote " & _
                       "SET UNIT_ID = '" & txtUnitNo.text & "' " & _
                       "WHERE UNIT_ID = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbDRCurrentPrint " & _
                       "SET UnitNumber = '" & txtUnitNo.text & "' " & _
                       "WHERE UnitNumber = '" & lblMainUnit(11).Caption & "';"
      conUnit_.Execute "UPDATE tlbFloor " & _
                       "SET UNIT_ID = '" & txtUnitNo.text & "' " & _
                       "WHERE UNIT_ID = '" & lblMainUnit(11).Caption & "';"
   End If

   Set conUnit_ = Nothing

   NEWMODE_ = False
   ComponentEnableModeUnit frmUnits2, DefaultMode
   cmdProperty.Enabled = True
   SEARCHUNITMODE_ = True
   txtUnitNo.Enabled = True
   fmeUnitLookup.Visible = False
   txtUnitName.Enabled = True
   tabUnits.Enabled = True
   cmdUnitLookup.Enabled = True
End Sub

Private Sub cmdSetInsuranceType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "ITYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsuranceType.Refresh
End Sub

Private Sub cmdSetInsurer_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "IRER"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsurer.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'IRER'"

   adoInsurer.RecordSource = sSQLQuery_
   adoInsurer.CommandType = adCmdText
   adoInsurer.Refresh
End Sub

Private Sub cmdSetUtilitiesType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UTIL"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoUtilitiesType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTIL'"

   adoUtilitiesType.RecordSource = sSQLQuery_
   adoUtilitiesType.CommandType = adCmdText
   adoUtilitiesType.Refresh
End Sub

Private Sub cmdUnitLookup_Click()
  ' On Error Resume Next
    txtSearchClientID.Enabled = True
    txtSearchClientName.Enabled = True
    txtSearchUnit.Locked = False
    txtSearchUnit.Enabled = True
    txtSearchName.Locked = False
    txtSearchName.Enabled = True
    txtSearchAddress.Locked = False
    txtSearchAddress.Enabled = True
    
    
   fmeUnitLookup.Left = Frame5.Left + txtUnitAddress1.Left
   fmeUnitLookup.Top = Frame5.Top + txtUnitAddress1.Top

   fmeUnitLookup.Visible = True
   fmeUnitLookup.ZOrder 0
   gridUnitLookup.Visible = True

   txtSearchUnit.SetFocus
   txtSearchUnit.text = ""
   txtSearchAddress.Enabled = True
   txtSearchName.Enabled = True
   
   '''''''''''''''''''''''''''Modified by Mahboob 15/03/2023 Change ID 1/Work Item 8 disable cancel and enable delete
   cmdCancelUnit.Enabled = False
   cmdDelete.Enabled = True
   ''''''''''''''''''''''''''''''end of modification
   
   If txtPropertyName.text = "" Then
      LoadGridUnitLookup ""
   Else
      LoadGridUnitLookup "AND (((UNITS.PROPERTYID) = '" & txtPropertyName.Tag & "'));"
   End If
End Sub

Private Sub cmdUnitMemoCancel_Click()
   UnitMemoButtonMode DefaultMode
End Sub

Private Sub cmdUnitMemoEdit_Click()
   UnitMemoButtonMode EditMode
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveUnitMemo Then
      ShowMsgInTaskBar "The memo has been saved successfully."
   Else
      ShowMsgInTaskBar "Could not save the memo. ", , "N"
   End If
   UnitMemoButtonMode DefaultMode
End Sub

Private Sub cmdUnitType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UTYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1
   
   adoUnitType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTYP'"

   adoUnitType.RecordSource = sSQLQuery_
   adoUnitType.CommandType = adCmdText
   adoUnitType.Refresh

   'cboUnitType.RowSource = adoUnitType.Recordset
   cboUnitType.BoundColumn = "CODE"
   cboUnitType.ListField = "VALUE"
End Sub

Private Sub cmdUploadImageAdd_Click()
   If MsgBox("Do you want to add new image?", vbQuestion + vbYesNo, "Image Attachment") = vbNo Then Exit Sub
   IMAGE_FILE_NAME_ = AddNewImage(imgUnitPicture, "Units", txtUnitNo.text, lblImageName)
   ShowMsgInTaskBar "Image has been uploaded successfull."
End Sub

Private Sub cmdUsage_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UUSE"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsUsage.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UUSE'"

   adoInsUsage.RecordSource = sSQLQuery_
   adoInsUsage.CommandType = adCmdText
   adoInsUsage.Refresh
End Sub

Private Sub cmdUStatus_Click()
   frmSecondaryCode.PRIMARY_CODE_SHOW = "USTA"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoUStatus.Refresh
End Sub

Private Sub cmdUtilitiesAttach_Click()
   Me.Enabled = False
   Load frmAttachment

   If UNIT_INSURANCE_NEW_ENTRY Then
      If INSURANCE_ID = "" Then INSURANCE_ID = UniqueID()
   Else
      INSURANCE_ID = gridSafety.TextMatrix(gridSafety.row, 0)
   End If

   HEALTH_N_SAFETY_ATTACH = False

   frmAttachment.OwnerID = INSURANCE_ID
   frmAttachment.CallerForm = "Unit_Insurance"
   frmAttachment.Show
End Sub

Private Sub cmdUtilitiesCancel_Click()
   UtilitiesButtonMode DefaultMode
   Frame5.Enabled = True
End Sub

Private Sub cmdUtilitiesEdit_Click()
   If txtUnitUtilitiesID.text = "" Then
       Exit Sub
   End If

   UtilitiesButtonMode EditMode
   UNIT_UTILITIES_NEW_ENTRY = False
   Frame5.Enabled = False
End Sub

Private Sub cmdUtilitiesNew_Click()
   UtilitiesButtonMode NewEntryMode

   UNIT_UTILITIES_NEW_ENTRY = True

   cboUtilitiesType.SetFocus
   Frame5.Enabled = False
End Sub

Private Sub cmdUtilitiesSave_Click()
   Dim rdoConn As New ADODB.Connection
   rdoConn.Open getConnectionString

   If SaveUnitUtilities(rdoConn) Then
      ShowMsgInTaskBar "The Utilities Information have been saved successfully."
      PopulateUtilities rdoConn
   Else
       ShowMsgInTaskBar "Could not save Utilities Information", , "N"
   End If

   UtilitiesButtonMode DefaultMode

   rdoConn.Close
   Set rdoConn = Nothing
   Frame5.Enabled = True
End Sub

Private Sub Form_Activate()
   If LOAD_UNIT_UNITID <> "" Then
      LoadUnitByUnitID
      Me.Caption = txtPropertyName.text + " - " + txtUnitName.text
   End If
End Sub

Private Sub Form_Load()
   'MousePointer = vbHourglass

   tabUnits.Tab = 2
   
   Me.Height = 8985
   Me.Width = 12600
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame5.BackColor = MODULEBACKCOLOR
   tabUnits.BackColor = MODULEBACKCOLOR

   Me.Caption = "Units"
   DSN_ALARM_ = "WD_ALARM"
   'ComponentEnableModeUnit frmUnits2, DefaultMode

   txtSearchUnit.Enabled = True
   NEWMODE_ = False
   SEARCHUNITMODE_ = True
   tabUnits.Enabled = False
   cmdProperty.Enabled = True
   tabUnits.Tab = 3                 ' select this tab first then come back to tab 0.
   tabUnits.Tab = 0                 ' its making problem

   '' Populate the codes
   PopulateCodes

   '' Button Modes''
   MaintenanceHistoryButtonMode DefaultMode
   UnitAnalysisButtonMode DefaultMode
   HealthSafetyButtonMode DefaultMode
   InsuranceButtonMode DefaultMode
   UtilitiesButtonMode DefaultMode
   UnitMemoButtonMode DefaultMode

   ''Set the grids

   Dim rdoConn As New ADODB.Connection
   rdoConn.Open getConnectionString

   SetGridUnitAnalysisHeader rdoConn
'   ConfigGridMaintenanceHistory rdoConn
   ConfigureGridSafety
   ConfigureGridUtilities
   FlexGridAccountHistoryConfigure

   rdoConn.Close
   Set rdoConn = Nothing
   'text mode
   txtPropertyName.Locked = True
   txtUnitName.text = ""
   txtUnitNo.text = ""
   txtUnitName.Locked = True
   txtUnitName.text = ""
   txtUnitAddress1.Locked = True
   txtUnitAddress1.text = ""
   txtUnitAddress2.Locked = True
   txtUnitAddress2.text = ""
   txtUnitAddress3.Locked = True
   txtUnitAddress3.text = ""
   txtUnitAddress4.Locked = True
   txtUnitAddress4.text = ""
   txtUnitPostCode.Locked = True
   txtUnitPostCode.text = ""
   cboUnitType.Locked = True
   cboUnitType.text = ""
   txtCurrentUsage.Locked = True
   txtCurrentUsage.text = ""
   cboStatus.Locked = True
   cboStatus.text = ""
   cboManagement.Locked = True
   cboManagement.text = ""
   cboCurrentTenant.Locked = True
   cboCurrentTenant.text = ""
   cboLandLord.Locked = True
   cboLandLord.text = ""
   txtTotalArea.Locked = True
   txtTotalArea.text = ""
   txtRentalPrice.Locked = True
   txtRentalPrice.text = ""
   '''''''''''''''''''''''''''''Modification By Mahboob Save and cancell button disable
   cmdSaveUnit.Enabled = False
   cmdCancelUnit.Enabled = False
      
   'End of addition
   
   Call WheelHook(Me.hWnd)
   'MousePointer = vbDefault
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   LOAD_UNIT_UNITID = ""
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub gridACHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridACHistory.ToolTipText = gridACHistory.TextMatrix(gridACHistory.MouseRow, gridACHistory.MouseCol)
End Sub

Private Sub gridInsurance_Click()
   InsuranceButtonMode GridRowOnSelection
End Sub

Private Sub gridInsurance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridInsurance.ToolTipText = gridInsurance.TextMatrix(gridInsurance.MouseRow, gridInsurance.MouseCol)
End Sub

Private Sub gridInsurance_RowColChange()
'   Dim iSelectedRow As Integer
'
'   iSelectedRow = gridInsurance.Row
'
'   cboInsurer.text = gridInsurance.TextMatrix(iSelectedRow, 1)
'   cboInsuranceType.Value = gridInsurance.TextMatrix(iSelectedRow, 2)
'   txtPolicyNo.text = gridInsurance.TextMatrix(iSelectedRow, 3)
'   txtSumInsured.text = gridInsurance.TextMatrix(iSelectedRow, 4)
'   txtAnnualPR.text = gridInsurance.TextMatrix(iSelectedRow, 5)
'   txtExpiryDate.text = gridInsurance.TextMatrix(iSelectedRow, 7)
'   txtTelephone.text = gridInsurance.TextMatrix(iSelectedRow, 11)

   populateControl Me, gridInsurance
End Sub

Private Sub gridMaintenanceHistory_Click()
   populateControl Me, gridMaintenanceHistory

   MaintenanceHistoryButtonMode GridRowOnSelection
End Sub

Private Sub gridMaintenanceHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridMaintenanceHistory.ToolTipText = gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.MouseRow, gridMaintenanceHistory.MouseCol)
End Sub

Private Sub gridMaintenanceHistory_RowColChange()
populateControl frmUnits2, gridMaintenanceHistory
End Sub

Private Sub gridSafety_Click()
   HealthSafetyButtonMode GridRowOnSelection
End Sub

Private Sub gridSafety_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridSafety.ToolTipText = gridSafety.TextMatrix(gridSafety.MouseRow, gridSafety.MouseCol)
End Sub

Private Sub gridSafety_RowColChange()
   populateControl Me, gridSafety
End Sub

Private Sub gridUnitAnalysis_Click()
   UnitAnalysisButtonMode GridRowOnSelection
End Sub

Private Sub gridUnitAnalysis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridUnitAnalysis.ToolTipText = gridUnitAnalysis.TextMatrix(gridUnitAnalysis.MouseRow, gridUnitAnalysis.MouseCol)
End Sub

Private Sub gridUnitAnalysis_RowColChange()
   populateControl Me, gridUnitAnalysis
End Sub

Private Sub gridUnitLookup_Click()
   SEARCHUNITMODE_ = False

   '' LOAD MAIN UNIT INFORMATION

   fmeLoading.Visible = True
   fmeLoading.Refresh
   'added by anol 23 Aug 2016
   txtPropertyName.text = gridUnitLookup.TextMatrix(gridUnitLookup.row, 6)
   txtPropertyName.Tag = gridUnitLookup.TextMatrix(gridUnitLookup.row, 7)
   'End of modification
   PopulateUnitInformation gridUnitLookup.TextMatrix(gridUnitLookup.row, 0)
   IMAGE_FILE_NAME_ = ImageLoader(imgUnitPicture, txtUnitNo.text, "UNITS", lblImageName)
   Me.Caption = txtPropertyName.text + " - " + txtUnitName.text

   '' LOAD UNIT DETAIL INFORMATION
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   PopulateTenantInformation adoConn
   PopulateClientInformation adoConn
   PopulateGridUnitAnalysis adoConn
   PopulateLeaseInformation adoConn
   LoadGridMaintenanceHistory adoConn
   PopulateHealthSafety adoConn
   PopulateInsurance adoConn
   PopulateUtilities adoConn
   RetrieveUnitMemo adoConn

   adoConn.Close
   Set adoConn = Nothing

   fmeLoading.Visible = False
   '' SET OTHERS
   fmeUnitLookup.Visible = False
   SEARCHUNITMODE_ = True
   tabUnits.Enabled = True
   gridUnitAnalysis.row = 0
   gridUnitAnalysis.col = 0
   gridUnitAnalysis.SetFocus

   cmdCopyUnit.Enabled = True
   cmdEditUnit.Enabled = True
   
   'Control State mode by anol 23 Aug 2016
   txtUnitAddress1.Locked = True
   txtUnitAddress2.Locked = True
   txtUnitAddress3.Locked = True
   txtUnitAddress4.Locked = True
   txtUnitPostCode.Locked = True
   cboUnitType.Locked = True
   txtCurrentUsage.Locked = True
   cboStatus.Locked = True
   cboManagement.Locked = True
   cboCurrentTenant.Locked = True
   cboLandLord.Locked = True
   txtTotalArea.Locked = True
   txtRentalPrice.Locked = True
   'End of addition
   
End Sub

Private Sub LoadUnitByUnitID()
   SEARCHUNITMODE_ = False

   '' LOAD MAIN UNIT INFORMATION

   fmeLoading.Visible = True
   fmeLoading.Refresh
   
   PopulateUnitInformation LOAD_UNIT_UNITID
   IMAGE_FILE_NAME_ = ImageLoader(imgUnitPicture, txtUnitNo.text, "UNITS", lblImageName)

   '' LOAD UNIT DETAIL INFORMATION
   Dim rdoConn As New ADODB.Connection
   rdoConn.Open getConnectionString

   PopulateTenantInformation rdoConn
   PopulateClientInformation rdoConn
   PopulateGridUnitAnalysis rdoConn
   PopulateLeaseInformation rdoConn
   LoadGridMaintenanceHistory rdoConn
   PopulateHealthSafety rdoConn
   PopulateInsurance rdoConn
   PopulateUtilities rdoConn
   RetrieveUnitMemo rdoConn
'   AccountHistory rdoConn

   rdoConn.Close
   Set rdoConn = Nothing

   fmeLoading.Visible = False
   '' SET OTHERS
   fmeUnitLookup.Visible = False
   SEARCHUNITMODE_ = True
   tabUnits.Enabled = True
   gridUnitAnalysis.row = 0
   gridUnitAnalysis.col = 0
   gridUnitAnalysis.SetFocus

   cmdCopyUnit.Enabled = True
   cmdEditUnit.Enabled = True
End Sub

Private Sub gridUnitLookup_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gridUnitLookup_Click
   End If
End Sub

Private Sub gridUtilities_Click()
   UtilitiesButtonMode GridRowOnSelection
End Sub

Private Sub gridUtilities_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridUtilities.ToolTipText = gridUtilities.TextMatrix(gridUtilities.MouseRow, gridUtilities.MouseCol)
End Sub

Private Sub gridUtilities_RowColChange()
   populateControl frmUnits2, gridUtilities
End Sub

Private Sub Label20_Click(Index As Integer)
   If Index = 0 Then                               ' Sort Tenant ID
      SortingGrid gridUnitLookup, Index, bSortingCol1
      bSortingCol1 = IIf(bSortingCol1, False, True)
      Label20(0).FontBold = True
      Label20(1).FontBold = False
      Label20(2).FontBold = False
   End If

   If Index = 1 Then                               ' Sort Tenant Name
      SortingGrid gridUnitLookup, Index, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      Label20(0).FontBold = False
      Label20(1).FontBold = True
      Label20(2).FontBold = False
   End If

   If Index = 2 Then                               ' Sort Unit Name
      SortingGrid gridUnitLookup, Index, bSortingCol3
      bSortingCol3 = IIf(bSortingCol3, False, True)
      Label20(0).FontBold = False
      Label20(1).FontBold = False
      Label20(2).FontBold = True
   End If
End Sub

Private Sub optAll_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget / Location"
   Label61(3).Caption = "Ref"
   Label61(1).Caption = "Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
End Sub

Private Sub optDiary_Click()
   Dim i As Integer

   Label61(10).Caption = "Location"
   Label61(3).Caption = "Diary No"
   Label61(1).Caption = "Event Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) = "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optJobs_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget"
   Label61(3).Caption = "Job No"
   Label61(1).Caption = "Maintenance Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) <> "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub txtAnalysisPercentage_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtAnalysisPercentage, KeyAscii, 4
End Sub

Private Sub txtAnalysisQuantity_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtAnalysisQuantity, KeyAscii, 0
End Sub

Private Sub txtAnalysisValue_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtAnalysisValue, KeyAscii
End Sub

Private Sub txtAnnualPR_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtAnnualPR, KeyAscii
End Sub

Private Sub txtChargeRate_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtChargeRate, KeyAscii
End Sub

Private Sub txtDateChk_Change()
   TextBoxChangeDate txtDateChk
End Sub

Private Sub txtDateChk_GotFocus()
   SelTxtInCtrl txtDateChk
End Sub

Private Sub txtDateChk_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateChk, KeyAscii
End Sub

Private Sub txtDateChk_LostFocus()
   TextBoxFormatDate txtDateChk
End Sub

Private Sub txtDateVacated_Change()
   TextBoxChangeDate txtDateVacated
End Sub

Private Sub txtDateVacated_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateVacated, KeyAscii
End Sub

Private Sub txtDateVacated_LostFocus()
   If txtDateVacated.text <> "" Then TextBoxFormatDate txtDateVacated
End Sub

Private Sub txtExpiryDate_Change()
   TextBoxChangeDate txtExpiryDate
End Sub

Private Sub txtExpiryDate_GotFocus()
   SelTxtInCtrl txtExpiryDate
End Sub

Private Sub txtExpiryDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtExpiryDate, KeyAscii
End Sub

Private Sub txtExpiryDate_LostFocus()
   TextBoxFormatDate txtExpiryDate
End Sub

Private Sub txtNextDueDate_Change()
   TextBoxChangeDate txtNextDueDate
End Sub

Private Sub txtNextDueDate_GotFocus()
   SelTxtInCtrl txtNextDueDate
End Sub

Private Sub txtNextDueDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtNextDueDate, KeyAscii
End Sub

Private Sub txtNextDueDate_LostFocus()
   TextBoxFormatDate txtNextDueDate
End Sub

Private Sub txtSearchAddress_Change()
'   Dim i As Integer
'
'   If Len(txtSearchAddress.text) > 0 Then
'      txtSearchUnit.text = ""
'      txtSearchName.text = ""
'   End If
'
'   For i = 1 To gridUnitLookup.Rows - 1
'      gridUnitLookup.RowHeight(i) = 240
'      If UCase(Left(gridUnitLookup.TextMatrix(i, 2), Len(txtSearchAddress.text))) <> UCase(txtSearchAddress.text) Then
'         gridUnitLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445, 00000442
'Modified By: Asif. 02 Aug 2014

   FilterUnitsList
End Sub

Private Sub txtSearchName_Change()
'   Dim i As Integer
'
'   If Len(txtSearchName.text) > 0 Then
'      txtSearchUnit.text = ""
'      txtSearchAddress.text = ""
'   End If
'
'   For i = 1 To gridUnitLookup.Rows - 1
'      gridUnitLookup.RowHeight(i) = 240
'      If UCase(Left(gridUnitLookup.TextMatrix(i, 1), Len(txtSearchName.text))) <> UCase(txtSearchName.text) Then
'         gridUnitLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445, 00000442
'Modified By: Asif. 02 Aug 2014

   FilterUnitsList
End Sub

Private Sub txtSearchUnit_Change()
'   Dim i As Integer
'
'   If Len(txtSearchUnit.text) > 0 Then
'      txtSearchName.text = ""
'      txtSearchAddress.text = ""
'   End If
'
'   For i = 1 To gridUnitLookup.Rows - 1
'      gridUnitLookup.RowHeight(i) = 240
'      If UCase(Left(gridUnitLookup.TextMatrix(i, 0), Len(txtSearchUnit.text))) <> UCase(txtSearchUnit.text) Then
'         gridUnitLookup.RowHeight(i) = 0
'      End If
'   Next i

'Resolved by BOSL
'Issue No: 0000445, 00000442
'Modified By: Asif. 02 Aug 2014

   FilterUnitsList
End Sub

Private Sub txtSearchUnit_KeyPress(KeyAscii As Integer)
   If Not SEARCHUNITMODE_ Then
       Exit Sub
   End If

   If KeyAscii = 13 Then
       gridUnitLookup.SetFocus
   End If
End Sub

Private Sub txtStartDate_Change()
   TextBoxChangeDate txtStartDate
End Sub

Private Sub txtStartDate_GotFocus()
   SelTxtInCtrl txtStartDate
End Sub

Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtStartDate, KeyAscii
End Sub

Private Sub txtStartDate_LostFocus()
   TextBoxFormatDate txtStartDate
End Sub

Private Sub txtSumInsured_KeyPress(KeyAscii As Integer)
   'Added By Samrat. 12/10/2006
   DigitTextKeyPress txtSumInsured, KeyAscii
End Sub

Private Sub txtTotalArea_KeyPress(KeyAscii As Integer)
    DigitTextKeyPress txtTotalArea, KeyAscii
    If KeyAscii = 13 Then
        If txtRentalPrice.Enabled Then txtRentalPrice.SetFocus
    End If
End Sub

Private Sub txtUnitAddress1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUnitAddress2.Enabled Then txtUnitAddress2.SetFocus
    End If
End Sub

Private Sub txtUnitAddress1_LostFocus()
   If txtUnitName.text = "" Then txtUnitName.text = txtUnitAddress1.text
End Sub

Private Sub txtUnitAddress2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUnitAddress3.Enabled Then txtUnitAddress3.SetFocus
    End If
End Sub

Private Sub txtUnitAddress3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUnitAddress4.Enabled Then txtUnitAddress4.SetFocus
    End If
End Sub



Private Sub txtUnitAddress4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUnitPostCode.Enabled Then txtUnitPostCode.SetFocus
    End If
End Sub

Private Sub txtUnitName_GotFocus()
   If txtUnitName.text = "" Then Exit Sub

   SelTxtInCtrl txtUnitName
End Sub

Public Sub PopulateCodes()
   Dim conLookup_ As New ADODB.Connection
   Dim rstLookup_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim i As Integer
   Dim PrimaryCode_ As String
   
   
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
     
   adoConn.Open getConnectionString

   'Set the RDO Connections to the dataset
   adoStatus.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'STAT'"

   adoStatus.RecordSource = sSQLQuery_
   adoStatus.CommandType = adCmdText
   adoStatus.Refresh

   adoUStatus.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'USTA'"         'Utility Status

   adoUStatus.RecordSource = sSQLQuery_
   adoUStatus.CommandType = adCmdText
   adoUStatus.Refresh

   adoUnitType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTYP'"

   adoUnitType.RecordSource = sSQLQuery_
   adoUnitType.CommandType = adCmdText
   adoUnitType.Refresh

   'cboUnitType.RowSource = adoUnitType.Recordset
   cboUnitType.BoundColumn = "CODE"
   cboUnitType.ListField = "VALUE"

   adoManagement.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'TNTYPE'"

   adoManagement.RecordSource = sSQLQuery_
   adoManagement.CommandType = adCmdText
   adoManagement.Refresh

   adoCurrentTenant.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT SAGEACCOUNTNUMBER, COMPANYNAME " & _
                 "FROM TENANTS "

   adoCurrentTenant.RecordSource = sSQLQuery_
   adoCurrentTenant.CommandType = adCmdText
   adoCurrentTenant.Refresh

   cboCurrentTenant.BoundColumn = "SAGEACCOUNTNUMBER"
   cboCurrentTenant.ListField = "COMPANYNAME"

   adoLandLord.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CLIENTID, CLIENTNAME " & _
                 "FROM CLIENT "

   adoLandLord.RecordSource = sSQLQuery_
   adoLandLord.CommandType = adCmdText
   adoLandLord.Refresh
   cboLandLord.BoundColumn = "CLIENTID"
   cboLandLord.ListField = "CLIENTNAME"

   adoMType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'MTYP'"

   adoMType.RecordSource = sSQLQuery_
   adoMType.CommandType = adCmdText
   adoMType.Refresh

   adoAnalysisType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'ATYP'"

   adoAnalysisType.RecordSource = sSQLQuery_
   adoAnalysisType.CommandType = adCmdText
   adoAnalysisType.Refresh

   adoSelectOption.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'SOPT'"

   adoSelectOption.RecordSource = sSQLQuery_
   adoSelectOption.CommandType = adCmdText
   adoSelectOption.Refresh
   
   adoSafetyType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'STYP'"

   adoSafetyType.RecordSource = sSQLQuery_
   adoSafetyType.CommandType = adCmdText
   adoSafetyType.Refresh

   adoSafetyStatus.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'SSTA'"
   adoSafetyStatus.RecordSource = sSQLQuery_
   adoSafetyStatus.CommandType = adCmdText
   adoSafetyStatus.Refresh

   adoUtilitiesType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTIL'"

   adoUtilitiesType.RecordSource = sSQLQuery_
   adoUtilitiesType.CommandType = adCmdText
   adoUtilitiesType.Refresh

   adoSupplier.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT SupplierID, SupplierName " & _
                 "FROM SUPPLIER "

   adoSupplier.RecordSource = sSQLQuery_
   adoSupplier.CommandType = adCmdText
   adoSupplier.Refresh

'   adoProperty.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT PROPERTYID, PROPERTYNAME " & _
'                 "FROM PROPERTY "
'   adoProperty.RecordSource = sSQLQuery_
'   adoProperty.CommandType = adCmdText
'   adoProperty.Refresh

   adoInspector.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IPT'"
   adoInspector.RecordSource = sSQLQuery_
   adoInspector.CommandType = adCmdText
   adoInspector.Refresh

   adoInsurer.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IRER'"
   adoInsurer.RecordSource = sSQLQuery_
   adoInsurer.CommandType = adCmdText
   adoInsurer.Refresh

   adoInsUsage.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'UUSE'"
   adoInsUsage.RecordSource = sSQLQuery_
   adoInsUsage.CommandType = adCmdText
   adoInsUsage.Refresh

   adoInsuranceType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'ITYP'"
   adoInsuranceType.RecordSource = sSQLQuery_
   adoInsuranceType.CommandType = adCmdText
   adoInsuranceType.Refresh

   adoConn.Close
   Set adoConn = Nothing
End Sub

Public Function PopulateUnitInformation(ByVal sUnitNumber As String) As Boolean
   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'Set the RDO Connections to the dataset
   conUnit_.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT UNITS.UnitNumber, UNITS.PropertyID, UNITS.UnitName, UNITS.UnitAddressLine1, " & _
               "UNITS.UnitAddressLine2, UNITS.UnitAddressLine3, UNITS.UnitAddressLine4, " & _
               "UNITS.UnitPostCode, UNITS.TenantCompanyName, UNITS.SageAccountNumber, " & _
               "UNITS.Frontage, UNITS.RateableValue, UNITS.RatesPayable, UNITS.GroundFloorArea, " & _
               "UNITS.MezzanineArea, UNITS.TotalArea, UNITS.UnitType, UNITS.LandLord, " & _
               "UNITS.Management, UNITS.Memo, RentalPrice " & _
           "FROM UNITS " & _
           "WHERE UNITS.UNITNUMBER = '" & sUnitNumber & "';"

   rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly
   'MsgBox sUnitNumber
   If rstUnit_.EOF Or rstUnit_.BOF Then
       ShowMsgInTaskBar "WARNING !! No information found for the specified unit.", , "N"
   End If

   While Not rstUnit_.EOF
      txtUnitNo.text = rstUnit_!UnitNumber
      txtUnitName.text = rstUnit_!UnitName
      txtPropertyName.Tag = IIf(IsNull(rstUnit_!propertyID), "", rstUnit_!propertyID)
      txtUnitAddress1.text = IIf(IsNull(rstUnit_!UnitAddressLine1), "", rstUnit_!UnitAddressLine1)
      txtUnitAddress2.text = IIf(IsNull(rstUnit_!UnitAddressLine2), "", rstUnit_!UnitAddressLine2)
      txtUnitAddress3.text = IIf(IsNull(rstUnit_!UnitAddressLine3), "", rstUnit_!UnitAddressLine3)
      txtUnitAddress4.text = IIf(IsNull(rstUnit_!UnitAddressLine4), "", rstUnit_!UnitAddressLine4)
      txtUnitPostCode.text = IIf(IsNull(rstUnit_!UnitPostCode), "", rstUnit_!UnitPostCode)

      cboUnitType.BoundText = IIf(IsNull(rstUnit_!UNITTYPE), "", rstUnit_!UNITTYPE)
      cboStatus.BoundText = UnitStatus(rstUnit_!UnitNumber, conUnit_)
      cboManagement.BoundText = IIf(IsNull(rstUnit_!MANAGEMENT), "", rstUnit_!MANAGEMENT)

      txtTotalArea.text = IIf(IsNull(rstUnit_!TotalArea), "0", rstUnit_!TotalArea)
      txtRentalPrice.text = IIf(IsNull(rstUnit_!RentalPrice), "", rstUnit_!RentalPrice)
      rstUnit_.MoveNext
   Wend

   rstUnit_.Close
   
   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT USAGE " & _
           "FROM LEASEDETAILS " & _
           "WHERE UNITNUMBER = '" & sUnitNumber & "' AND STATUS = True;"

   rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly
   'MsgBox sUnitNumber
   If rstUnit_.EOF Or rstUnit_.BOF Then
       txtCurrentUsage.text = ""
   End If

   While Not rstUnit_.EOF
      txtCurrentUsage.text = IIf(IsNull(rstUnit_!Usage), "", rstUnit_!Usage)
      rstUnit_.MoveNext
   Wend

   rstUnit_.Close
   
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing
End Function

Private Function UnitStatus(szUnitNumber As String, conUnit As ADODB.Connection) As String
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT LeaseDetails.* " & _
                "FROM LeaseDetails, UnitS " & _
                "WHERE LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
                  "LeaseDetails.Status = True And " & _
                  "Units.UnitNumber = '" & szUnitNumber & "';"

   rstUnit_.Open sSQLQuery_, conUnit, adOpenStatic, adLockReadOnly

   If Not rstUnit_.EOF Then
      UnitStatus = "Y"
   Else
      UnitStatus = "N"
   End If

   rstUnit_.Close
   Set rstUnit_ = Nothing
End Function

Public Sub RetrieveUnitMemo(ByVal conUnitMemo_ As ADODB.Connection)
   Dim rstUnitMemo_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   On Error Resume Next

   txtUnitMemo.text = ""

   sSQLQuery_ = "SELECT MEMO " & _
                "FROM UNITS WHERE UNITNUMBER = '" & txtUnitNo.text & "'"
   rstUnitMemo_.Open sSQLQuery_, conUnitMemo_, adOpenStatic, adLockReadOnly

   txtUnitMemo.text = IIf(IsNull(rstUnitMemo_!Memo), "", rstUnitMemo_!Memo) 'rstUnitMemo_!Memo '
   
   Call LoadAttachmentFiles(cmbFiles, txtUnitNo.text, "Unit")

   rstUnitMemo_.Close
   Set rstUnitMemo_ = Nothing
End Sub

Public Sub PopulateTenantInformation(ByVal conTenant As ADODB.Connection)
   Dim rstTenant As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT TENANTS.*, LEASEDETAILS.UNITNUMBER " & _
                "FROM TENANTS, LEASEDETAILS, UNITS " & _
                "WHERE TENANTS.SAGEACCOUNTNUMBER = LEASEDETAILS.SAGEACCOUNTNUMBER AND " & _
                  "LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " & _
                  "LEASEDETAILS.STATUS = TRUE AND " & _
                  "UNITS.OCCUPIED = 'Y' AND UNITS.UNITNUMBER = '" & txtUnitNo.text & "';"

   rstTenant.Open sSQLQuery_, conTenant, adOpenStatic, adLockReadOnly

   If rstTenant.EOF Or rstTenant.BOF Then
      lblTenantCompany.Caption = ""
      lblTenantSageAcc.Caption = ""
      lblTenantContact.Caption = ""
      lblTenantDirectLine.Caption = ""
      lblTenantMobile.Caption = ""
      lblTenantEmail.Caption = ""
      cboCurrentTenant.BoundText = ""
   End If

   While Not rstTenant.EOF
      cboCurrentTenant.BoundText = IIf(IsNull(rstTenant!CompanyName), "", rstTenant!CompanyName)
      lblTenantCompany.Caption = IIf(IsNull(rstTenant!CompanyName), "", rstTenant!CompanyName)
      lblTenantSageAcc.Caption = IIf(IsNull(rstTenant!SageAccountNumber), "", rstTenant!SageAccountNumber)
      lblTenantContact.Caption = IIf(IsNull(rstTenant!Contact1), "", rstTenant!Contact1)
      lblTenantDirectLine.Caption = IIf(IsNull(rstTenant!DirectLine1), "", rstTenant!DirectLine1)
      lblTenantMobile.Caption = IIf(IsNull(rstTenant!HOTelephone), "", rstTenant!HOTelephone)
      lblTenantEmail.Caption = IIf(IsNull(rstTenant!Email1), "", rstTenant!Email1)

      cboCurrentTenant.BoundText = IIf(IsNull(rstTenant!SageAccountNumber), "", rstTenant!SageAccountNumber)
      rstTenant.MoveNext
   Wend
   rstTenant.Close
   Set rstTenant = Nothing
End Sub

Public Sub PopulateLeaseInformation(ByVal conLease As ADODB.Connection)
   Dim rstLease As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT " & _
     " LeaseDetails.LeaseID, LeaseDetails.CompanyName, LeaseDetails.UnitNumber, " & _
     " Client.ClientName, " & _
     " Property.PropertyName, " & _
     " LeaseDetails.TypeOfStore, LeaseDetails.StartDate, LeaseDetails.EndDate, " & _
     " IIF(LeaseDetails.BRPayable='Y','Yes','No') AS BRPAYABLE, IIF(LeaseDetails.SCPayable='Y','Yes','No') AS SCPAYABLE, " & _
     " LeaseDetails.InsurancePayable, LeaseDetails.HoldingOver, " & _
     " RentAnalysis.RentReviewDate, " & _
     " T_BRFREQUENCY.FREQUENCY AS BR_FREQ " & _
   " FROM LeaseDetails, Client,  Property, UNITS, LRentCharges, RentAnalysis, " & _
     " (SELECT ID, FREQUENCY FROM FREQUENCIES) AS T_BRFREQUENCY " & _
   " WHERE " & _
     " LEASEDETAILS.UNITNUMBER = UNITS.UNITNUMBER AND " & _
     " UNITS.PROPERTYID = PROPERTY.PROPERTYID AND " & _
     " PROPERTY.CLIENTID = CLIENT.CLIENTID AND " & _
     " LRentCharges.BRFREQUENCY = T_BRFREQUENCY.ID AND " & _
     " LeaseDetails.LeaseID = LRentCharges.LeaseID AND " & _
     " LeaseDetails.UNITNUMBER = '" & txtUnitNo.text & "' AND " & _
     " RentAnalysis.LeaseID = LeaseDetails.LeaseID AND " & _
     " LEASEDETAILS.STATUS = TRUE;"
   
   rstLease.Open sSQLQuery_, conLease, adOpenStatic, adLockReadOnly
   'MsgBox sUnitNumber
   If rstLease.EOF Or rstLease.BOF Then
       lblLeaseHeading.Caption = "<No Lease setup for this unit>"
      
       lblLeaseReference.Caption = ""
       lblLeaseProperty.Caption = ""
       lblLeaseType.Caption = ""
       lblLeaseStartDate.Caption = ""
       lblLeaseExpiryDate.Caption = ""
       lblLeaseInsurance.Caption = ""
       lblLeaseHoldingOver.Caption = ""
       lblLeaseRentPayable.Caption = ""
       lblLeaseSCPayable.Caption = ""
       lblLeaseRentReview.Caption = ""
       lblLeaseReviewFreq.Caption = ""
   End If
   
   While Not rstLease.EOF
       lblLeaseHeading.Caption = ""
       lblLeaseReference.Caption = IIf(IsNull(rstLease!LeaseID), "", rstLease!LeaseID)
       lblLeaseProperty.Caption = IIf(IsNull(rstLease!PropertyName), "", rstLease!PropertyName)
       lblLeaseType.Caption = IIf(IsNull(rstLease!TYPEOFSTORE), "", rstLease!TYPEOFSTORE)
       lblLeaseStartDate.Caption = IIf(IsNull(rstLease!StartDate), "", rstLease!StartDate)
       lblLeaseInsurance.Caption = IIf(IsNull(rstLease!InsurancePayable), "No", IIf(rstLease!InsurancePayable = "Y", "Yes", "No"))
       lblLeaseHoldingOver.Caption = IIf(IsNull(rstLease!HoldingOver), "No", IIf(rstLease!HoldingOver = "False", "Yes", "No"))
       lblLeaseExpiryDate.Caption = IIf(IsNull(rstLease!EndDate), "", rstLease!EndDate)
       lblLeaseRentPayable.Caption = IIf(IsNull(rstLease!BRPayable), "No", rstLease!BRPayable)
       lblLeaseSCPayable.Caption = IIf(IsNull(rstLease!SCPayable), "No", rstLease!SCPayable)
       lblLeaseRentReview.Caption = IIf(IsNull(rstLease!RentReviewDate), "", rstLease!RentReviewDate)
       lblLeaseReviewFreq.Caption = IIf(IsNull(rstLease!BR_FREQ), "", rstLease!BR_FREQ)

       rstLease.MoveNext
   Wend

   rstLease.Close
   Set rstLease = Nothing
End Sub

Public Sub PopulateClientInformation(ByVal conClient As ADODB.Connection)
   Dim rstClient As New ADODB.Recordset
   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT CLIENT.CLIENTID, CLIENT.CLIENTNAME,  " & _
           "CLIENT.CLIENTADDRESSLINE1, CLIENT.CLIENTADDRESSLINE2,  CLIENT.CLIENTADDRESSLINE3, " & _
           "CLIENT.CLIENTPOSTCODE, CLIENT.CLIENTOFFICETEL,  " & _
           "CLIENT.CLIENTMOBILE, CLIENT.CLIENTOFFICEEMAIL  " & _
           "FROM CLIENT, PROPERTY, UNITS " & _
           "WHERE CLIENT.CLIENTID = PROPERTY.CLIENTID " & _
           "AND PROPERTY.PROPERTYID = UNITS.PROPERTYID " & _
           "AND UNITS.UNITNUMBER = '" & txtUnitNo.text & "'"

   rstClient.Open sSQLQuery_, conClient, adOpenStatic, adLockReadOnly

   If rstClient.EOF Or rstClient.BOF Then
       lblClientName.Caption = ""
       lblClientAddress1.Caption = ""
       lblClientAddress2.Caption = ""
       lblClientTelephone.Caption = ""
       lblClientMobile.Caption = ""
       lblClientEmail.Caption = ""
   End If

   While Not rstClient.EOF
       cboLandLord.BoundText = IIf(IsNull(rstClient!ClientID), "", rstClient!ClientID)
       lblClientName.Caption = IIf(IsNull(rstClient!ClientName), "", rstClient!ClientName)
       lblClientAddress1.Caption = IIf(IsNull(rstClient!ClientAddressLine1), "", _
                                         rstClient!ClientAddressLine1) & " " & _
                                         IIf(IsNull(rstClient!ClientAddressLine2), "", _
                                         rstClient!ClientAddressLine2)
       lblClientAddress2.Caption = IIf(IsNull(rstClient!ClientAddressLine3), "", _
                                         rstClient!ClientAddressLine3) & " " & _
                                         IIf(IsNull(rstClient!ClientPostCode), "", _
                                         rstClient!ClientPostCode)
       lblClientTelephone.Caption = IIf(IsNull(rstClient!ClientOfficeTel), "", _
                                         rstClient!ClientOfficeTel)
       lblClientMobile.Caption = IIf(IsNull(rstClient!ClientMobile), "", _
                                         rstClient!ClientMobile)
       lblClientEmail.Caption = IIf(IsNull(rstClient!ClientOfficeEmail), "", _
                                         rstClient!ClientOfficeEmail)
       rstClient.MoveNext
   Wend

   rstClient.Close
   Set rstClient = Nothing
End Sub

Private Function DoBrowse() As String
' Browse for a file
    Dim obj As New CDialog
    With obj                'find a graphics file
        .Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNFileMustExist + cdlOFNExplorer
        .Filter = "Graphics File|*.bmp;*.ico;*.emf;*.wmf;*.jpg;*.gif|All Files|*.*"
        .DialogTitle = "Select a File"
        .InitDir = Trim$(App.Path)
        .lHwnd = Me.hWnd
        .ShowOpen
        If Not .Cancelled Then
            DoBrowse = .Filename
        Else
            DoBrowse = "NONE"
        End If
    End With
    Set obj = Nothing
End Function

Public Function SaveUnitInformation(conUnit_ As ADODB.Connection) As Boolean
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   If NEWMODE_ Then
       sSQLQuery_ = "SELECT UnitNumber, PropertyID, UnitName, UnitAddressLine1, " & _
                  "UnitAddressLine2, UnitAddressLine3, UnitAddressLine4, " & _
                  "UnitPostCode, Occupied, TenantCompanyName, SageAccountNumber, " & _
                  "Frontage, RateableValue, RatesPayable, GroundFloorArea, " & _
                  "MezzanineArea, TotalArea, UnitType, LandLord, Management, Memo, RentalPrice " & _
                    "FROM UNITS " & _
                    "WHERE UNITNUMBER = '" & lblMainUnit(11).Caption & "'"

       rstUnit_.Open sSQLQuery_, conUnit_, adOpenStatic, adLockReadOnly

       If rstUnit_.EOF Or rstUnit_.BOF Then
           If AddUnitInformation Then
               ShowMsgInTaskBar "The Unit Information added successfully."
               SaveUnitInformation = True
           Else
               ShowMsgInTaskBar "An error occured while saving the Unit Information.", , "N"
               SaveUnitInformation = False
           End If
       Else
           If (MsgBox("WARNING ! The Unit Number entered already exists. Do you want to update the information", vbYesNo, "Save Unit Information") = vbYes) Then
               lblMainUnit(11).Caption = txtUnitNo.text
               If UpdateUnitInformation Then
                   ShowMsgInTaskBar "The Unit Information updated successfully."
                   SaveUnitInformation = True
               Else
                   ShowMsgInTaskBar "An error occured while updating the Unit Information.", , "N"
                   SaveUnitInformation = False
               End If
           End If
       End If

       rstUnit_.Close
'       conUnit_.Close
       Set rstUnit_ = Nothing
'       Set conUnit_ = Nothing
       SaveUnitInformation = True
       Exit Function
   Else
       If (MsgBox("Do you wish to update the Unit Information", vbYesNo + vbQuestion, "Save Unit Information") = vbYes) Then
           If UpdateUnitInformation Then
               ShowMsgInTaskBar "The Unit Information updated successfully"
               SaveUnitInformation = True
               Exit Function
           Else
'               ShowMsgInTaskBar "An error occured while updating the Unit Information", , "N"
               SaveUnitInformation = False
               Exit Function
           End If
       End If
   End If

Exception:
'   ShowMsgInTaskBar ERR.Number & " - " & ERR.description, , "N"
   If rstUnit_.State = 1 Then
         rstUnit_.Close
         Set rstUnit_ = Nothing
   End If
'   If conUnit_.State = 1 Then
'         conUnit_.Close
'         Set conUnit_ = Nothing
'   End If
   SaveUnitInformation = False
End Function

Public Function AddUnitInformation() As Boolean
    Dim conUnit_ As New ADODB.Connection
    Dim rstUnit_ As New ADODB.Recordset
    Dim sSQLQuery_ As String

'    On Error GoTo Exception
    'Set the RDO Connections to the dataset
    conUnit_.Open getConnectionString

    sSQLQuery_ = "SELECT UnitNumber, PropertyID, UnitName, UnitAddressLine1, " & _
                     "UnitAddressLine2, UnitAddressLine3, UnitAddressLine4, " & _
                     "UnitPostCode, Occupied, TenantCompanyName, SageAccountNumber, " & _
                     "Frontage, RateableValue, RatesPayable, GroundFloorArea, " & _
                     "MezzanineArea, TotalArea, UnitType, LandLord, Management, Memo, RentalPrice,CreatedBy,CreatedDate " & _
                 "FROM UNITS"
    rstUnit_.Open sSQLQuery_, conUnit_, adOpenDynamic, adLockOptimistic
    rstUnit_.AddNew
    rstUnit_!CreatedBy = User
    rstUnit_!CreatedDate = Now
    
    rstUnit_!UnitNumber = Trim(txtUnitNo.text)
    rstUnit_!propertyID = txtPropertyName.Tag
    rstUnit_!UnitName = Trim(txtUnitName.text)
    rstUnit_!UnitAddressLine1 = txtUnitAddress1.text
    rstUnit_!UnitAddressLine2 = txtUnitAddress2.text
    rstUnit_!UnitAddressLine3 = txtUnitAddress3.text
    rstUnit_!UnitAddressLine4 = txtUnitAddress4.text
    rstUnit_!UnitPostCode = txtUnitPostCode.text
    rstUnit_!UNITTYPE = cboUnitType.BoundText
    rstUnit_!MANAGEMENT = cboManagement.BoundText
    rstUnit_!TenantCompanyName = cboCurrentTenant.text
    rstUnit_!LANDLORD = cboLandLord.BoundText
    rstUnit_!SageAccountNumber = cboCurrentTenant.BoundText
    rstUnit_!RentalPrice = txtRentalPrice.text
    If NEWMODE_ Then rstUnit_!OCCUPIED = "N"
    
'** User should input the total are of the unit in the global information of the unit.
    rstUnit_!TotalArea = IIf(txtTotalArea.text = "", 0, txtTotalArea.text)
'    If Not txtAnalysisTotalArea.text = "" Then
'        rstUnit_!TotalArea = CDbl(txtAnalysisTotalArea.text)
'    End If
    rstUnit_.Update
    rstUnit_.Close
    conUnit_.Close
    Set rstUnit_ = Nothing
    Set conUnit_ = Nothing
'
    AddUnitInformation = True
    Exit Function
'
Exception:
    ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
    If rstUnit_.State = 1 Then
        rstUnit_.Close
        Set rstUnit_ = Nothing
    End If
    If conUnit_.State = 1 Then
        conUnit_.Close
        Set conUnit_ = Nothing
    End If
   
    
    AddUnitInformation = False
End Function

Public Function UpdateUnitInformation() As Boolean
   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
    If lblMainUnit(11).Caption = "(Unit Code)" Then
        Exit Function
    End If
   On Error GoTo Exception
   'Set the RDO Connections to the dataset
   conUnit_.Open getConnectionString

   sSQLQuery_ = "SELECT UnitNumber, PropertyID, UnitName, UnitAddressLine1, " & _
                  "UnitAddressLine2, UnitAddressLine3, UnitAddressLine4, " & _
                  "UnitPostCode, Occupied, TenantCompanyName, SageAccountNumber, " & _
                  "Frontage, RateableValue, RatesPayable, GroundFloorArea, " & _
                  "MezzanineArea, TotalArea, UnitType, LandLord, Management, Memo, RentalPrice " & _
                "FROM UNITS " & _
                "WHERE UNITNUMBER = '" & lblMainUnit(11).Caption & "'"
   rstUnit_.Open sSQLQuery_, conUnit_, adOpenDynamic, adLockOptimistic

   rstUnit_!propertyID = txtPropertyName.Tag
   rstUnit_!UnitName = txtUnitName.text
   rstUnit_!UnitAddressLine1 = txtUnitAddress1.text
   rstUnit_!UnitAddressLine2 = txtUnitAddress2.text
   rstUnit_!UnitAddressLine3 = txtUnitAddress3.text
   rstUnit_!UnitAddressLine4 = txtUnitAddress4.text
   rstUnit_!UnitPostCode = txtUnitPostCode.text
   rstUnit_!UNITTYPE = cboUnitType.BoundText
   rstUnit_!MANAGEMENT = cboManagement.BoundText
   rstUnit_!LANDLORD = cboLandLord.BoundText
'** If the total unit modified and there is any lease against the unit then
'   we might need to change the value of the price per square foot/metre of the SC.
   If rstUnit_!TotalArea <> IIf(txtTotalArea.text = "", 0, txtTotalArea.text) Then
      UpdateLeaseInformation conUnit_, txtUnitNo.text, txtTotalArea.text
   End If
'** User should input the total are of the unit in the global information of the unit.
   rstUnit_!TotalArea = IIf(txtTotalArea.text = "", 0, txtTotalArea.text)
   rstUnit_!RentalPrice = IIf(txtRentalPrice.text = "", "", txtRentalPrice.text)

   rstUnit_.Update

   rstUnit_.Close
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing
   UpdateUnitInformation = True
   Exit Function

Exception:

   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
   rstUnit_.Close
   conUnit_.Close
   Set rstUnit_ = Nothing
   Set conUnit_ = Nothing
   UpdateUnitInformation = False
End Function

Public Sub MaintenanceHistoryButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
    Case ComponentMode.DefaultMode
        cmdNewMHistory.Enabled = True
        cmdEditMHistory.Enabled = False
        gridMaintenanceHistory.Enabled = True
        
    Case ComponentMode.GridRowOnSelection
        cmdNewMHistory.Enabled = True
        cmdEditMHistory.Enabled = True
        gridMaintenanceHistory.Enabled = True
    
    Case ComponentMode.NewEntryMode
        cmdNewMHistory.Enabled = False
        cmdEditMHistory.Enabled = False
        gridMaintenanceHistory.Enabled = False

    Case ComponentMode.EditMode
         cmdNewMHistory.Enabled = False
         cmdEditMHistory.Enabled = False
         gridMaintenanceHistory.Enabled = False
   End Select
End Sub

Public Sub UnitAnalysisButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
        Case ComponentMode.DefaultMode
            cmdAnalysisNew.Enabled = True
            cmdAnalysisEdit.Enabled = False
            cmdAnalysisSave.Enabled = False
            cmdAnalysisCancel.Enabled = False
            
            gridUnitAnalysis.Enabled = True
        
            cboAnalysisType.Enabled = False
            cmdAnalysis.Enabled = False
            txtAnalysisDescription.Enabled = False
            cboAnalysisOption.Enabled = False
            txtAnalysisValue.Enabled = False
            txtAnalysisQuantity.Enabled = False
            txtAnalysisPercentage.Enabled = False
                
        
        Case ComponentMode.GridRowOnSelection
            cmdAnalysisNew.Enabled = True
            cmdAnalysisEdit.Enabled = True
            cmdAnalysisSave.Enabled = False
            cmdAnalysisCancel.Enabled = False
            
            gridUnitAnalysis.Enabled = True
        
        Case ComponentMode.NewEntryMode
            cmdAnalysisNew.Enabled = False
            cmdAnalysisEdit.Enabled = False
            cmdAnalysisSave.Enabled = True
            cmdAnalysisCancel.Enabled = True
            
            gridUnitAnalysis.Enabled = False
        
            cboAnalysisType.Enabled = True
            cmdAnalysis.Enabled = True
            txtAnalysisDescription.Enabled = True
            txtAnalysisDescription.text = ""
            cboAnalysisOption.Enabled = True
            txtAnalysisValue.Enabled = True
            txtAnalysisValue.text = ""
            txtAnalysisQuantity.Enabled = True
            txtAnalysisQuantity.text = ""
            txtAnalysisPercentage.Enabled = True
            txtAnalysisPercentage.text = ""
            
        Case ComponentMode.EditMode
            cmdAnalysisNew.Enabled = False
            cmdAnalysisEdit.Enabled = False
            cmdAnalysisSave.Enabled = True
            cmdAnalysisCancel.Enabled = True
            
            gridUnitAnalysis.Enabled = False
        
            cboAnalysisType.Enabled = True
            cmdAnalysis.Enabled = True
            txtAnalysisDescription.Enabled = True
            cboAnalysisOption.Enabled = True
            txtAnalysisValue.Enabled = True
            txtAnalysisQuantity.Enabled = True
            txtAnalysisPercentage.Enabled = True
    End Select
End Sub

Public Sub HealthSafetyButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control

   Select Case mode
      Case ComponentMode.DefaultMode
         cmdSafetyNew.Enabled = True
         cmdSafetyEdit.Enabled = False
         cmdSafetySave.Enabled = False
         cmdSafetyCancel.Enabled = False
         cmdAttachment.Enabled = False

         gridSafety.Enabled = True

         cboSafetyType.Enabled = False
         cboSafetyType.text = ""
         cmdSafety.Enabled = False
         txtRef.Enabled = False
         txtRef.text = ""
         txtDateChk.Enabled = False
         txtDateChk.text = ""
         txtNextDueDate.Enabled = False
         txtNextDueDate.text = ""
         cboSchedule.Enabled = False
         cboSchedule.text = ""
         cboInspectedBy.Enabled = False
         cboInspectedBy.text = ""
         txtSafetyTelephone.Enabled = False
         txtSafetyTelephone.text = ""
         txtComment.Enabled = False
         txtComment.text = ""
         chkCertificate.Enabled = False
         chkCertificate.Value = 0
         chkAlarm.Enabled = False
         chkAlarm.Value = 0

      Case ComponentMode.GridRowOnSelection
         cmdSafetyNew.Enabled = True
         cmdSafetyEdit.Enabled = True
         cmdSafetySave.Enabled = False
         cmdSafetyCancel.Enabled = False
         cmdAttachment.Enabled = False

         gridSafety.Enabled = True

      Case ComponentMode.NewEntryMode
         cmdSafetyNew.Enabled = False
         cmdSafetyEdit.Enabled = False
         cmdSafetySave.Enabled = True
         cmdSafetyCancel.Enabled = True
         cmdAttachment.Enabled = True

         gridSafety.Enabled = False

         cboSafetyType.Enabled = True
         cboSafetyType.text = ""
         cmdSafety.Enabled = True
         txtRef.Enabled = True
         txtRef.text = ""
         txtNextDueDate.Enabled = True
         txtNextDueDate.text = ""
         txtDateChk.Enabled = True
         txtDateChk.text = "" 'Format(Date, "dd/mm/yyyy")
         txtNextDueDate.text = ""
         cboSchedule.Enabled = True
         cboSchedule.text = ""
         cboInspectedBy.Enabled = True
         cboInspectedBy.text = ""
         txtSafetyTelephone.Enabled = True
         txtSafetyTelephone.text = ""
         txtComment.Enabled = True
         txtComment.text = ""
         chkCertificate.Enabled = True
         chkAlarm.Enabled = True
         chkCertificate.Value = 0
         chkAlarm.Value = 0

      Case ComponentMode.EditMode
         cmdSafetyNew.Enabled = False
         cmdSafetyEdit.Enabled = False
         cmdSafetySave.Enabled = True
         cmdSafetyCancel.Enabled = True
         cmdAttachment.Enabled = True

         gridSafety.Enabled = False

         cboSafetyType.Enabled = True
         cmdSafety.Enabled = True
         txtRef.Enabled = True
         txtNextDueDate.Enabled = True
         txtDateChk.Enabled = True
         cboSchedule.Enabled = True
         cboInspectedBy.Enabled = True
         txtSafetyTelephone.Enabled = True
         txtComment.Enabled = True
         chkCertificate.Enabled = True
         chkAlarm.Enabled = True
   End Select
End Sub

Public Sub InsuranceButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control
   Select Case mode
      Case ComponentMode.DefaultMode
         cmdInsuranceNew.Enabled = True
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = False
         cmdInsuranceCancel.Enabled = False

         gridInsurance.Enabled = True
         fraInsurance.Enabled = False

         cboInsurer.text = ""
         cboInsuranceType.text = ""
         txtPolicyNo.text = ""
         txtSumInsured.text = ""
         txtAnnualPR.text = ""
         txtStartDate.text = ""
         txtExpiryDate.text = ""
         cboUsage.text = ""
         txtTelephone.text = ""
         txtComments.text = ""

         cboInsurer.Enabled = True
         cboInsuranceType.Enabled = True
         txtPolicyNo.Enabled = True
         txtSumInsured.Enabled = True
         txtAnnualPR.Enabled = True
         txtStartDate.Enabled = True
         txtExpiryDate.Enabled = True
         cboUsage.Enabled = True
         txtTelephone.Enabled = True
         txtComments.Enabled = True

      Case ComponentMode.GridRowOnSelection
         cmdInsuranceNew.Enabled = True
         cmdInsuranceEdit.Enabled = True
         cmdInsuranceSave.Enabled = False
         cmdInsuranceCancel.Enabled = False

         gridInsurance.Enabled = True

      Case ComponentMode.NewEntryMode
         cmdInsuranceNew.Enabled = False
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = True
         cmdInsuranceCancel.Enabled = True

         gridInsurance.Enabled = False
         fraInsurance.Enabled = True

         cboInsurer.text = ""
         cboInsuranceType.text = ""
         txtPolicyNo.text = ""
         txtSumInsured.text = ""
         txtAnnualPR.text = ""
         txtStartDate.text = ""
         txtExpiryDate.text = ""
         cboUsage.text = ""
         txtTelephone.text = ""
         txtComments.text = ""

      Case ComponentMode.EditMode
         cmdInsuranceNew.Enabled = False
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = True
         cmdInsuranceCancel.Enabled = True

         gridInsurance.Enabled = False
         fraInsurance.Enabled = True
   End Select
End Sub

Public Sub UtilitiesButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control

   Select Case mode
      Case ComponentMode.DefaultMode
         cmdUtilitiesNew.Enabled = True
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = False
         cmdUtilitiesCancel.Enabled = False

         gridUtilities.Enabled = True

         cboUtilitiesType.Enabled = False
         cboUtilitiesType.text = ""
         cboAuthority_Supplier.Enabled = False
         cboAuthority_Supplier.text = ""
         txtUtilitiesReference.Enabled = False
         txtUtilitiesReference.text = ""
         cboUnitUtilityStatus.Enabled = False
         cboUnitUtilityStatus.text = ""
         txtUnitUtilityStDt.Enabled = False
         txtUnitUtilityStDt.text = ""
         txtDateVacated.Enabled = False
         txtDateVacated.text = ""
         txtChargeRate.Enabled = False
         txtChargeRate.text = ""
         txtUnitUtilityIniReading.Enabled = False
         txtUnitUtilityIniReading.text = ""
         txtFinalReading.Enabled = False
         txtFinalReading.text = ""
         txtUnitUtilityCom.Enabled = False
         txtUnitUtilityCom.text = ""

      Case ComponentMode.GridRowOnSelection
         cmdUtilitiesNew.Enabled = True
         cmdUtilitiesEdit.Enabled = True
         cmdUtilitiesSave.Enabled = False
         cmdUtilitiesCancel.Enabled = False

         gridUtilities.Enabled = True

      Case ComponentMode.NewEntryMode
         cmdUtilitiesNew.Enabled = False
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = True
         cmdUtilitiesCancel.Enabled = True

         gridUtilities.Enabled = False

         cboUtilitiesType.Enabled = True
         cboUtilitiesType.text = ""
         cboAuthority_Supplier.Enabled = True
         cboAuthority_Supplier.text = ""
         txtUtilitiesReference.Enabled = True
         txtUtilitiesReference.text = ""
         cboUnitUtilityStatus.Enabled = True
         cboUnitUtilityStatus.text = ""
         txtUnitUtilityStDt.Enabled = True
         txtUnitUtilityStDt.text = ""
         txtDateVacated.Enabled = True
         txtDateVacated.text = ""
         txtChargeRate.Enabled = True
         txtChargeRate.text = ""
         txtUnitUtilityIniReading.Enabled = True
         txtUnitUtilityIniReading.text = ""
         txtFinalReading.Enabled = True
         txtFinalReading.text = ""
         txtUnitUtilityCom.Enabled = True
         txtUnitUtilityCom.text = ""

      Case ComponentMode.EditMode
         cmdUtilitiesNew.Enabled = False
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = True
         cmdUtilitiesCancel.Enabled = True

         gridUtilities.Enabled = False

         cboUtilitiesType.Enabled = True
         cboAuthority_Supplier.Enabled = True
         txtUtilitiesReference.Enabled = True
         cboUnitUtilityStatus.Enabled = True
         txtUnitUtilityStDt.Enabled = True
         txtDateVacated.Enabled = True
         txtChargeRate.Enabled = True
         txtUnitUtilityIniReading.Enabled = True
         txtFinalReading.Enabled = True
         txtUnitUtilityCom.Enabled = True
    End Select
End Sub

Public Sub UnitMemoButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    
    Select Case mode
        Case ComponentMode.DefaultMode
            cmdUnitMemoEdit.Enabled = True
            cmdUnitMemoSave.Enabled = False
            cmdUnitMemoCancel.Enabled = False
            
            txtUnitMemo.Enabled = False
            Frame17.Enabled = False
            
        Case ComponentMode.GridRowOnSelection
            cmdUnitMemoEdit.Enabled = True
            cmdUnitMemoSave.Enabled = False
            cmdUnitMemoCancel.Enabled = False
            
            txtUnitMemo.Enabled = False
            Frame17.Enabled = False

        Case ComponentMode.NewEntryMode
            cmdUnitMemoEdit.Enabled = False
            cmdUnitMemoSave.Enabled = True
            cmdUnitMemoCancel.Enabled = True
            
            txtUnitMemo.Enabled = True
            Frame17.Enabled = True
                    
        Case ComponentMode.EditMode
            cmdUnitMemoEdit.Enabled = False
            cmdUnitMemoSave.Enabled = True
            cmdUnitMemoCancel.Enabled = True
            
            txtUnitMemo.Enabled = True
            Frame17.Enabled = True
    End Select
End Sub

Public Sub ConfigGridMaintenanceHistory(ByVal rstMHistory_ As ADODB.Recordset)
   Dim iColumn As Integer
   Dim oColumn As ADODB.Field

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = rstMHistory_.Fields.Count + 1

   For iColumn = 1 To 10
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
'   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn - 1).Left - 70
   gridMaintenanceHistory.ColWidth(6) = 0
   gridMaintenanceHistory.ColWidth(11) = 900
   For iColumn = 12 To rstMHistory_.Fields.Count
      gridMaintenanceHistory.ColWidth(iColumn) = 0
   Next iColumn

   iColumn = 0
   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.RowHeight(0) = 0
   For Each oColumn In rstMHistory_.Fields
      gridMaintenanceHistory.TextMatrix(0, iColumn) = oColumn.Name
      gridMaintenanceHistory.col = iColumn
      gridMaintenanceHistory.CellFontBold = True
      iColumn = iColumn + 1
   Next oColumn
End Sub

Public Sub SetGridUnitAnalysisHeader(ByVal conUnitAnalysis As ADODB.Connection)
   Dim rstUnitAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
      sSQLQuery_ = "SELECT UnitAnalysisID,AnalysisType, AnalysisDescription, AnalysisOption, " & _
      "AnalysisValue, AnalysisQuantity, AnalysisPercentage " & _
      "FROM UnitAnalysis " & _
      "WHERE UNITNUMBER = '" & txtUnitNo.text & "' "
                    
   rstUnitAnalysis_.Open sSQLQuery_, conUnitAnalysis, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   iRow = 1

   gridUnitAnalysis.Clear
   gridUnitAnalysis.Rows = 2
   gridUnitAnalysis.Cols = 7
   
   gridUnitAnalysis.ColWidth(0) = 0
   gridUnitAnalysis.ColWidth(1) = cboAnalysisType.Width + cmdAnalysis.Width
   gridUnitAnalysis.ColWidth(2) = txtAnalysisDescription.Width
   gridUnitAnalysis.ColWidth(3) = cboAnalysisOption.Width
   gridUnitAnalysis.ColWidth(4) = txtAnalysisValue.Width
   gridUnitAnalysis.ColWidth(5) = txtAnalysisQuantity.Width
   gridUnitAnalysis.ColWidth(6) = txtAnalysisPercentage.Width
   gridUnitAnalysis.RowHeight(0) = 0

   Dim oColumn As ADODB.Field
   Dim iColumn As Integer
   iColumn = 0

   gridUnitAnalysis.Cols = rstUnitAnalysis_.Fields.Count
   For Each oColumn In rstUnitAnalysis_.Fields
        gridUnitAnalysis.TextMatrix(0, iColumn) = oColumn.Name
        iColumn = iColumn + 1
   Next oColumn

   'SetMaintenanceHistoryControl

   SetControlStyle gridUnitAnalysis
   rstUnitAnalysis_.Close

   Set rstUnitAnalysis_ = Nothing
End Sub

Private Sub ConfigureGridSafety()
   Dim szHeader As String

   szHeader$ = "UnitSafetyID|SafetyType|Schedule|Ref|DateChk|NextDueDate|" & _
               "InspectedBy|<SafetyTelephone|<Comment|Alarm|Certificate|Attach"

   gridSafety.Clear
   gridSafety.FormatString = szHeader
   gridSafety.Rows = 2
   gridSafety.Cols = 12
   gridSafety.RowHeight(0) = 0

   gridSafety.ColWidth(0) = 0
   gridSafety.ColWidth(1) = Label41(1).Left - Label41(0).Left
   gridSafety.ColWidth(2) = Label41(2).Left - Label41(1).Left
   gridSafety.ColWidth(3) = Label41(3).Left - Label41(2).Left
   gridSafety.ColWidth(4) = Label41(4).Left - Label41(3).Left
   gridSafety.ColWidth(5) = Label41(5).Left - Label41(4).Left
   gridSafety.ColWidth(6) = Label41(6).Left - Label41(5).Left
   gridSafety.ColWidth(7) = Label41(7).Left - Label41(6).Left
   gridSafety.ColWidth(8) = VerticalLabel(0).Left - Label41(7).Left
   gridSafety.ColWidth(9) = VerticalLabel(1).Left - VerticalLabel(0).Left
   gridSafety.ColWidth(10) = VerticalLabel(2).Left - VerticalLabel(1).Left
   gridSafety.ColWidth(11) = gridSafety.Width + gridSafety.Left - VerticalLabel(2).Left - 300
End Sub

Private Sub ConfigureGridUtilities()
   Dim szHeader As String

   szHeader$ = "UnitUtilitiesID|Occupier|UtilitiesType|Authority_Supplier|UtilitiesReference|UnitUtilityStatus" & _
               "|UnitUtilityStDt|DateVacated|ChargeRate|UnitUtilityIniReading|FinalReading|UnitUtilityCom"

   gridUtilities.Clear
   gridUtilities.FormatString = szHeader
   gridUtilities.Rows = 2
   gridUtilities.Cols = 12
   gridUtilities.RowHeight(0) = 0

   gridUtilities.ColWidth(0) = 0
   gridUtilities.ColWidth(1) = Label82(1).Left - Label82(0).Left
   gridUtilities.ColWidth(2) = Label82(2).Left - Label82(1).Left
   gridUtilities.ColWidth(3) = Label82(3).Left - Label82(2).Left
   gridUtilities.ColWidth(4) = Label82(4).Left - Label82(3).Left
   gridUtilities.ColWidth(5) = Label82(5).Left - Label82(4).Left
   gridUtilities.ColWidth(6) = Label82(6).Left - Label82(5).Left
   gridUtilities.ColWidth(7) = Label82(7).Left - Label82(6).Left
   gridUtilities.ColWidth(8) = Label82(8).Left - Label82(7).Left
   gridUtilities.ColWidth(9) = Label82(9).Left - Label82(8).Left
   gridUtilities.ColWidth(10) = Label82(10).Left - Label82(9).Left
   gridUtilities.ColWidth(11) = gridUtilities.Width + gridUtilities.Left - Label82(10).Left - 300
End Sub

Public Sub LoadGridMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection)
   Dim rstMHistory_ As New ADODB.Recordset
   Dim szSQL As String
' Comment out by anol 20161121 view job was not working
'   szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY'), S.Value, " & _
'                "H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
'                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
'                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
'                "H.Detail, H.ActualCost, H.ReportedBy, " & _
'                "H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent, " & _
'                "H.MaintenanceType " & _
'           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
'           "WHERE H.UnitNumber = '" & txtUnitNo.text & "' " & _
'               "AND S.Code = H.MaintenanceType " & _
'               "AND S.PrimaryCode = 'MTYP' " & _
'           "ORDER BY H.ReportedDate DESC;"

'Debug.Print szSQL
szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID, '', P.PropertyName , '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S , Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE H.UnitNumber = '" & txtUnitNo.text & "' AND  U.UnitNumber = L.UnitNumber AND U.PropertyID= H.PropertyID AND  H.PropertyID = P.PropertyID AND H.ReportedBy=L.SageAccountNumber AND " & _
               "S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' " & _
           "ORDER BY H.ReportedDate DESC;"
           
   rstMHistory_.Open szSQL, conMHistory_, adOpenStatic, adLockReadOnly

   ConfigGridMaintenanceHistory rstMHistory_

   If rstMHistory_.EOF Then
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
      Exit Sub
   Else
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
   End If

   populateGridDefinedHeader conMHistory_, szSQL, gridMaintenanceHistory

   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.col = 0
End Sub

Public Sub PopulateGridUnitAnalysis(ByVal conUnitAnalysis As ADODB.Connection)

   Dim rstUnitAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
      sSQLQuery_ = "SELECT UnitAnalysis.UnitAnalysisID, " & _
      "UnitAnalysis.AnalysisType, " & _
      "UnitAnalysis.AnalysisDescription, " & _
      "UnitAnalysis.AnalysisOption, UnitAnalysis.AnalysisValue, " & _
      "UnitAnalysis.AnalysisQuantity, UnitAnalysis.AnalysisPercentage, " & _
      "SecondaryCode.value " & _
      "FROM UnitAnalysis, SecondaryCode " & _
      "WHERE UnitAnalysis.UNITNUMBER = '" & txtUnitNo.text & "' " & _
      "AND SecondaryCode.Code = UnitAnalysis.AnalysisType " & _
      "AND SecondaryCode.PrimaryCode = 'ATYP'"

   rstUnitAnalysis_.Open sSQLQuery_, conUnitAnalysis, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   Dim dAnalysisArea As Double
   iRow = 1

   gridUnitAnalysis.Clear
   gridUnitAnalysis.Rows = 2
   gridUnitAnalysis.Cols = 9

   SetGridUnitAnalysisHeader conUnitAnalysis

   While Not rstUnitAnalysis_.EOF
      gridUnitAnalysis.TextMatrix(iRow, 0) = rstUnitAnalysis_!UnitAnalysisID
      gridUnitAnalysis.TextMatrix(iRow, 1) = rstUnitAnalysis_!Value
      gridUnitAnalysis.TextMatrix(iRow, 2) = rstUnitAnalysis_!AnalysisDescription
      gridUnitAnalysis.TextMatrix(iRow, 3) = rstUnitAnalysis_!AnalysisOption
      gridUnitAnalysis.TextMatrix(iRow, 4) = rstUnitAnalysis_!AnalysisValue
      dAnalysisArea = dAnalysisArea + CDbl(rstUnitAnalysis_!AnalysisValue * rstUnitAnalysis_!AnalysisQuantity)
      gridUnitAnalysis.TextMatrix(iRow, 5) = rstUnitAnalysis_!AnalysisQuantity
      gridUnitAnalysis.TextMatrix(iRow, 6) = rstUnitAnalysis_!AnalysisPercentage

      rstUnitAnalysis_.MoveNext
      gridUnitAnalysis.AddItem ""
      iRow = iRow + 1
   Wend

   rstUnitAnalysis_.Close
   txtAnalysisTotalArea.text = dAnalysisArea
   Set rstUnitAnalysis_ = Nothing
End Sub

Public Sub PopulateHealthSafety(ByVal conSafety As ADODB.Connection)
   Dim rstSafety As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   sSQLQuery_ = "SELECT U.UnitSafetyID, U.SafetyInspection, U.Attachment, " & _
      "U.NextDueDate, U.SafetyStatus, U.SafetyTelephone, " & _
      "U.Certificate, U.LastDate, U.Comment, " & _
      "U.Alarm, S.Value AS SV, SC.Value AS SS, SE.Value AS InspectedBy " & _
   "FROM ((UnitSafety AS U LEFT JOIN SecondaryCode AS S ON U.SafetyType = S.Code) " & _
      "LEFT JOIN SecondaryCode AS SC ON U.SafetyStatus = SC.Code) " & _
      "LEFT JOIN SecondaryCode AS SE ON U.InspectedBy =  SE.Code " & _
   "WHERE U.UNITNUMBER = '" & txtUnitNo.text & "' AND Module = 'U';"
'Debug.Print sSQLQuery_
   rstSafety.Open sSQLQuery_, conSafety, adOpenDynamic, adLockOptimistic

   Dim iRow As Integer
   iRow = 1

   ConfigureGridSafety

   While Not rstSafety.EOF
      gridSafety.TextMatrix(iRow, 0) = rstSafety!UnitSafetyID
      gridSafety.TextMatrix(iRow, 1) = rstSafety!SV
      gridSafety.TextMatrix(iRow, 2) = IIf(IsNull(rstSafety!SS), "", rstSafety!SS)
      gridSafety.TextMatrix(iRow, 3) = IIf(IsNull(rstSafety!SafetyInspection), "", rstSafety!SafetyInspection)
      gridSafety.TextMatrix(iRow, 4) = IIf(IsNull(rstSafety!LastDate), "", Format(rstSafety!LastDate, "dd/mm/yyyy"))
      gridSafety.TextMatrix(iRow, 5) = IIf(IsNull(rstSafety!NextDueDate), "", rstSafety!NextDueDate)
      gridSafety.TextMatrix(iRow, 6) = IIf(IsNull(rstSafety!InspectedBy), "", rstSafety!InspectedBy)
      gridSafety.TextMatrix(iRow, 7) = IIf(IsNull(rstSafety!SafetyTelephone), "", rstSafety!SafetyTelephone)
      gridSafety.TextMatrix(iRow, 8) = IIf(IsNull(rstSafety!comment), "", rstSafety!comment)
      gridSafety.TextMatrix(iRow, 9) = IIf(IIf(IsNull(rstSafety!Alarm), "N", rstSafety!Alarm) = "Y", "Yes", "No")
      gridSafety.TextMatrix(iRow, 10) = IIf(IIf(IsNull(rstSafety!Certificate), "N", rstSafety!Certificate), "Yes", "No")
      gridSafety.TextMatrix(iRow, 11) = IIf(IIf(IsNull(rstSafety!attachment), "N", rstSafety!attachment) = "Y", "Yes", "No")

      rstSafety.MoveNext
      gridSafety.AddItem ""
      iRow = iRow + 1
   Wend

   rstSafety.Close
   Set rstSafety = Nothing
   gridSafety.row = 0
End Sub

Public Sub PopulateInsurance(ByVal conInsurance As ADODB.Connection)
   Dim rstInsurance As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next

   sSQLQuery_ = "SELECT I.*, " & _
                     "SC1.value AS Insu, SC2.value AS InsType, " & _
                     "SC3.value as U " & _
               "FROM ((PropertyInsurance AS I LEFT JOIN SecondaryCode AS SC1 ON " & _
                     "(I.Insurer = SC1.Code AND SC1.PrimaryCode = 'IRER')) " & _
                  "LEFT JOIN SecondaryCode AS SC2 ON (I.InsuranceType = SC2.Code AND SC2.PrimaryCode = 'ITYP')) " & _
                  "LEFT JOIN SecondaryCode AS SC3 ON (I.Usage = SC3.Code AND SC3.PrimaryCode = 'UUSE') " & _
               "WHERE I.PropertyID = '" & txtUnitNo.text & "' AND I.Module = 'U';"

'Debug.Print sSQLQuery_
   rstInsurance.Open sSQLQuery_, conInsurance, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   ConfigureGridInsurance

   iRow = 1

   While Not rstInsurance.EOF
      gridInsurance.TextMatrix(iRow, 0) = rstInsurance!PropertyInsuranceID
      gridInsurance.TextMatrix(iRow, 1) = IIf(IsNull(rstInsurance!Insu), "", rstInsurance!Insu)
      gridInsurance.TextMatrix(iRow, 2) = IIf(IsNull(rstInsurance!InsType), "", rstInsurance!InsType)
      gridInsurance.TextMatrix(iRow, 3) = rstInsurance!PolicyNo
      gridInsurance.TextMatrix(iRow, 4) = IIf(IsNull(rstInsurance!SumInsured), "0.00", Format(rstInsurance!SumInsured, "0.00"))
      gridInsurance.TextMatrix(iRow, 5) = IIf(IsNull(rstInsurance!AnnualPR), "0.00", Format(rstInsurance!AnnualPR, "0.00"))
      gridInsurance.TextMatrix(iRow, 6) = IIf(IsNull(rstInsurance!StartDate), "", Format(rstInsurance!StartDate, "dd/mm/yyyy"))
      gridInsurance.TextMatrix(iRow, 7) = IIf(IsNull(rstInsurance!ExpiryDate), "", rstInsurance!ExpiryDate)
      gridInsurance.TextMatrix(iRow, 8) = IIf(IsNull(rstInsurance!Telephone), "", rstInsurance!Telephone)
      gridInsurance.TextMatrix(iRow, 9) = IIf(IsNull(rstInsurance!u), "", rstInsurance!u)
      gridInsurance.TextMatrix(iRow, 10) = IIf(IsNull(rstInsurance!Comments), "", rstInsurance!Comments)
      gridInsurance.TextMatrix(iRow, 11) = IIf(IIf(IsNull(rstInsurance!attachment), "N", rstInsurance!attachment) = "Y", "Yes", "No")
      rstInsurance.MoveNext
      gridInsurance.AddItem ""
      iRow = iRow + 1
   Wend

   rstInsurance.Close
   Set rstInsurance = Nothing
   gridInsurance.row = 0
   gridInsurance.col = 0
End Sub

Public Sub ConfigureGridInsurance()
   Dim szHeader As String

   szHeader$ = "PropertyInsuranceID|<Insurer|<InsuranceType|<" & _
               "PolicyNo|>SumInsured|>AnnualPR|<StartDate|<ExpiryDate|<" & _
               "Telephone|<Usage|<Comments|<Attachment"
                  
   gridInsurance.Clear
   gridInsurance.FormatString = szHeader
   gridInsurance.Rows = 2
   gridInsurance.Cols = 12
   gridInsurance.RowHeight(0) = 0

   gridInsurance.ColWidth(0) = 0                                        'ID
   gridInsurance.ColWidth(1) = Label6(1).Left - Label6(0).Left          'Insurer
   gridInsurance.ColWidth(2) = Label6(2).Left - Label6(1).Left          'Ins Type
   gridInsurance.ColWidth(3) = Label6(3).Left - Label6(2).Left          'Policy No
   gridInsurance.ColWidth(4) = Label6(4).Left - Label6(3).Left          'Sum Ins
   gridInsurance.ColWidth(5) = Label6(5).Left - Label6(4).Left          'Annual PR
   gridInsurance.ColWidth(6) = Label6(6).Left - Label6(5).Left          'St Date
   gridInsurance.ColWidth(7) = Label6(7).Left - Label6(6).Left          'Exp Date
   gridInsurance.ColWidth(8) = Label6(8).Left - Label6(7).Left          'Tel
   gridInsurance.ColWidth(9) = Label6(9).Left - Label6(8).Left          'Usage
   gridInsurance.ColWidth(10) = Label6(10).Left - Label6(9).Left        'Comment
   gridInsurance.ColWidth(11) = gridInsurance.Width + gridInsurance.Left - (Label6(10).Left + fraInsurance.Left) - 200  'Attach
End Sub

Public Sub PopulateUtilities(ByVal conUtilities As ADODB.Connection)
   Dim rstUtilities As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim iRow As Integer

   'On Error Resume Next
   sSQLQuery_ = _
         "SELECT U.*, " & _
              "SC1.value AS UT, SC2.value AS US, S.SupplierName " & _
         "FROM ((UnitUtilities AS U LEFT JOIN SecondaryCode AS SC1 ON U.UtilitiesType = SC1.Code) " & _
              "LEFT JOIN SecondaryCode AS SC2 ON U.Status = SC2.Code) INNER JOIN " & _
              "Supplier AS S ON U.Authority_Supplier = S.SupplierID " & _
         "WHERE U.UNITNUMBER = '" & txtUnitNo.text & "' " & _
              "AND SC1.PrimaryCode = 'UTIL' " & _
              "AND SC2.PrimaryCode = 'USTA';"
'Debug.Print sSQLQuery_
   rstUtilities.Open sSQLQuery_, conUtilities, adOpenStatic, adLockReadOnly

   iRow = 1

   ConfigureGridUtilities

   While Not rstUtilities.EOF
      gridUtilities.TextMatrix(iRow, 0) = IIf(IsNull(rstUtilities!UnitUtilitiesID), "", rstUtilities!UnitUtilitiesID)
      gridUtilities.TextMatrix(iRow, 1) = IIf(IsNull(rstUtilities!Occupier), "", rstUtilities!Occupier)
      gridUtilities.TextMatrix(iRow, 2) = IIf(IsNull(rstUtilities!UT), "", rstUtilities!UT)
      gridUtilities.TextMatrix(iRow, 3) = IIf(IsNull(rstUtilities!Authority_Supplier), "", rstUtilities!SupplierName)
      gridUtilities.TextMatrix(iRow, 4) = IIf(IsNull(rstUtilities!UtilitiesReference), "", rstUtilities!UtilitiesReference)
      gridUtilities.TextMatrix(iRow, 5) = IIf(IsNull(rstUtilities!US), "", rstUtilities!US)
      gridUtilities.TextMatrix(iRow, 6) = IIf(IsNull(rstUtilities!StartDate), "", Format(rstUtilities!StartDate, "dd/mm/yyyy"))
      gridUtilities.TextMatrix(iRow, 7) = IIf(IsNull(rstUtilities!DateVacated), "", Format(rstUtilities!DateVacated, "dd/mm/yyyy"))
      gridUtilities.TextMatrix(iRow, 8) = IIf(IsNull(rstUtilities!ChargeRate), "", Format(rstUtilities!ChargeRate, "0.00"))
      gridUtilities.TextMatrix(iRow, 9) = IIf(IsNull(rstUtilities!InitialReading), "", rstUtilities!InitialReading)
      gridUtilities.TextMatrix(iRow, 10) = IIf(IsNull(rstUtilities!FinalReading), "", rstUtilities!FinalReading)
      gridUtilities.TextMatrix(iRow, 11) = IIf(IsNull(rstUtilities!Comments), "", rstUtilities!Comments)

      rstUtilities.MoveNext
      gridUtilities.AddItem ""
      iRow = iRow + 1
   Wend

   rstUtilities.Close
   Set rstUtilities = Nothing

   gridUtilities.row = 0
End Sub
'
'Public Function SaveUnitMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection) As Boolean
'    Dim rstMHistory_ As New ADODB.Recordset
'    Dim rstDEL_MHistory As New ADODB.Recordset
'    Dim rstID         As New ADODB.Recordset
'    Dim sSQLQuery_ As String
'    Dim sSQLDelete As String
'    Dim sSQLFilter As String
'    Dim iRowIndex As Integer
'    Dim lTableID      As Long
'
'    sSQLFilter = ""
'
'    On Error GoTo Exception
'
'    If Not M_HISTORY_NEW_ENTRY_ Then
'        sSQLFilter = "WHERE UNITNUMBER = '" & txtUnitNo.text & "' AND ID = " & txtID.text & ""
'    Else
'        sSQLFilter = ""
'    End If
'
'    sSQLQuery_ = "SELECT ID, UnitNumber, MaintenanceType, ReportedDate, Description, " & _
'                    "DateCompleted, TaskOwner, Contact, RemindDate, Alarm, EstimateCost, REMINDER_ID " & _
'                "FROM UNITMAINTHISTORY " & sSQLFilter
'
'   rstMHistory_.Open sSQLQuery_, conMHistory_, adOpenDynamic, adLockOptimistic
'
'    'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
'    If M_HISTORY_NEW_ENTRY_ Then
'       sSQLQuery_ = "SELECT MAX(ID) AS M_ID FROM UnitMAINTHISTORY;"
'       rstID.Open sSQLQuery_, conMHistory_, adOpenDynamic, adLockOptimistic
'       lTableID = IIf(IsNull(rstID!M_ID), 0, rstID!M_ID) + 1
'       rstID.Close
'       Set rstID = Nothing
'
'       rstMHistory_.AddNew
'    End If
'
'    rstMHistory_!UnitNumber = txtUnitNo.text
'    rstMHistory_!MaintenanceType = cboMaintenanceType.BoundText
'    rstMHistory_!ReportedDate = IIf(dtpReportedDate.text = "", Format(Date, "dd mmmm yyyy"), Format(dtpReportedDate.text, "dd mmmm yyyy"))
'    rstMHistory_!description = IIf(txtDescription.text = "", "", txtDescription.text)
'    rstMHistory_!DateCompleted = IIf(dtpDateCompleted.text = "", Null, Format(dtpDateCompleted.text, "dd mmmm yyyy"))
'    rstMHistory_!TaskOwner = IIf(txtTaskOwner.text = "", "", txtTaskOwner.text)
'    rstMHistory_!Contact = IIf(txtContact.text = "", "", txtContact.text)
'    rstMHistory_!RemindDate = IIf(dtpRemindDate.text = "", Null, Format(dtpRemindDate.text, "dd mmmm yyyy"))
'
'    If chkAlarm.Value = 1 And dtpRemindDate.text <> "" Then
'      rstMHistory_!Alarm = True
'
'      If M_HISTORY_NEW_ENTRY_ Then
'         rstMHistory_!Reminder_ID = NewReminder(Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), "083000", txtDescription.text, "UnitMAINTHISTORY", CStr(lTableID))
'      Else
'         If IsNull(rstMHistory_!Reminder_ID) Then
'            rstMHistory_!Reminder_ID = NewReminder(Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), "083000", txtDescription.text, "UnitMAINTHISTORY", CStr(lTableID))
'         Else
'            UpdateReminder rstMHistory_!Reminder_ID, Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), "083000", txtDescription.text
'         End If
'      End If
'   Else
'      rstMHistory_!Alarm = False
'
'      If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.Row, 8) = "Y" Then
'         ClearReminder rstMHistory_!Reminder_ID
'      End If
'   End If
'
'    rstMHistory_.Update
'
'    rstMHistory_.Close
'
'    Set rstMHistory_ = Nothing
'
'    SaveUnitMaintenanceHistory = True
'    Exit Function
'
'Exception:
'
'    'MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
'    rstMHistory_.Close
'    conMHistory_.Close
'    Set rstMHistory_ = Nothing
'    Set conMHistory_ = Nothing
'    SaveUnitMaintenanceHistory = False
'End Function

Public Function SaveUnitAnalysis(ByVal conUnitAnalysis As ADODB.Connection) As Boolean
   Dim rstUnitAnalysis_ As New ADODB.Recordset
   Dim rstUnit_ As New ADODB.Recordset

   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   If Not UNIT_ANALYSIS_NEW_ENTRY Then
       sSQLFilter = "WHERE UNITNUMBER = '" & txtUnitNo.text & "' AND UnitAnalysisID = " & txtUnitAnalysisID.text & ""
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT UnitAnalysisID, UnitNumber, AnalysisType, AnalysisDescription, " & _
                  "AnalysisOption, AnalysisValue, AnalysisQuantity, AnalysisPercentage " & _
                "FROM UNITANALYSIS " & sSQLFilter

   rstUnitAnalysis_.Open sSQLQuery_, conUnitAnalysis, adOpenDynamic, adLockOptimistic

   If UNIT_ANALYSIS_NEW_ENTRY Then
       rstUnitAnalysis_.AddNew
   End If

   rstUnitAnalysis_!UnitNumber = txtUnitNo.text
   rstUnitAnalysis_!AnalysisType = IIf(cboAnalysisType.BoundText <> "", cboAnalysisType.BoundText, "0")
   rstUnitAnalysis_!AnalysisDescription = IIf(txtAnalysisDescription.text <> "", txtAnalysisDescription.text, "")
   rstUnitAnalysis_!AnalysisOption = IIf(cboAnalysisOption.BoundText <> "", cboAnalysisOption.BoundText, "")
   rstUnitAnalysis_!AnalysisValue = IIf(txtAnalysisValue.text <> "", txtAnalysisValue.text, "0")
   rstUnitAnalysis_!AnalysisQuantity = IIf(txtAnalysisQuantity.text <> "", txtAnalysisQuantity.text, "0")
   rstUnitAnalysis_!AnalysisPercentage = IIf(txtAnalysisPercentage.text <> "", txtAnalysisPercentage.text, "0")

   rstUnitAnalysis_.Update

   rstUnitAnalysis_.Close

   Set rstUnit_ = Nothing
   Set rstUnitAnalysis_ = Nothing

   SaveUnitAnalysis = True
End Function

Public Function SaveHealthSafety(ByVal conSafety As ADODB.Connection) As Boolean
   Dim rstSafety As New ADODB.Recordset
   Dim sSQLQuery_ As String, sSQLDelete As String
   Dim sSQLFilter As String, iRowIndex As Integer

   sSQLFilter = ""

   'On Error GoTo Exception

   If Not HEALTH_SAFETY_NEW_ENTRY Then
       sSQLFilter = "WHERE UNITNUMBER = '" & txtUnitNo.text & "' AND UnitSafetyID = '" & txtUnitSafetyID.text & "' AND Module = 'U'"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT U.UnitSafetyID, U.Attachment, " & _
                  "U.SafetyInspection, U.UnitNumber, " & _
                  "U.NextDueDate, U.SafetyStatus, " & _
                  "U.InspectedBy, U.SafetyTelephone, " & _
                  "U.Certificate, U.LastDate, U.Comment, " & _
                  "U.Alarm, U.SafetyType, U.Module, U.spare1 " & _
                "FROM UnitSafety AS U " & sSQLFilter

   rstSafety.Open sSQLQuery_, conSafety, adOpenDynamic, adLockOptimistic

   If HEALTH_SAFETY_NEW_ENTRY Then
      rstSafety.AddNew

      rstSafety!UnitSafetyID = UniqueID()

      If chkAlarm.Value = 1 Then
         rstSafety!spare1 = NewReminder(Format(txtNextDueDate.text, "YYYYMMDD"), "010000", _
                           "Health and safety issue for the unit no. " & txtUnitNo.text, _
                           "UnitSafety", rstSafety!UnitSafetyID)
         rstSafety!Alarm = "Y"
      End If
   Else
'      rstSafety!UnitSafetyID = HEALTH_SAFETY_ID
      If rstSafety!Alarm = "Y" And chkAlarm.Value = 0 Then _
         ClearReminder rstSafety!spare1
      If rstSafety!Alarm = "Y" And chkAlarm.Value = 1 Then _
         UpdateReminder rstSafety!spare1, Format(txtNextDueDate.text, "YYYYMMDD"), "010000", _
                        "Health and safety issue for the unit no. " & txtUnitNo.text
   End If
   rstSafety!UnitNumber = txtUnitNo.text
   rstSafety!SafetyType = cboSafetyType.BoundText
   rstSafety!SafetyInspection = txtRef.text
   rstSafety!NextDueDate = IIf(txtNextDueDate.text = "", Null, Format(txtNextDueDate.text, "dd mmmm yyyy"))
   rstSafety!SafetyStatus = cboSchedule.BoundText
   rstSafety!InspectedBy = cboInspectedBy.BoundText
   rstSafety!SafetyTelephone = txtSafetyTelephone.text
   rstSafety!Certificate = IIf(chkCertificate.Value = 1, True, False)
   rstSafety!LastDate = IIf(txtDateChk.text = "", Format(txtDateChk, "dd mmmm yyyy"), Format(txtDateChk.text, "dd mmmm yyyy"))
   rstSafety!comment = txtComment.text
   rstSafety!attachment = IIf(HEALTH_N_SAFETY_ATTACH, "Y", "N")
   rstSafety!Module = "U"
   rstSafety.Update

   rstSafety.Close
   Set rstSafety = Nothing

   SaveHealthSafety = True
   Exit Function

Exception:
   'MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
   rstSafety.Close
   conSafety.Close
   Set rstSafety = Nothing
   Set conSafety = Nothing
   SaveHealthSafety = False
End Function

Public Function SaveUnitUtilities(ByVal conUtilities As ADODB.Connection) As Boolean
   If cboUtilitiesType.text = "" Then
      ShowMsgInTaskBar "Please select utility type to save."
      Exit Function
   End If
   If cboAuthority_Supplier.text = "" Then
      ShowMsgInTaskBar "Please select supplier to save."
      Exit Function
   End If
   If cboUnitUtilityStatus.text = "" Then
      ShowMsgInTaskBar "Please select the status of the utility to save."
      Exit Function
   End If

   Dim rstUtilities As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   'On Error GoTo Exception

   If Not UNIT_UTILITIES_NEW_ENTRY Then
       sSQLFilter = "WHERE UNITNUMBER = '" & txtUnitNo.text & "' AND " & _
                          "UnitUtilitiesID = " & txtUnitUtilitiesID.text & " AND " & _
                          "Module = 'U';"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM UnitUtilities " & sSQLFilter

   rstUtilities.Open sSQLQuery_, conUtilities, adOpenDynamic, adLockOptimistic

   If UNIT_UTILITIES_NEW_ENTRY Then
      rstUtilities.AddNew
      rstUtilities!Occupier = lblTenantSageAcc.Caption
   End If

   rstUtilities!UnitNumber = txtUnitNo.text
   rstUtilities!UtilitiesType = cboUtilitiesType.BoundText
   rstUtilities!Authority_Supplier = cboAuthority_Supplier.BoundText
   rstUtilities!UtilitiesReference = txtUtilitiesReference.text
   rstUtilities!Status = cboUnitUtilityStatus.BoundText
   rstUtilities!StartDate = IIf(txtUnitUtilityStDt.text = "", Null, Format(txtUnitUtilityStDt.text, "dd mmmm yyyy"))
   rstUtilities!DateVacated = IIf(txtDateVacated.text = "", Null, Format(txtDateVacated.text, "dd mmmm yyyy"))
   rstUtilities!ChargeRate = IIf(txtChargeRate.text = "", "0", txtChargeRate.text)
   rstUtilities!InitialReading = txtUnitUtilityIniReading.text
   rstUtilities!FinalReading = txtFinalReading.text
   rstUtilities!Comments = txtUnitUtilityCom.text
   rstUtilities!Module = "U"
   rstUtilities.Update

   'Next iRowIndex
   rstUtilities.Close

   Set rstUtilities = Nothing

   SaveUnitUtilities = True
   Exit Function

Exception:

   'MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
   rstUtilities.Close

   Set rstUtilities = Nothing

   SaveUnitUtilities = False
End Function

Public Function SaveUnitMemo() As Boolean
    Dim conUnitMemo_ As New ADODB.Connection
    Dim rstUnitMemo_ As New ADODB.Recordset
    Dim sSQLQuery_, sSQLFilter As String

    'On Error GoTo Exception
    
    'Set the RDO Connections to the dataset
    conUnitMemo_.Open getConnectionString

    sSQLFilter = "WHERE UNITNUMBER = '" & txtUnitNo.text & "'"

    sSQLQuery_ = "SELECT UnitNumber, PropertyID, UnitName, UnitAddressLine1, " & _
               "UnitAddressLine2, UnitAddressLine3, UnitAddressLine4, " & _
               "UnitPostCode, Occupied, TenantCompanyName, SageAccountNumber, " & _
               "Frontage, RateableValue, RatesPayable, GroundFloorArea, " & _
               "MezzanineArea, TotalArea, UnitType, LandLord, Management, Memo, RentalPrice " & _
    "FROM Units " & sSQLFilter

    rstUnitMemo_.Open sSQLQuery_, conUnitMemo_, adOpenDynamic, adLockOptimistic

    If txtUnitMemo.text = "" Then
        rstUnitMemo_!Memo = "<No memo saved>"
    Else
        rstUnitMemo_!Memo = txtUnitMemo.text
    End If
    rstUnitMemo_.Update
    
    rstUnitMemo_.Close
    conUnitMemo_.Close
    Set rstUnitMemo_ = Nothing
    Set conUnitMemo_ = Nothing
    SaveUnitMemo = True
    Exit Function

Exception:
    
    ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
    rstUnitMemo_.Close
    conUnitMemo_.Close
    Set rstUnitMemo_ = Nothing
    Set conUnitMemo_ = Nothing
    SaveUnitMemo = False
End Function

Public Function SaveUnitInsurance(ByVal conInsurance As ADODB.Connection) As Boolean
   If cboInsurer.text = "" Then
       ShowMsgInTaskBar "Please enter the name of the Insurer."
       SaveUnitInsurance = False
       Exit Function
   End If

   If txtPolicyNo.text = "" Then
       ShowMsgInTaskBar "Please enter the Policy No."
       SaveUnitInsurance = False
       Exit Function
   End If

   Dim rstInsurance As New ADODB.Recordset
   Dim rstDEL_MHistory As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   'On Error GoTo Exception
   'Set the RDO Connections to the dataset

   If Not UNIT_INSURANCE_NEW_ENTRY Then
       sSQLFilter = "WHERE PropertyID = '" & txtUnitNo.text & "' AND " & _
                        "PropertyInsuranceID = '" & txtPropertyInsuranceID.text & "' AND " & _
                        "Module = 'U';"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyInsurance " & sSQLFilter
'Debug.Print sSQLQuery_
   rstInsurance.Open sSQLQuery_, conInsurance, adOpenDynamic, adLockOptimistic

   'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
   If UNIT_INSURANCE_NEW_ENTRY Then rstInsurance.AddNew
   If INSURANCE_ID = "" Then
      rstInsurance!PropertyInsuranceID = UniqueID()
   Else
      rstInsurance!PropertyInsuranceID = INSURANCE_ID
   End If

   rstInsurance!propertyID = txtUnitNo.text
   rstInsurance!Insurer = cboInsurer.BoundText
   rstInsurance!InsuranceType = cboInsuranceType.BoundText
   rstInsurance!PolicyNo = txtPolicyNo.text
   rstInsurance!SumInsured = IIf(txtSumInsured.text = "", "0", txtSumInsured.text)
   rstInsurance!AnnualPR = IIf(txtAnnualPR.text = "", "0", txtAnnualPR.text)
   rstInsurance!StartDate = IIf(txtStartDate.text = "", Null, Format(txtStartDate.text, "dd mmmm yyyy"))
   rstInsurance!ExpiryDate = IIf(txtExpiryDate.text = "", Null, Format(txtExpiryDate.text, "dd mmmm yyyy"))
   rstInsurance!Telephone = txtTelephone.text
   rstInsurance!Usage = cboUsage.BoundText
   rstInsurance!Comments = txtComments.text
   rstInsurance!Module = "U"
   rstInsurance!attachment = IIf(HEALTH_N_SAFETY_ATTACH, "Y", "N")

   rstInsurance.Update
   rstInsurance.Close
   Set rstInsurance = Nothing

   SaveUnitInsurance = True
   Exit Function

Exception:

   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
   rstInsurance.Close

   Set rstInsurance = Nothing

   SaveUnitInsurance = False
End Function

Private Sub AccountHistory(ByVal Conn1 As ADODB.Connection)
   Dim rstBP As New ADODB.Recordset
   Dim szBP As String, i As Integer

   FlexGridAccountHistoryConfigure

'   szBP = "SELECT LeaseDetails.CompanyName, LeaseDetails.SageAccountNumber, " & _
'                  "LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.TerminateDate, " & _
'                  "LeaseDetails.OLED, PropertyInsurance.Usage " & _
'          "FROM LeaseDetails, PropertyInsurance " & _
'          "WHERE LeaseDetails.Status = False AND " & _
'               "LeaseDetails.LeaseID = PropertyInsurance.LeaseID AND " & _
'               "LeaseDetails.UnitNumber = '" & txtUnitNo.text & "' AND " & _
'               "PropertyInsurance.PropertyInsuranceID IN " & _
'               "(SELECT MAX(PropertyInsurance.PropertyInsuranceID) " & _
'                "FROM PropertyInsurance, LeaseDetails  " & _
'                "GROUP BY PropertyInsurance.LeaseID" & _
'               ");"
    szBP = "SELECT LeaseDetails.CompanyName, LeaseDetails.SageAccountNumber, " & _
               "LeaseDetails.StartDate, LeaseDetails.EndDate, LeaseDetails.TerminateDate, " & _
               "LeaseDetails.OLED, LeaseDetails.Usage " & _
           "FROM LeaseDetails " & _
               "WHERE LeaseDetails.Status = False AND " & _
               "LeaseDetails.UnitNumber = '" & txtUnitNo.text & "';"

'Debug.Print szBP
   rstBP.Open szBP, Conn1, adOpenStatic, adLockReadOnly

   Dim iRow As Integer

   iRow = 1

   While Not rstBP.EOF
      'gridACHistory.RowHeight(gridACHistory.Rows - 1) = 285
      gridACHistory.TextMatrix(iRow, 0) = IIf(IsNull(rstBP!SageAccountNumber), "", rstBP!SageAccountNumber)
      gridACHistory.TextMatrix(iRow, 1) = IIf(IsNull(rstBP!CompanyName), "", rstBP!CompanyName)
      gridACHistory.TextMatrix(iRow, 2) = IIf(IsNull(rstBP!StartDate), "", rstBP!StartDate)

      If (rstBP!OLED = "False") Then
         gridACHistory.TextMatrix(iRow, 3) = IIf(IsNull(rstBP!EndDate), "", rstBP!EndDate)
      Else
         gridACHistory.TextMatrix(iRow, 3) = IIf(IsNull(rstBP!TerminateDate), "", rstBP!TerminateDate)
      End If

      gridACHistory.TextMatrix(iRow, 4) = IIf(IsNull(rstBP!Usage), "", rstBP!Usage)

      rstBP.MoveNext
      iRow = iRow + 1
      If Not rstBP.EOF Then gridACHistory.AddItem ""
   Wend

   rstBP.Close
   Set rstBP = Nothing

   gridACHistory.Sort = flexSortGenericAscending
   gridACHistory.ColAlignment(0) = vbAlignLeft
   gridACHistory.ColAlignment(1) = vbLeftJustify
   gridACHistory.ColAlignment(4) = vbLeftJustify
End Sub

Private Sub FlexGridAccountHistoryConfigure()
   gridACHistory.Clear
   gridACHistory.Rows = 2
   gridACHistory.Cols = 5
   
   gridACHistory.ColWidth(0) = 1500
   gridACHistory.TextMatrix(0, 0) = "A/C Code"

   gridACHistory.ColWidth(1) = 3000
   gridACHistory.TextMatrix(0, 1) = "Tenant Name"

   gridACHistory.ColWidth(2) = 1400
   gridACHistory.TextMatrix(0, 2) = "Start Date"

   gridACHistory.ColWidth(3) = 1400
   gridACHistory.TextMatrix(0, 3) = "End Date"

   gridACHistory.ColWidth(4) = 4500
   gridACHistory.TextMatrix(0, 4) = "Usage"

'   lblMainUnit(14).Left = 400
'   lblMainUnit(15).Left = 2600
'   lblMainUnit(16).Left = 5000
'   lblMainUnit(17).Left = 6500
'   lblMainUnit(18).Left = 9000

   gridACHistory.RowHeight(0) = 0
End Sub

Private Function SetAlarm() As Boolean
   Dim conAlarm_ As New ADODB.Connection
   Dim rstAlarm_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'Set the RDO Connections to the dataset
   conAlarm_.Open "DSN=" & DSN_ALARM_ & ";UID=;PWD="

   sSQLQuery_ = "SELECT * " & _
             "FROM tlbReminder"

   rstAlarm_.Open sSQLQuery_, conAlarm_, adOpenStatic, adLockReadOnly

   If rstAlarm_.EOF Or rstAlarm_.BOF Then
       If (AddUnitInformation) Then
           ShowMsgInTaskBar "The Unit Information added successfully."
       Else
           ShowMsgInTaskBar "An error occured while saving the Unit Information.", , "N"
       End If
   Else
       If (MsgBox("WARNING ! The Unit Number entered already exists. Do you want to update the information", vbYesNo, "Save Unit Information") = vbYes) Then
           UpdateUnitInformation
       End If
   End If

   rstAlarm_.Close
   conAlarm_.Close
   Set rstAlarm_ = Nothing
   Set conAlarm_ = Nothing
   SetAlarm = True
End Function

Public Function GenerateUnitNumber() As String
   If txtPropertyName.Tag = "" Then Exit Function
   
   Dim conUnitNumber As New ADODB.Connection
   Dim rstUnitNumber As New ADODB.Recordset
   Dim adoUnitNumber As ADODB.Recordset
   Dim sSQLQuery_ As String, bUnitNum As Boolean
   Dim MAX_UNIT_ As String
   Dim UNIT_NUMBER_ As String

   'Set the RDO Connections to the dataset
   conUnitNumber.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT MAX(RIGHT(Units.UnitNumber,3)) + 1 AS  MAX_UNIT " & _
                "From Units " & _
                "WHERE Units.PropertyID = '" & txtPropertyName.Tag & "';"
'Debug.Print sSQLQuery_
   Set adoUnitNumber = New ADODB.Recordset
   adoUnitNumber.Open sSQLQuery_, conUnitNumber, adOpenDynamic, adLockReadOnly
   
   If adoUnitNumber.EOF = True Then
      MAX_UNIT_ = "1"
   Else
      While Not adoUnitNumber.EOF
         MAX_UNIT_ = IIf(IsNull(adoUnitNumber!MAX_UNIT), "1", adoUnitNumber!MAX_UNIT)
         adoUnitNumber.MoveNext
      Wend
   End If
   adoUnitNumber.Close
   bUnitNum = False
   Do
      GenerateUnitNumber = txtPropertyName.Tag & "-" & Lpad(MAX_UNIT_, "0", 3)
      sSQLQuery_ = "SELECT Units.UnitNumber " & _
                   "From Units " & _
                   "WHERE Units.UnitNumber = '" & GenerateUnitNumber & "';"
      adoUnitNumber.Open sSQLQuery_, conUnitNumber, adOpenStatic, adLockReadOnly

      If adoUnitNumber.EOF Then
         bUnitNum = True
      Else
         adoUnitNumber.Close
         MAX_UNIT_ = MAX_UNIT_ + 1
      End If
   Loop Until bUnitNum

   adoUnitNumber.Close
   conUnitNumber.Close
   Set adoUnitNumber = Nothing
   Set conUnitNumber = Nothing
End Function

Public Function SetTotalArea()
   Dim iCount, iQuantity As Integer
   Dim dTotalArea, dValue As Double

   Dim conUnit_ As New ADODB.Connection
   Dim rstUnit_ As New ADODB.Recordset

   dTotalArea = 0
   For iCount = 1 To gridUnitAnalysis.Rows - 2
      dValue = CDbl(gridUnitAnalysis.TextMatrix(iCount, 4))
      iQuantity = CInt(gridUnitAnalysis.TextMatrix(iCount, 5))
      dTotalArea = dTotalArea + (dValue * iQuantity)
   Next iCount

   txtAnalysisTotalArea.text = dTotalArea

'** Samrat 06/06/2006 ****
'* I have stopped the following codes to modify the Unit form.
'* The total area of the unit should not be updated from unit analysis.
'* User should input the total area of the unit in the global area of
'* unit information. My concern is user might not know the area of each room,
'* but user should know the total area of the unit.
'********************************************************************
'   Dim sSQLQuery_ As String
'
'   sSQLQuery_ = "SELECT * FROM UNITS WHERE UNITNUMBER = '" & txtUnitNo.text & "'"
'
'   conUnit_.open  getConnectionString
'
'   Set rstUnit_ = conUnit_.OpenResultset(sSQLQuery_,adOpenDynamic,adLockOptimistic
'
'   rstUnit_.Edit
'   rstUnit_!TotalArea = IIf(txtAnalysisTotalArea.text <> "", txtAnalysisTotalArea.text, 0)
'   rstUnit_.Update
'
'   rstUnit_.Close
'   conUnit_.Close
'   Set rstUnit_ = Nothing
'   Set conUnit_ = Nothing
End Function

Private Sub txtUnitName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
'        If cmdUnitLookup.Enabled Then
'            cmdUnitLookup.SetFocus
'        Else
'            If txtUnitNo.Enabled Then txtUnitNo.SetFocus
'        End If
         txtUnitAddress1.SetFocus
    End If
End Sub

Private Sub txtUnitNo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            If txtUnitName.Enabled Then txtUnitName.SetFocus
        End If
        
        '''''''''''''''''Modification by Md. Mahboob 20230317 Change ID 2/Work Item 4-:  To stop allow space
    If Chr(KeyAscii) = " " Then
        KeyAscii = 0 'cancel the space character
    End If
    ''''''''End of modification
End Sub

Private Sub txtUnitNo_LostFocus()
   If txtUnitNo.text <> "" Then txtUnitNo.text = UCase(txtUnitNo.text)

   Dim conUnitNumber As New ADODB.Connection
   Dim rstUnitNumber As New ADODB.Recordset
   Dim sSQLQuery_    As String
   Dim bUnitID       As Boolean

   'Set the RDO Connections to the dataset
   conUnitNumber.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT UnitNumber " & _
                "FROM   Units " & _
                "WHERE  UnitNumber = '" & txtUnitNo.text & "';"
'Debug.Print sSQLQuery_
   rstUnitNumber.Open sSQLQuery_, conUnitNumber, adOpenStatic, adLockReadOnly

   bUnitID = Not rstUnitNumber.EOF

   rstUnitNumber.Close
   conUnitNumber.Close
   Set rstUnitNumber = Nothing
   Set conUnitNumber = Nothing

   If NEWMODE_ And bUnitID Then
      txtUnitNo.text = ""
      txtUnitNo.SetFocus
      ShowMsgInTaskBar "Unit ID already exits. Please enter unique unit id.", , "N"
   Else
      If txtUnitNo.text <> "" And txtUnitNo.text <> lblMainUnit(11).Caption And bUnitID Then
         txtUnitNo.text = ""
         txtUnitNo.SetFocus
         ShowMsgInTaskBar "Unit ID already exits. Please enter unique unit id.", , "N"
      End If
   End If
End Sub

Private Sub txtUnitPostCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          FocusControl cboUnitType
    End If
End Sub

Private Sub txtUnitUtilityStDt_Change()
   TextBoxChangeDate txtUnitUtilityStDt
End Sub

Private Sub txtUnitUtilityStDt_GotFocus()
   SelTxtInCtrl txtUnitUtilityStDt
End Sub

Private Sub txtUnitUtilityStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtUnitUtilityStDt, KeyAscii
End Sub

Private Sub txtUnitUtilityStDt_LostFocus()
   TextBoxFormatDate txtUnitUtilityStDt
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
End Sub
