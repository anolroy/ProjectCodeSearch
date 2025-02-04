VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientNew4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   12555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21105
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientNew4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12555
   ScaleWidth      =   21105
   Begin VB.Frame fraSearch 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Automatic Demand Generate:"
      ForeColor       =   &H00FF00FF&
      Height          =   2220
      Left            =   9360
      TabIndex        =   499
      Top             =   10530
      Visible         =   0   'False
      Width           =   3715
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E5E5E5&
         Height          =   2100
         Index           =   0
         Left            =   40
         ScaleHeight     =   2040
         ScaleWidth      =   3555
         TabIndex        =   500
         Top             =   90
         Width           =   3615
         Begin VB.CommandButton cmdCloseSearch 
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
            Left            =   3330
            Style           =   1  'Graphical
            TabIndex        =   507
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchToD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2025
            MaxLength       =   80
            TabIndex        =   506
            Top             =   1125
            Width           =   1380
         End
         Begin VB.TextBox txtSearchFromD 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   80
            TabIndex        =   505
            Top             =   1125
            Width           =   1290
         End
         Begin VB.TextBox txtSearchRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   504
            Top             =   790
            Width           =   2685
         End
         Begin VB.TextBox txtSearchNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   10
            TabIndex        =   503
            Top             =   450
            Width           =   2685
         End
         Begin VB.CommandButton cmdSearchCancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   2055
            TabIndex        =   502
            Top             =   1635
            Width           =   1200
         End
         Begin VB.CommandButton cmdSearchOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   501
            Top             =   1605
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   96
            Left            =   135
            TabIndex        =   510
            Top             =   1125
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref."
            Height          =   210
            Index           =   95
            Left            =   135
            TabIndex        =   509
            Top             =   810
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   195
            Index           =   94
            Left            =   135
            TabIndex        =   508
            Top             =   495
            Width           =   210
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00C0FFFF&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   30
            Left            =   0
            Top             =   260
            Width           =   3855
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   55
            Index           =   2
            Left            =   0
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2610
      TabIndex        =   214
      Top             =   10575
      Visible         =   0   'False
      Width           =   6630
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
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   4335
         Left            =   90
         TabIndex        =   215
         Top             =   675
         Width           =   6435
         _ExtentX        =   11351
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
      Begin VB.Shape Shape2 
         Height          =   5010
         Left            =   0
         Top             =   0
         Width           =   6585
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   219
         Top             =   180
         Width           =   1230
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "2170;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   315
         TabIndex        =   218
         Top             =   390
         Width           =   1305
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2302;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   220
         Top             =   390
         Width           =   3555
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6271;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   217
         Top             =   195
         Width           =   2100
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "3704;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   2
         Left            =   5175
         TabIndex        =   216
         Top             =   180
         Width           =   915
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "1614;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   5265
         TabIndex        =   222
         Top             =   390
         Width           =   1170
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2064;450"
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
         Left            =   90
         Top             =   135
         Width           =   6345
      End
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   9075
      Left            =   90
      TabIndex        =   69
      Top             =   1125
      Width           =   22875
      _ExtentX        =   40349
      _ExtentY        =   16007
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmClientNew4.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Property"
      TabPicture(1)   =   "frmClientNew4.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblImageName"
      Tab(1).Control(1)=   "cmdImgLeftMove"
      Tab(1).Control(2)=   "imgPremises"
      Tab(1).Control(3)=   "tvwLandLord"
      Tab(1).Control(4)=   "fraOccupied"
      Tab(1).Control(5)=   "fraType"
      Tab(1).Control(6)=   "cmdImgDelete"
      Tab(1).Control(7)=   "cmdUploadImageAdd"
      Tab(1).Control(8)=   "imgList"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Agreement"
      TabPicture(2)   =   "frmClientNew4.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fraagreement"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bank Accounts"
      TabPicture(3)   =   "frmClientNew4.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraBank(1)"
      Tab(3).Control(1)=   "cmdSavePaymentDetails"
      Tab(3).Control(2)=   "cmdPaymentTypeCancel"
      Tab(3).Control(3)=   "cmdPaymentTypeUpdate"
      Tab(3).Control(4)=   "Frame6"
      Tab(3).Control(5)=   "cmdAddNewBank"
      Tab(3).Control(6)=   "cmdSaveBank"
      Tab(3).Control(7)=   "cmdDeleteBank"
      Tab(3).Control(8)=   "cmdCancelBank"
      Tab(3).Control(9)=   "cmdSetDefaultAC"
      Tab(3).Control(10)=   "cmdBACS"
      Tab(3).Control(11)=   "cmdEdit"
      Tab(3).Control(12)=   "fraBank(0)"
      Tab(3).Control(13)=   "Frame14"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Account History"
      TabPicture(4)   =   "frmClientNew4.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command1"
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(2)=   "cmdSearch"
      Tab(4).Control(3)=   "cmdClientFilter"
      Tab(4).Control(4)=   "chkShowOutstanding"
      Tab(4).Control(5)=   "cmdSupplierFilter"
      Tab(4).Control(6)=   "flxACHistory"
      Tab(4).Control(7)=   "flxACHistorySplit"
      Tab(4).Control(8)=   "Label11(22)"
      Tab(4).Control(9)=   "Label11(0)"
      Tab(4).Control(10)=   "txtFilterClient"
      Tab(4).Control(11)=   "Label1(92)"
      Tab(4).Control(12)=   "txtACBalanceByCl"
      Tab(4).Control(13)=   "Label1(91)"
      Tab(4).Control(14)=   "txtSupplierFilter"
      Tab(4).Control(15)=   "txtSearchRef1"
      Tab(4).Control(16)=   "Label1(90)"
      Tab(4).Control(17)=   "Label11(5)"
      Tab(4).Control(18)=   "Label11(10)"
      Tab(4).Control(19)=   "Label11(11)"
      Tab(4).Control(20)=   "Label11(12)"
      Tab(4).Control(21)=   "Label11(13)"
      Tab(4).Control(22)=   "Label11(21)"
      Tab(4).Control(23)=   "Label11(19)"
      Tab(4).Control(24)=   "Label11(14)"
      Tab(4).Control(25)=   "Label11(20)"
      Tab(4).Control(26)=   "Label11(15)"
      Tab(4).Control(27)=   "Label11(16)"
      Tab(4).Control(28)=   "Label11(17)"
      Tab(4).Control(29)=   "Label11(18)"
      Tab(4).Control(30)=   "Label11(1)"
      Tab(4).Control(31)=   "Label11(2)"
      Tab(4).Control(32)=   "Label11(4)"
      Tab(4).Control(33)=   "Label11(6)"
      Tab(4).Control(34)=   "Label11(7)"
      Tab(4).Control(35)=   "Label11(8)"
      Tab(4).Control(36)=   "Label11(9)"
      Tab(4).Control(37)=   "Label11(3)"
      Tab(4).Control(38)=   "lblGridCaption(1)"
      Tab(4).ControlCount=   39
      TabCaption(5)   =   "Global Settings"
      TabPicture(5)   =   "frmClientNew4.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Memo/Attachment"
      TabPicture(6)   =   "frmClientNew4.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame17(6)"
      Tab(6).Control(1)=   "Frame17(5)"
      Tab(6).Control(2)=   "Frame17(4)"
      Tab(6).Control(3)=   "Frame17(3)"
      Tab(6).Control(4)=   "Frame17(1)"
      Tab(6).Control(5)=   "Frame17(2)"
      Tab(6).Control(6)=   "Frame17(0)"
      Tab(6).Control(7)=   "Frame2"
      Tab(6).Control(8)=   "Shape1(4)"
      Tab(6).ControlCount=   9
      Begin VB.Frame Frame17 
         Caption         =   "Produce Client Statement Report Template"
         Height          =   735
         Index           =   6
         Left            =   -74910
         TabIndex        =   537
         Top             =   8280
         Width           =   12795
         Begin VB.CommandButton cmdBrowsFile 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   3
            Left            =   7380
            Style           =   1  'Graphical
            TabIndex        =   541
            Top             =   270
            Width           =   345
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   19
            Left            =   9450
            TabIndex        =   540
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   18
            Left            =   8100
            TabIndex        =   539
            Top             =   270
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   20
            Left            =   10935
            TabIndex        =   538
            Top             =   270
            Width           =   1350
         End
         Begin MSForms.TextBox txtComments2 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   542
            Top             =   270
            Width           =   7185
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12674;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Preview Client Statement Report Template"
         Height          =   735
         Index           =   5
         Left            =   -74910
         TabIndex        =   531
         Top             =   7560
         Width           =   12795
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   17
            Left            =   10935
            TabIndex        =   535
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   15
            Left            =   8100
            TabIndex        =   534
            Top             =   270
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   16
            Left            =   9450
            TabIndex        =   533
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdBrowsFile 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   1
            Left            =   7380
            Style           =   1  'Graphical
            TabIndex        =   532
            Top             =   270
            Width           =   345
         End
         Begin MSForms.TextBox txtComments2 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   536
            Top             =   270
            Width           =   7185
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12674;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Lessee Account History Report Template"
         Height          =   735
         Index           =   4
         Left            =   -74910
         TabIndex        =   525
         Top             =   6840
         Width           =   12795
         Begin VB.CommandButton cmdBrowsFile 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   0
            Left            =   7380
            Style           =   1  'Graphical
            TabIndex        =   529
            Top             =   270
            Width           =   345
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   13
            Left            =   9450
            TabIndex        =   528
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   12
            Left            =   8100
            TabIndex        =   527
            Top             =   270
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   14
            Left            =   10935
            TabIndex        =   526
            Top             =   270
            Width           =   1350
         End
         Begin MSForms.TextBox txtComments2 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   530
            Top             =   270
            Width           =   7185
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12674;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Old"
         Height          =   465
         Left            =   -56190
         TabIndex        =   514
         Top             =   8190
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   465
         Left            =   -55200
         TabIndex        =   513
         Top             =   8190
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   -66855
         Style           =   1  'Graphical
         TabIndex        =   511
         Top             =   7470
         Width           =   1080
      End
      Begin VB.CommandButton cmdClientFilter 
         Caption         =   "..."
         Height          =   315
         Left            =   -71850
         Style           =   1  'Graphical
         TabIndex        =   491
         Top             =   450
         Width           =   330
      End
      Begin VB.CheckBox chkShowOutstanding 
         Caption         =   "Show Outstanding only"
         Height          =   240
         Left            =   -67350
         TabIndex        =   490
         Top             =   495
         Width           =   2490
      End
      Begin VB.CommandButton cmdSupplierFilter 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69735
         Style           =   1  'Graphical
         TabIndex        =   489
         Top             =   450
         Width           =   330
      End
      Begin VB.Frame Frame17 
         Caption         =   "Lessee Statement Report Template"
         Height          =   735
         Index           =   3
         Left            =   -74910
         TabIndex        =   483
         Top             =   6075
         Width           =   12795
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   11
            Left            =   10935
            TabIndex        =   524
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   9
            Left            =   8100
            TabIndex        =   523
            Top             =   270
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   10
            Left            =   9450
            TabIndex        =   486
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdBrowsFile 
            Caption         =   ".."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   2
            Left            =   7380
            Style           =   1  'Graphical
            TabIndex        =   484
            Top             =   270
            Width           =   345
         End
         Begin MSForms.TextBox txtComments2 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   485
            Top             =   270
            Width           =   7185
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12674;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Bank Account Details :"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3780
         Index           =   1
         Left            =   -67800
         TabIndex        =   440
         Top             =   405
         Width           =   7905
         Begin VB.TextBox txtAvailableBankBalance 
            Alignment       =   1  'Right Justify
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
            Left            =   6705
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   522
            Top             =   2655
            Width           =   1050
         End
         Begin VB.TextBox txtRetention 
            Alignment       =   1  'Right Justify
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
            Left            =   3645
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   520
            Top             =   2655
            Width           =   1140
         End
         Begin VB.TextBox txtBankBalance 
            Alignment       =   1  'Right Justify
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   518
            Top             =   2655
            Width           =   1140
         End
         Begin VB.CommandButton cmdPaymentTypeNew 
            Caption         =   "..."
            Height          =   285
            Index           =   2
            Left            =   6300
            TabIndex        =   463
            Top             =   585
            Width           =   315
         End
         Begin VB.CommandButton cmdPaymentTypeNew 
            Caption         =   "..."
            Height          =   285
            Index           =   1
            Left            =   5940
            TabIndex        =   462
            Top             =   585
            Width           =   315
         End
         Begin VB.TextBox txtBank_AC_Name 
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   453
            Top             =   900
            Width           =   4470
         End
         Begin VB.TextBox txtBANK_SC 
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   452
            Top             =   1248
            Width           =   4470
         End
         Begin VB.TextBox txtBANK_AC_NUM 
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
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   451
            Top             =   1584
            Width           =   4470
         End
         Begin VB.TextBox txtOverDraft 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   13
            TabIndex        =   450
            Top             =   2325
            Width           =   1770
         End
         Begin VB.CommandButton cmdAddEditBankCode 
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
            Left            =   6345
            Picture         =   "frmClientNew4.frx":098E
            Style           =   1  'Graphical
            TabIndex        =   449
            Top             =   240
            Width           =   600
         End
         Begin VB.CheckBox chkOverDraft 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   448
            Top             =   2325
            Width           =   255
         End
         Begin VB.CheckBox chkConsolidated 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5730
            TabIndex        =   447
            Top             =   2340
            Width           =   255
         End
         Begin VB.TextBox txtNCCODE 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   446
            Top             =   225
            Width           =   990
         End
         Begin VB.TextBox txtNominal 
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
            Left            =   2475
            Locked          =   -1  'True
            TabIndex        =   445
            Top             =   225
            Width           =   3420
         End
         Begin VB.CommandButton cmdNC 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5895
            Style           =   1  'Graphical
            TabIndex        =   444
            Top             =   225
            Width           =   345
         End
         Begin VB.CommandButton cmdconsolidatedAccountName 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5985
            Style           =   1  'Graphical
            TabIndex        =   443
            Top             =   1935
            Width           =   345
         End
         Begin VB.TextBox txtconsolidatedAccountName 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   442
            Top             =   1935
            Width           =   4500
         End
         Begin VB.TextBox txtPaymentMethod 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   441
            Top             =   540
            Width           =   4455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Available Bank Balance:"
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
            Index           =   98
            Left            =   4950
            TabIndex        =   521
            Top             =   2700
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Retentions:"
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
            Index           =   97
            Left            =   2790
            TabIndex        =   519
            Top             =   2700
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Balance:"
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
            Index           =   93
            Left            =   135
            TabIndex        =   517
            Top             =   2700
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
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
            Index           =   57
            Left            =   120
            TabIndex        =   461
            Top             =   870
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
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
            Index           =   58
            Left            =   120
            TabIndex        =   460
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
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
            Index           =   59
            Left            =   120
            TabIndex        =   459
            Top             =   1575
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Allow Overdraft:"
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
            Index           =   60
            Left            =   120
            TabIndex        =   458
            Top             =   2325
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method:"
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
            Index           =   56
            Left            =   120
            TabIndex        =   457
            Top             =   558
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Code:"
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
            Index           =   17
            Left            =   120
            TabIndex        =   456
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Consolidated"
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
            Left            =   4635
            TabIndex        =   455
            Top             =   2385
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Consolidated AC :"
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
            Index           =   47
            Left            =   120
            TabIndex        =   454
            Top             =   1935
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdSavePaymentDetails 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -57135
         TabIndex        =   421
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdPaymentTypeCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -55830
         TabIndex        =   423
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdPaymentTypeUpdate 
         Caption         =   "U&pdate Payment Details"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -59385
         TabIndex        =   418
         Top             =   6750
         Width           =   2160
      End
      Begin VB.Frame Frame6 
         Caption         =   "Client Payment Details:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -59880
         TabIndex        =   395
         Top             =   450
         Width           =   5640
         Begin VB.CheckBox chkUsePayableTemplate 
            Caption         =   "Check1"
            Height          =   225
            Left            =   3015
            TabIndex        =   420
            Top             =   3960
            Width           =   240
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   14
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   414
            Top             =   2385
            Width           =   3015
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   11
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   411
            Top             =   1215
            Width           =   3015
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   12
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   412
            Top             =   1620
            Width           =   3015
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   13
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   413
            Top             =   2025
            Width           =   3015
         End
         Begin VB.CommandButton cmdRemittanceTemplate 
            Caption         =   "..."
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
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   416
            Top             =   3150
            Width           =   345
         End
         Begin VB.TextBox txtRemittanceTemplate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1665
            TabIndex        =   415
            Top             =   3150
            Width           =   2955
         End
         Begin VB.TextBox txtRenSummaryStatement 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1665
            TabIndex        =   417
            Top             =   3555
            Width           =   2955
         End
         Begin VB.CommandButton cmdBrowseTemplate 
            Caption         =   "..."
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
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   419
            Top             =   3555
            Width           =   345
         End
         Begin VB.CommandButton cmdPaymentType 
            Caption         =   "..."
            Height          =   285
            Left            =   4590
            TabIndex        =   408
            Top             =   450
            Width           =   315
         End
         Begin VB.TextBox txtPaymentType 
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
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   407
            Top             =   450
            Width           =   2835
         End
         Begin VB.CommandButton cmdPaymentTypeNew 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   4950
            TabIndex        =   409
            Top             =   450
            Width           =   315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Use Payable Type Statement Template:"
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
            Index           =   15
            Left            =   45
            TabIndex        =   479
            Top             =   4005
            Width           =   2700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Report Templates :"
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
            Index           =   88
            Left            =   90
            TabIndex        =   478
            Top             =   2880
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Payment Ref:"
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
            Index           =   89
            Left            =   90
            TabIndex        =   477
            Top             =   2385
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
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
            Index           =   86
            Left            =   90
            TabIndex        =   476
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code :"
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
            Index           =   87
            Left            =   90
            TabIndex        =   475
            Top             =   1605
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number :"
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
            Index           =   67
            Left            =   90
            TabIndex        =   474
            Top             =   1980
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Remittance Template:"
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
            Index           =   14
            Left            =   45
            TabIndex        =   424
            Top             =   3195
            Width           =   1515
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Statement Template:"
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
            Index           =   13
            Left            =   45
            TabIndex        =   422
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Terms:"
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
            Index           =   85
            Left            =   90
            TabIndex        =   406
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type:"
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
            Index           =   84
            Left            =   90
            TabIndex        =   405
            Top             =   450
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   83
            Left            =   2955
            TabIndex        =   396
            Top             =   870
            Width           =   735
         End
         Begin MSForms.TextBox txtPaymentTerms 
            Height          =   285
            Left            =   1710
            TabIndex        =   410
            Top             =   810
            Width           =   1005
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "1773;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6765
         Index           =   5
         Left            =   -74820
         TabIndex        =   235
         Top             =   495
         Width           =   19860
         Begin VB.TextBox txtSearchProperties2 
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
            Left            =   3150
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   236
            Top             =   360
            Width           =   1050
         End
         Begin TabDlg.SSTab TabGlobalSettingSub 
            Height          =   6405
            Left            =   4365
            TabIndex        =   239
            Top             =   180
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   11298
            _Version        =   393216
            Style           =   1
            Tabs            =   1
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   " Fees and Charges  Global Settings"
            TabPicture(0)   =   "frmClientNew4.frx":0F18
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame1(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            Begin VB.Frame Frame1 
               Height          =   5955
               Index           =   4
               Left            =   90
               TabIndex        =   240
               Top             =   360
               Width           =   15090
               Begin VB.TextBox txtNoOfDaysToSendMFB4Due 
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
                  Left            =   2295
                  Locked          =   -1  'True
                  MaxLength       =   5
                  TabIndex        =   249
                  Top             =   315
                  Width           =   1125
               End
               Begin VB.CommandButton cmdAutoSetup 
                  BackColor       =   &H80000013&
                  Caption         =   "Auto Date Fill"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   1
                  Left            =   765
                  MaskColor       =   &H00FFFFFF&
                  TabIndex        =   247
                  Top             =   5175
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSaveFEEnCharge 
                  Caption         =   "&Save"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   8100
                  TabIndex        =   245
                  Top             =   5220
                  Width           =   1215
               End
               Begin VB.CommandButton cmdEditFeeNChargePaydates 
                  Caption         =   "&Edit"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   6750
                  TabIndex        =   243
                  Top             =   5220
                  Width           =   1215
               End
               Begin VB.CommandButton cmdCancelFeenCharge 
                  Caption         =   "Canc&el"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   9405
                  TabIndex        =   241
                  Top             =   5220
                  Width           =   1215
               End
               Begin TabDlg.SSTab tabDates 
                  Height          =   3435
                  Index           =   1
                  Left            =   180
                  TabIndex        =   252
                  Top             =   1125
                  Width           =   10545
                  _ExtentX        =   18600
                  _ExtentY        =   6059
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   4
                  TabsPerRow      =   4
                  TabHeight       =   520
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  TabCaption(0)   =   "Monthly Payment Dates"
                  TabPicture(0)   =   "frmClientNew4.frx":0F34
                  Tab(0).ControlEnabled=   -1  'True
                  Tab(0).Control(0)=   "fraPaymentDate(11)"
                  Tab(0).Control(0).Enabled=   0   'False
                  Tab(0).Control(1)=   "fraPaymentDate(10)"
                  Tab(0).Control(1).Enabled=   0   'False
                  Tab(0).Control(2)=   "fraPaymentDate(9)"
                  Tab(0).Control(2).Enabled=   0   'False
                  Tab(0).ControlCount=   3
                  TabCaption(1)   =   "Quarterly Payment Dates"
                  TabPicture(1)   =   "frmClientNew4.frx":0F50
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "fraPaymentDate(12)"
                  Tab(1).ControlCount=   1
                  TabCaption(2)   =   "Half Yearly payments"
                  TabPicture(2)   =   "frmClientNew4.frx":0F6C
                  Tab(2).ControlEnabled=   0   'False
                  Tab(2).Control(0)=   "fraPaymentDate(13)"
                  Tab(2).ControlCount=   1
                  TabCaption(3)   =   "Yearly payments"
                  TabPicture(3)   =   "frmClientNew4.frx":0F88
                  Tab(3).ControlEnabled=   0   'False
                  Tab(3).Control(0)=   "fraPaymentDate(14)"
                  Tab(3).ControlCount=   1
                  Begin VB.Frame fraPaymentDate 
                     Caption         =   "Yearly Payment Date"
                     Enabled         =   0   'False
                     Height          =   3000
                     Index           =   6
                     Left            =   -74940
                     TabIndex        =   347
                     Top             =   405
                     Width           =   11865
                     Begin VB.ComboBox cboYDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   1
                        Left            =   3135
                        TabIndex        =   349
                        Top             =   540
                        Width           =   615
                     End
                     Begin VB.ComboBox cboYMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   1
                        Left            =   3855
                        TabIndex        =   348
                        Top             =   540
                        Width           =   1335
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Once:"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   4
                        Left            =   2505
                        TabIndex        =   350
                        Top             =   540
                        Width           =   465
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Caption         =   "Half Yearly Payment Dates"
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Myriad Condensed Web"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2970
                     Index           =   7
                     Left            =   -74895
                     TabIndex        =   340
                     Top             =   420
                     Width           =   11775
                     Begin VB.ComboBox cboHMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   2
                        Left            =   4800
                        TabIndex        =   344
                        Top             =   975
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboHDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   2
                        Left            =   4080
                        TabIndex        =   343
                        Top             =   975
                        Width           =   615
                     End
                     Begin VB.ComboBox cboHMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   3
                        Left            =   4800
                        TabIndex        =   342
                        Top             =   495
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboHDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   3
                        Left            =   4080
                        TabIndex        =   341
                        Top             =   495
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "First"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   27
                        Left            =   3375
                        TabIndex        =   346
                        Top             =   495
                        Width           =   330
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Second"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   28
                        Left            =   3390
                        TabIndex        =   345
                        Top             =   975
                        Width           =   585
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Caption         =   "Quarterly Payment Dates"
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Myriad Condensed Web"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2925
                     Index           =   8
                     Left            =   -74910
                     TabIndex        =   327
                     Top             =   510
                     Width           =   11790
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   4
                        Left            =   5175
                        TabIndex        =   335
                        Top             =   1845
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   4
                        Left            =   4455
                        TabIndex        =   334
                        Top             =   1845
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   5
                        Left            =   5175
                        TabIndex        =   333
                        Top             =   1365
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   5
                        Left            =   4455
                        TabIndex        =   332
                        Top             =   1365
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   6
                        Left            =   5175
                        TabIndex        =   331
                        Top             =   885
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   6
                        Left            =   4455
                        TabIndex        =   330
                        Top             =   885
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   7
                        Left            =   5175
                        TabIndex        =   329
                        Top             =   405
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   7
                        Left            =   4455
                        TabIndex        =   328
                        Top             =   405
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Fourth"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   29
                        Left            =   3240
                        TabIndex        =   339
                        Top             =   1845
                        Width           =   525
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Third"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   30
                        Left            =   3240
                        TabIndex        =   338
                        Top             =   1365
                        Width           =   405
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Second"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   31
                        Left            =   3240
                        TabIndex        =   337
                        Top             =   885
                        Width           =   585
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "First"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   32
                        Left            =   3240
                        TabIndex        =   336
                        Top             =   405
                        Width           =   330
                     End
                  End
                  Begin VB.Frame Frame12 
                     Caption         =   "Quarterly Payment Dates"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2295
                     Index           =   1
                     Left            =   -72180
                     TabIndex        =   314
                     Top             =   960
                     Width           =   3015
                     Begin VB.ComboBox cboM4 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   322
                        Top             =   1800
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD4 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   321
                        Top             =   1800
                        Width           =   615
                     End
                     Begin VB.ComboBox cboM3 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   320
                        Top             =   1320
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD3 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   319
                        Top             =   1320
                        Width           =   615
                     End
                     Begin VB.ComboBox cboM2 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   318
                        Top             =   840
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD2 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   317
                        Top             =   840
                        Width           =   615
                     End
                     Begin VB.ComboBox cboM1 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   316
                        Top             =   360
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD1 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   315
                        Top             =   360
                        Width           =   615
                     End
                     Begin VB.Label Label47 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Fourth"
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
                        Index           =   1
                        Left            =   120
                        TabIndex        =   326
                        Top             =   1800
                        Width           =   615
                     End
                     Begin VB.Label Label46 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Third"
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
                        Index           =   1
                        Left            =   120
                        TabIndex        =   325
                        Top             =   1320
                        Width           =   615
                     End
                     Begin VB.Label Label45 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Second"
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
                        Index           =   1
                        Left            =   120
                        TabIndex        =   324
                        Top             =   840
                        Width           =   615
                     End
                     Begin VB.Label Label44 
                        Alignment       =   1  'Right Justify
                        Caption         =   "First"
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
                        Index           =   1
                        Left            =   360
                        TabIndex        =   323
                        Top             =   360
                        Width           =   375
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2970
                     Index           =   9
                     Left            =   6840
                     TabIndex        =   305
                     Top             =   420
                     Width           =   3555
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   42
                        Left            =   2385
                        TabIndex        =   309
                        Top             =   2070
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   41
                        Left            =   2385
                        TabIndex        =   308
                        Top             =   1590
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   40
                        Left            =   2385
                        TabIndex        =   307
                        Top             =   1110
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   39
                        Left            =   2385
                        TabIndex        =   306
                        Top             =   630
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Caption         =   "December:"
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
                        Index           =   46
                        Left            =   810
                        TabIndex        =   437
                        Top             =   2115
                        Width           =   1185
                     End
                     Begin VB.Label Label1 
                        Caption         =   "November:"
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
                        Index           =   45
                        Left            =   810
                        TabIndex        =   436
                        Top             =   1620
                        Width           =   1050
                     End
                     Begin VB.Label Label1 
                        Caption         =   "October:"
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
                        Index           =   44
                        Left            =   810
                        TabIndex        =   435
                        Top             =   1125
                        Width           =   1140
                     End
                     Begin VB.Label Label1 
                        Caption         =   "September:"
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
                        Index           =   43
                        Left            =   810
                        TabIndex        =   434
                        Top             =   675
                        Width           =   1005
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "12th"
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
                        Index           =   33
                        Left            =   1950
                        TabIndex        =   313
                        Top             =   2130
                        Width           =   330
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "11th"
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
                        Index           =   34
                        Left            =   1950
                        TabIndex        =   312
                        Top             =   1650
                        Width           =   330
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "10th"
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
                        Index           =   53
                        Left            =   1950
                        TabIndex        =   311
                        Top             =   1170
                        Width           =   330
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "9th"
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
                        Index           =   54
                        Left            =   1995
                        TabIndex        =   310
                        Top             =   690
                        Width           =   240
                     End
                  End
                  Begin VB.Frame Frame9 
                     Caption         =   "Half Yearly Payment Dates"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1575
                     Index           =   1
                     Left            =   -71580
                     TabIndex        =   298
                     Top             =   1260
                     Width           =   3135
                     Begin VB.ComboBox cboM6 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   302
                        Top             =   840
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD6 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   301
                        Top             =   840
                        Width           =   615
                     End
                     Begin VB.ComboBox cboM5 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   300
                        Top             =   360
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboD5 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   840
                        TabIndex        =   299
                        Top             =   360
                        Width           =   615
                     End
                     Begin VB.Label Label22 
                        Alignment       =   1  'Right Justify
                        Caption         =   "First"
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
                        Index           =   1
                        Left            =   240
                        TabIndex        =   304
                        Top             =   360
                        Width           =   495
                     End
                     Begin VB.Label Label21 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Second"
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
                        Index           =   1
                        Left            =   120
                        TabIndex        =   303
                        Top             =   840
                        Width           =   615
                     End
                  End
                  Begin VB.Frame Frame8 
                     Caption         =   "Yearly Payment Date"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   975
                     Index           =   1
                     Left            =   -71317
                     TabIndex        =   295
                     Top             =   1440
                     Width           =   2535
                     Begin VB.ComboBox cboD7 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   240
                        TabIndex        =   297
                        Top             =   360
                        Width           =   615
                     End
                     Begin VB.ComboBox cboM7 
                        Enabled         =   0   'False
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   1
                        Left            =   960
                        TabIndex        =   296
                        Top             =   360
                        Width           =   1335
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2970
                     Index           =   10
                     Left            =   3435
                     TabIndex        =   286
                     Top             =   420
                     Width           =   3285
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   35
                        Left            =   2070
                        TabIndex        =   290
                        Top             =   630
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   36
                        Left            =   2070
                        TabIndex        =   289
                        Top             =   1110
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   37
                        Left            =   2070
                        TabIndex        =   288
                        Top             =   1590
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   38
                        Left            =   2070
                        TabIndex        =   287
                        Top             =   2070
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Caption         =   "August:"
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
                        Index           =   42
                        Left            =   675
                        TabIndex        =   433
                        Top             =   2115
                        Width           =   825
                     End
                     Begin VB.Label Label1 
                        Caption         =   "July:"
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
                        Index           =   41
                        Left            =   675
                        TabIndex        =   432
                        Top             =   1620
                        Width           =   375
                     End
                     Begin VB.Label Label1 
                        Caption         =   "June:"
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
                        Index           =   40
                        Left            =   675
                        TabIndex        =   431
                        Top             =   1125
                        Width           =   375
                     End
                     Begin VB.Label Label1 
                        Caption         =   "May:"
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
                        Index           =   39
                        Left            =   675
                        TabIndex        =   430
                        Top             =   675
                        Width           =   375
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "7th"
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
                        Index           =   55
                        Left            =   1725
                        TabIndex        =   294
                        Top             =   1650
                        Width           =   240
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "5th"
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
                        Index           =   61
                        Left            =   1725
                        TabIndex        =   293
                        Top             =   690
                        Width           =   240
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "6th"
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
                        Index           =   62
                        Left            =   1725
                        TabIndex        =   292
                        Top             =   1170
                        Width           =   240
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "8th"
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
                        Index           =   63
                        Left            =   1725
                        TabIndex        =   291
                        Top             =   2130
                        Width           =   240
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   2970
                     Index           =   11
                     Left            =   90
                     TabIndex        =   277
                     Top             =   420
                     Width           =   3240
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   34
                        Left            =   2430
                        TabIndex        =   281
                        Top             =   2070
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   33
                        Left            =   2430
                        TabIndex        =   280
                        Top             =   1590
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   32
                        Left            =   2430
                        TabIndex        =   279
                        Top             =   1110
                        Width           =   615
                     End
                     Begin VB.ComboBox cboDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Index           =   31
                        ItemData        =   "frmClientNew4.frx":0FA4
                        Left            =   2430
                        List            =   "frmClientNew4.frx":0FA6
                        TabIndex        =   278
                        Top             =   630
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Caption         =   "April:"
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
                        Index           =   38
                        Left            =   630
                        TabIndex        =   429
                        Top             =   2115
                        Width           =   375
                     End
                     Begin VB.Label Label1 
                        Caption         =   "March:"
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
                        Index           =   37
                        Left            =   630
                        TabIndex        =   428
                        Top             =   1620
                        Width           =   1185
                     End
                     Begin VB.Label Label1 
                        Caption         =   "February:"
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
                        Index           =   36
                        Left            =   630
                        TabIndex        =   427
                        Top             =   1125
                        Width           =   1140
                     End
                     Begin VB.Label Label1 
                        Caption         =   "January:"
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
                        Index           =   35
                        Left            =   630
                        TabIndex        =   426
                        Top             =   675
                        Width           =   870
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "4th"
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
                        Index           =   64
                        Left            =   1950
                        TabIndex        =   285
                        Top             =   2130
                        Width           =   240
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "3rd"
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
                        Index           =   65
                        Left            =   1950
                        TabIndex        =   284
                        Top             =   1650
                        Width           =   240
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "2nd"
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
                        Index           =   66
                        Left            =   1950
                        TabIndex        =   283
                        Top             =   1170
                        Width           =   270
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        Caption         =   "1st"
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
                        Index           =   68
                        Left            =   1815
                        TabIndex        =   282
                        Top             =   690
                        Width           =   375
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Myriad Condensed Web"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   3015
                     Index           =   12
                     Left            =   -74910
                     TabIndex        =   264
                     Top             =   315
                     Width           =   10350
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   14
                        Left            =   5175
                        TabIndex        =   272
                        Top             =   1845
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   14
                        Left            =   4455
                        TabIndex        =   271
                        Top             =   1845
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   13
                        Left            =   5175
                        TabIndex        =   270
                        Top             =   1365
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   13
                        Left            =   4455
                        TabIndex        =   269
                        Top             =   1365
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   12
                        Left            =   5175
                        TabIndex        =   268
                        Top             =   885
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   12
                        Left            =   4455
                        TabIndex        =   267
                        Top             =   885
                        Width           =   615
                     End
                     Begin VB.ComboBox cboQMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   11
                        Left            =   5175
                        TabIndex        =   266
                        Top             =   405
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboQDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   11
                        Left            =   4455
                        TabIndex        =   265
                        Top             =   405
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Fourth"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   21
                        Left            =   3240
                        TabIndex        =   276
                        Top             =   1845
                        Width           =   525
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Third"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   69
                        Left            =   3240
                        TabIndex        =   275
                        Top             =   1365
                        Width           =   405
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Second"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   70
                        Left            =   3240
                        TabIndex        =   274
                        Top             =   885
                        Width           =   585
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "First"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   71
                        Left            =   3240
                        TabIndex        =   273
                        Top             =   405
                        Width           =   330
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Myriad Condensed Web"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   3105
                     Index           =   13
                     Left            =   -74910
                     TabIndex        =   257
                     Top             =   270
                     Width           =   10380
                     Begin VB.ComboBox cboHMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   6
                        Left            =   4800
                        TabIndex        =   261
                        Top             =   975
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboHDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   6
                        Left            =   4080
                        TabIndex        =   260
                        Top             =   975
                        Width           =   615
                     End
                     Begin VB.ComboBox cboHMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   5
                        Left            =   4800
                        TabIndex        =   259
                        Top             =   495
                        Width           =   1335
                     End
                     Begin VB.ComboBox cboHDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   5
                        Left            =   4080
                        TabIndex        =   258
                        Top             =   495
                        Width           =   615
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "First"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   72
                        Left            =   3375
                        TabIndex        =   263
                        Top             =   495
                        Width           =   330
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        AutoSize        =   -1  'True
                        Caption         =   "Second"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   73
                        Left            =   3390
                        TabIndex        =   262
                        Top             =   975
                        Width           =   585
                     End
                  End
                  Begin VB.Frame fraPaymentDate 
                     Enabled         =   0   'False
                     Height          =   3045
                     Index           =   14
                     Left            =   -74910
                     TabIndex        =   253
                     Top             =   315
                     Width           =   10380
                     Begin VB.ComboBox cboYDay 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   2
                        Left            =   3135
                        TabIndex        =   255
                        Top             =   540
                        Width           =   615
                     End
                     Begin VB.ComboBox cboYMth 
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   330
                        Index           =   2
                        Left            =   3855
                        TabIndex        =   254
                        Top             =   540
                        Width           =   1335
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Once:"
                        BeginProperty Font 
                           Name            =   "Myriad Web"
                           Size            =   9
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   210
                        Index           =   74
                        Left            =   2505
                        TabIndex        =   256
                        Top             =   540
                        Width           =   465
                     End
                  End
               End
               Begin MSForms.Label Label7 
                  Height          =   300
                  Index           =   12
                  Left            =   270
                  TabIndex        =   351
                  Top             =   360
                  Width           =   8535
                  VariousPropertyBits=   8388627
                  Caption         =   "Generate Fees and Charges                                            days before due date"
                  Size            =   "15055;529"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPropertySelection2 
            Height          =   5820
            Left            =   90
            TabIndex        =   425
            Top             =   720
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   10266
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   12648447
            ForeColorSel    =   4210752
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Property Search:"
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
            Left            =   1935
            TabIndex        =   238
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label19 
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
            Index           =   0
            Left            =   270
            TabIndex        =   237
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6225
         Left            =   45
         TabIndex        =   172
         Top             =   1170
         Width           =   18105
         Begin VB.CheckBox chkStatementAddress 
            Caption         =   "Statement Address"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3555
            TabIndex        =   473
            Top             =   315
            Width           =   2085
         End
         Begin VB.CheckBox chkClientAddress 
            Caption         =   "Client Address"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1800
            TabIndex        =   472
            Top             =   315
            Width           =   1680
         End
         Begin VB.CheckBox chkConsolidatedStatement 
            Caption         =   "Consolidated Statement"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            TabIndex        =   16
            Top             =   315
            Width           =   4110
         End
         Begin VB.Frame Frame1 
            Caption         =   "Client Address:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5490
            Index           =   0
            Left            =   90
            TabIndex        =   181
            Top             =   675
            Width           =   5970
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   4
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   21
               Top             =   1575
               Width           =   4680
            End
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   18
               Top             =   600
               Width           =   4680
            End
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   5
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   1920
               Width           =   1815
            End
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   19
               Top             =   930
               Width           =   4680
            End
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   0
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   17
               Top             =   270
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   4
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   27
               Top             =   3990
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   26
               Top             =   3630
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   3270
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   2910
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   0
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   2550
               Width           =   4680
            End
            Begin VB.TextBox txtClientAddressLine1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   20
               Top             =   1245
               Width           =   4680
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   5
               Left            =   1185
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   28
               Top             =   4365
               Width           =   1455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
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
               Index           =   8
               Left            =   240
               TabIndex        =   189
               Top             =   285
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
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
               Index           =   9
               Left            =   240
               TabIndex        =   188
               Top             =   1920
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Tel:"
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
               Index           =   10
               Left            =   240
               TabIndex        =   187
               Top             =   2925
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Email:"
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
               Index           =   13
               Left            =   240
               TabIndex        =   186
               Top             =   3630
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile:"
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
               Index           =   12
               Left            =   240
               TabIndex        =   185
               Top             =   3270
               Width           =   525
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Email:"
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
               Index           =   14
               Left            =   240
               TabIndex        =   184
               Top             =   3990
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Tel:"
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
               Index           =   11
               Left            =   240
               TabIndex        =   183
               Top             =   2610
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Group Code:"
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
               Index           =   26
               Left            =   240
               TabIndex        =   182
               Top             =   4365
               Width           =   870
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Client Statement Address:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5490
            Index           =   1
            Left            =   6120
            TabIndex        =   178
            Top             =   675
            Width           =   5925
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   10
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   3960
               Width           =   4545
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   9
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   3600
               Width           =   4545
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   8
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   3240
               Width           =   4545
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   7
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   36
               Top             =   2880
               Width           =   4545
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   6
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   35
               Top             =   2520
               Width           =   4545
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   19
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   33
               Top             =   1575
               Width           =   4500
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   15
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   29
               Top             =   240
               Width           =   4500
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   17
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   31
               Top             =   900
               Width           =   4500
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   20
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   1920
               Width           =   1275
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   16
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   30
               Top             =   570
               Width           =   4500
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   18
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   70
               TabIndex        =   32
               Top             =   1230
               Width           =   4500
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Tel:"
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
               Index           =   52
               Left            =   360
               TabIndex        =   471
               Top             =   2565
               Width           =   705
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Email:"
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
               Index           =   51
               Left            =   360
               TabIndex        =   470
               Top             =   3990
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile:"
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
               Index           =   50
               Left            =   360
               TabIndex        =   469
               Top             =   3270
               Width           =   525
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Email:"
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
               Index           =   49
               Left            =   360
               TabIndex        =   468
               Top             =   3630
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Tel:"
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
               Index           =   48
               Left            =   360
               TabIndex        =   467
               Top             =   2925
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
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
               Index           =   16
               Left            =   360
               TabIndex        =   180
               Top             =   1920
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
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
               Index           =   15
               Left            =   360
               TabIndex        =   179
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Client Registered Office:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5430
            Index           =   2
            Left            =   12105
            TabIndex        =   174
            Top             =   675
            Width           =   5925
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   26
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   2025
               Width           =   4230
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   23
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   960
               Width           =   4230
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   27
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2385
               Width           =   1185
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   24
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   1290
               Width           =   4230
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   22
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   600
               Width           =   4230
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   21
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   240
               Width           =   4230
            End
            Begin VB.TextBox txtClientHomeTel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   25
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   1650
               Width           =   4230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
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
               Index           =   23
               Left            =   180
               TabIndex        =   177
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
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
               Index           =   24
               Left            =   180
               TabIndex        =   176
               Top             =   2460
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Company No"
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
               Index           =   25
               Left            =   180
               TabIndex        =   175
               Top             =   240
               Width           =   885
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Send Statement to:"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Index           =   71
            Left            =   225
            TabIndex        =   173
            Top             =   315
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   45
         TabIndex        =   171
         Top             =   360
         Width           =   18105
         Begin VB.CommandButton cmdSaveClient 
            Caption         =   "&Save"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   5085
            TabIndex        =   12
            Top             =   225
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteClient 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   9225
            TabIndex        =   14
            Top             =   225
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditClient 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   2790
            TabIndex        =   11
            Top             =   225
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelChange 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   7200
            TabIndex        =   13
            Top             =   225
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddNewClient 
            Caption         =   "&New"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   495
            TabIndex        =   10
            Top             =   225
            Width           =   1215
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   11385
            TabIndex        =   15
            Top             =   225
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdAddNewBank 
         Caption         =   "&Add Bank Account"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74820
         TabIndex        =   163
         Top             =   6750
         Width           =   1755
      End
      Begin VB.CommandButton cmdSaveBank 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -63930
         TabIndex        =   162
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteBank 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72975
         TabIndex        =   161
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelBank 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -62655
         TabIndex        =   160
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetDefaultAC 
         Caption         =   "Set &Default A/C"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65415
         TabIndex        =   159
         Top             =   6750
         Width           =   1455
      End
      Begin VB.CommandButton cmdBACS 
         Caption         =   "&BACS"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66720
         TabIndex        =   158
         Top             =   6750
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Update Bank Details"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71490
         TabIndex        =   157
         Top             =   6750
         Width           =   2160
      End
      Begin VB.PictureBox Picture1 
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
         Height          =   405
         Index           =   1
         Left            =   8820
         ScaleHeight     =   405
         ScaleWidth      =   4545
         TabIndex        =   155
         Top             =   45
         Visible         =   0   'False
         Width           =   4545
         Begin VB.Label Label3 
            Alignment       =   2  'Center
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
            Left            =   450
            TabIndex        =   156
            Top             =   105
            Width           =   3960
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Comment 1:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   1
         Left            =   -74910
         TabIndex        =   150
         Top             =   4725
         Width           =   12795
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   3
            Left            =   8100
            TabIndex        =   153
            Top             =   240
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   9450
            TabIndex        =   152
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   10935
            TabIndex        =   151
            Top             =   225
            Width           =   1350
         End
         Begin MSForms.TextBox txtComments1 
            Height          =   315
            Left            =   120
            TabIndex        =   154
            Top             =   225
            Width           =   7140
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12594;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Comment 2:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   2
         Left            =   -74910
         TabIndex        =   145
         Top             =   5370
         Width           =   12795
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Index           =   8
            Left            =   10935
            TabIndex        =   148
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Index           =   7
            Left            =   9450
            TabIndex        =   147
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Edit"
            Height          =   315
            Index           =   6
            Left            =   8100
            TabIndex        =   146
            Top             =   240
            Width           =   1260
         End
         Begin MSForms.TextBox txtComments2 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   7185
            VariousPropertyBits=   746604575
            MaxLength       =   250
            Size            =   "12674;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attachment Files:"
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
         Height          =   720
         Index           =   0
         Left            =   -74910
         TabIndex        =   140
         Top             =   4005
         Width           =   12795
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Open File"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   9450
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   8100
            Style           =   1  'Graphical
            TabIndex        =   142
            Top             =   240
            Width           =   1260
         End
         Begin VB.CommandButton cmdClientAddAtch 
            Caption         =   "&Delete File"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   10935
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   144
            Top             =   270
            Width           =   7185
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "12674;503"
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.Frame Fraagreement 
         Caption         =   "Agreement Details"
         Height          =   7755
         Left            =   -74955
         TabIndex        =   138
         Top             =   360
         Width           =   22785
         Begin VB.TextBox txtPropertySearchSel1 
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
            Left            =   3600
            MaxLength       =   12
            TabIndex        =   212
            Top             =   405
            Width           =   1050
         End
         Begin VB.Frame fraPaymentDate 
            Caption         =   "Agreement Terms"
            Height          =   2490
            Index           =   15
            Left            =   4770
            TabIndex        =   190
            Top             =   675
            Width           =   17970
            Begin VB.TextBox txtComparenextDueDate1 
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
               Left            =   3645
               MaxLength       =   12
               TabIndex        =   480
               Top             =   225
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.TextBox txtAgreementEndDate 
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
               Left            =   5805
               MaxLength       =   12
               TabIndex        =   193
               Top             =   765
               Width           =   1050
            End
            Begin VB.TextBox txtAgreementStartDate 
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
               Left            =   2250
               MaxLength       =   12
               TabIndex        =   192
               Top             =   765
               Width           =   1050
            End
            Begin VB.TextBox txtREVIEW_DATE 
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
               Left            =   9765
               MaxLength       =   12
               TabIndex        =   194
               Top             =   765
               Width           =   1050
            End
            Begin VB.CommandButton cmdAgrTopEdit 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3465
               TabIndex        =   196
               Top             =   1320
               Width           =   1815
            End
            Begin VB.CommandButton cmdAgrTopSave 
               Caption         =   "Save"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   5325
               TabIndex        =   198
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CommandButton cmdCanelAgree 
               Caption         =   "Cancel"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   6510
               TabIndex        =   199
               Top             =   1320
               Width           =   1095
            End
            Begin MSForms.Label Label7 
               Height          =   195
               Index           =   16
               Left            =   4005
               TabIndex        =   200
               Top             =   810
               Width           =   1515
               VariousPropertyBits=   276824083
               Caption         =   "Agreement End Date:"
               Size            =   "2672;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label7 
               Height          =   195
               Index           =   15
               Left            =   450
               TabIndex        =   197
               Top             =   810
               Width           =   1590
               VariousPropertyBits=   276824083
               Caption         =   "Agreement Start Date:"
               Size            =   "2805;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label Label7 
               Height          =   195
               Index           =   11
               Left            =   7875
               TabIndex        =   195
               Top             =   810
               Width           =   1770
               VariousPropertyBits=   276824083
               Caption         =   "Agreement Review Date:"
               Size            =   "3122;344"
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin TabDlg.SSTab tabAgreement 
            Height          =   4380
            Left            =   45
            TabIndex        =   139
            Top             =   3285
            Width           =   22680
            _ExtentX        =   40005
            _ExtentY        =   7726
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   15
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "   Management Fees               "
            TabPicture(0)   =   "frmClientNew4.frx":0FA8
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Shape3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fraPaymentDate(17)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "     Rent Payable                "
            TabPicture(1)   =   "frmClientNew4.frx":0FC4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraPaymentDate(16)"
            Tab(1).ControlCount=   1
            Begin VB.Frame fraPaymentDate 
               Caption         =   " Management Fee"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3885
               Index           =   17
               Left            =   45
               TabIndex        =   209
               Top             =   450
               Width           =   22605
               Begin VB.CommandButton cmdPrintAgreement 
                  Caption         =   "Print Management Fee Agreement"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   355
                  Left            =   7605
                  Style           =   1  'Graphical
                  TabIndex        =   516
                  Top             =   3405
                  Width           =   3360
               End
               Begin VB.CommandButton cmdFix 
                  Caption         =   "Fix LCD"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   355
                  Left            =   6255
                  Style           =   1  'Graphical
                  TabIndex        =   482
                  Top             =   3405
                  Width           =   1245
               End
               Begin VB.CommandButton cmdAdvanceProgr 
                  Caption         =   "Advance"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   355
                  Left            =   19710
                  Style           =   1  'Graphical
                  TabIndex        =   481
                  Top             =   3405
                  Visible         =   0   'False
                  Width           =   1380
               End
               Begin VB.CommandButton cmdClose2 
                  Caption         =   "Close"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   18135
                  TabIndex        =   392
                  Top             =   3420
                  Width           =   1215
               End
               Begin VB.CommandButton cmdDeleteMgtFee 
                  Caption         =   "&Delete"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   15570
                  TabIndex        =   390
                  Top             =   3420
                  Width           =   1215
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   12
                  Left            =   11925
                  Style           =   1  'Graphical
                  TabIndex        =   378
                  Top             =   540
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   11
                  Left            =   8955
                  Style           =   1  'Graphical
                  TabIndex        =   375
                  Top             =   540
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   10
                  Left            =   7605
                  Style           =   1  'Graphical
                  TabIndex        =   373
                  Top             =   540
                  Width           =   345
               End
               Begin VB.TextBox txtSTART_DATE 
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
                  Left            =   9315
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   376
                  Top             =   540
                  Width           =   915
               End
               Begin VB.TextBox txtManagingAgentAC 
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
                  Left            =   4995
                  Locked          =   -1  'True
                  TabIndex        =   370
                  Top             =   540
                  Width           =   1275
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   9
                  Left            =   2835
                  Style           =   1  'Graphical
                  TabIndex        =   367
                  Top             =   540
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   8
                  Left            =   6030
                  Style           =   1  'Graphical
                  TabIndex        =   394
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   7
                  Left            =   4635
                  Style           =   1  'Graphical
                  TabIndex        =   369
                  Top             =   540
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   6
                  Left            =   6300
                  Style           =   1  'Graphical
                  TabIndex        =   371
                  Top             =   540
                  Width           =   345
               End
               Begin VB.TextBox txtChargeType 
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
                  Left            =   90
                  Locked          =   -1  'True
                  TabIndex        =   366
                  Top             =   540
                  Width           =   2715
               End
               Begin VB.TextBox txtDemandTypemngtFee 
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
                  Left            =   4455
                  Locked          =   -1  'True
                  TabIndex        =   393
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   1500
               End
               Begin VB.TextBox txtFundMngtFee 
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
                  Left            =   3195
                  Locked          =   -1  'True
                  TabIndex        =   368
                  Top             =   540
                  Width           =   1410
               End
               Begin VB.CommandButton cmdAgmntEdit 
                  Caption         =   "&Edit"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   12885
                  TabIndex        =   388
                  Top             =   3405
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAgmntSave 
                  Caption         =   "&Save"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   14280
                  TabIndex        =   389
                  Top             =   3405
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAgmntAddNew 
                  Caption         =   "Add New"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   11565
                  TabIndex        =   387
                  Top             =   3405
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAgmntCancel 
                  Caption         =   "Cancel"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   16875
                  TabIndex        =   391
                  Top             =   3420
                  Width           =   1215
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxManagementFee 
                  Height          =   2220
                  Left            =   90
                  TabIndex        =   210
                  Top             =   930
                  Width           =   22380
                  _ExtentX        =   39476
                  _ExtentY        =   3916
                  _Version        =   393216
                  ForeColor       =   0
                  Cols            =   6
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   0
                  BackColorSel    =   12648447
                  ForeColorSel    =   4210752
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
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   6
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
               Begin MSForms.TextBox txtLastChargeDate 
                  Height          =   285
                  Left            =   17460
                  TabIndex        =   383
                  Top             =   540
                  Width           =   1035
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1826;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   1
                  Left            =   20835
                  TabIndex        =   439
                  Top             =   270
                  Width           =   660
                  VariousPropertyBits=   276824083
                  Caption         =   "End Date"
                  Size            =   "1164;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   26
                  Left            =   13500
                  TabIndex        =   438
                  Top             =   270
                  Width           =   1440
                  VariousPropertyBits=   8388627
                  Caption         =   "Amount/Percentage"
                  Size            =   "2540;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtAmount 
                  Height          =   285
                  Left            =   13500
                  TabIndex        =   380
                  Top             =   540
                  Width           =   1395
                  VariousPropertyBits=   679495711
                  MaxLength       =   12
                  BorderStyle     =   1
                  Size            =   "2461;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox txtCapAmount 
                  Height          =   285
                  Left            =   19575
                  TabIndex        =   385
                  Top             =   540
                  Width           =   960
                  VariousPropertyBits=   679495707
                  MaxLength       =   12
                  BorderStyle     =   1
                  Size            =   "1693;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox txtChargeBasis 
                  Height          =   285
                  Left            =   7965
                  TabIndex        =   374
                  Top             =   540
                  Width           =   990
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1746;503"
                  SpecialEffect   =   0
                  FontName        =   "Calibri"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtChargingMethod 
                  Height          =   285
                  Left            =   6705
                  TabIndex        =   372
                  Top             =   540
                  Width           =   900
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1587;503"
                  SpecialEffect   =   0
                  FontName        =   "Calibri"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtFrequecymngtFee 
                  Height          =   285
                  Left            =   10305
                  TabIndex        =   377
                  Top             =   540
                  Width           =   1530
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "2699;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtStopDatemngtFee 
                  Height          =   285
                  Left            =   18540
                  TabIndex        =   384
                  Top             =   540
                  Width           =   990
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1746;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox txtPeriod 
                  Height          =   285
                  Left            =   16515
                  TabIndex        =   382
                  Top             =   540
                  Width           =   900
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1587;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.TextBox txtTotalAmountPerYear 
                  Height          =   285
                  Left            =   15030
                  TabIndex        =   381
                  Top             =   540
                  Width           =   1395
                  VariousPropertyBits=   679495711
                  MaxLength       =   12
                  BorderStyle     =   1
                  Size            =   "2461;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   31
                  Left            =   19710
                  TabIndex        =   365
                  Top             =   270
                  Width           =   870
                  VariousPropertyBits=   276824083
                  Caption         =   "Cap Amount"
                  Size            =   "1535;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   30
                  Left            =   18765
                  TabIndex        =   364
                  Top             =   270
                  Width           =   720
                  VariousPropertyBits=   276824083
                  Caption         =   "Stop Date"
                  Size            =   "1270;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   29
                  Left            =   17415
                  TabIndex        =   363
                  Top             =   270
                  Width           =   1245
                  VariousPropertyBits=   276824083
                  Caption         =   "Last Charge  Date"
                  Size            =   "2196;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   28
                  Left            =   16515
                  TabIndex        =   362
                  Top             =   270
                  Width           =   840
                  VariousPropertyBits=   276824083
                  Caption         =   "Each Period"
                  Size            =   "1482;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   21
                  Left            =   6705
                  TabIndex        =   361
                  Top             =   270
                  Width           =   1095
                  VariousPropertyBits=   276824083
                  Caption         =   "Charge method"
                  Size            =   "1931;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   17
                  Left            =   135
                  TabIndex        =   360
                  Top             =   270
                  Width           =   900
                  VariousPropertyBits=   276824083
                  Caption         =   "Charge Type"
                  Size            =   "1588;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   18
                  Left            =   3375
                  TabIndex        =   359
                  Top             =   3645
                  Visible         =   0   'False
                  Width           =   975
                  VariousPropertyBits=   276824083
                  Caption         =   "Demand Type"
                  Size            =   "1720;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   22
                  Left            =   8025
                  TabIndex        =   358
                  Top             =   270
                  Width           =   915
                  VariousPropertyBits=   276824083
                  Caption         =   "Charge Basis"
                  Size            =   "1614;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   23
                  Left            =   9420
                  TabIndex        =   357
                  Top             =   270
                  Width           =   735
                  VariousPropertyBits=   276824083
                  Caption         =   "Start Date"
                  Size            =   "1296;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   24
                  Left            =   10395
                  TabIndex        =   356
                  Top             =   270
                  Width           =   765
                  VariousPropertyBits=   276824083
                  Caption         =   "Frequency "
                  Size            =   "1349;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   19
                  Left            =   3240
                  TabIndex        =   355
                  Top             =   270
                  Width           =   375
                  VariousPropertyBits=   276824083
                  Caption         =   "Fund"
                  Size            =   "661;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   25
                  Left            =   12300
                  TabIndex        =   354
                  Top             =   270
                  Width           =   1080
                  VariousPropertyBits=   276824083
                  Caption         =   "Next Due Date"
                  Size            =   "1905;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   27
                  Left            =   15045
                  TabIndex        =   353
                  Top             =   270
                  Width           =   1335
                  VariousPropertyBits=   8388627
                  Caption         =   "Total Amount/Year"
                  Size            =   "2355;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   20
                  Left            =   4905
                  TabIndex        =   352
                  Top             =   270
                  Width           =   1470
                  VariousPropertyBits=   276824083
                  Caption         =   "Managing Agent A/C"
                  Size            =   "2593;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label7 
                  Height          =   255
                  Index           =   13
                  Left            =   150
                  TabIndex        =   211
                  Top             =   3360
                  Width           =   5880
                  ForeColor       =   4194368
                  BackColor       =   -2147483634
                  VariousPropertyBits=   8388627
                  Size            =   "10372;450"
                  BorderColor     =   -2147483631
                  BorderStyle     =   1
                  FontName        =   "Myriad Condensed Web"
                  FontHeight      =   180
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtEND_DATE 
                  Height          =   285
                  Left            =   20565
                  TabIndex        =   386
                  Top             =   540
                  Width           =   1050
                  VariousPropertyBits=   679495707
                  BorderStyle     =   1
                  Size            =   "1852;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.TextBox txtNtDueDate 
                  Height          =   285
                  Left            =   12330
                  TabIndex        =   379
                  Top             =   540
                  Width           =   1125
                  VariousPropertyBits=   679495711
                  BorderStyle     =   1
                  Size            =   "1984;503"
                  SpecialEffect   =   0
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
                  ParagraphAlign  =   2
               End
            End
            Begin VB.Frame fraPaymentDate 
               Caption         =   "Rent Payable"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3840
               Index           =   16
               Left            =   -74955
               TabIndex        =   201
               Top             =   450
               Width           =   19410
               Begin VB.TextBox txtPayeeType 
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
                  Left            =   7425
                  Locked          =   -1  'True
                  TabIndex        =   227
                  Top             =   495
                  Width           =   1545
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   1
                  Left            =   9000
                  Style           =   1  'Graphical
                  TabIndex        =   228
                  Top             =   495
                  Width           =   345
               End
               Begin VB.CommandButton cmdClose3 
                  Caption         =   "Close"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   17685
                  TabIndex        =   251
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdDeleteRentPayable 
                  Caption         =   "&Delete"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   15030
                  TabIndex        =   248
                  Top             =   3240
                  Width           =   1305
               End
               Begin VB.TextBox txtPercentage 
                  Alignment       =   1  'Right Justify
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
                  Left            =   14310
                  Locked          =   -1  'True
                  MaxLength       =   5
                  TabIndex        =   233
                  Top             =   495
                  Width           =   1365
               End
               Begin VB.TextBox txtPayFund 
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
                  Left            =   4455
                  Locked          =   -1  'True
                  TabIndex        =   225
                  Top             =   495
                  Width           =   2535
               End
               Begin VB.TextBox txtPayableBasis 
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
                  Left            =   11520
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   231
                  Top             =   495
                  Width           =   2310
               End
               Begin VB.TextBox txtPayableType 
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
                  Left            =   135
                  Locked          =   -1  'True
                  TabIndex        =   221
                  Top             =   495
                  Width           =   3885
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   5
                  Left            =   13860
                  Style           =   1  'Graphical
                  TabIndex        =   232
                  Top             =   495
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   3
                  Left            =   11025
                  Style           =   1  'Graphical
                  TabIndex        =   230
                  Top             =   495
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   2
                  Left            =   6975
                  Style           =   1  'Graphical
                  TabIndex        =   226
                  Top             =   495
                  Width           =   345
               End
               Begin VB.CommandButton cmdCommandArray 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   0
                  Left            =   4050
                  Style           =   1  'Graphical
                  TabIndex        =   223
                  Top             =   495
                  Width           =   345
               End
               Begin VB.TextBox txtClientLandlord 
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
                  Left            =   9450
                  Locked          =   -1  'True
                  TabIndex        =   229
                  Top             =   495
                  Width           =   1500
               End
               Begin VB.CommandButton cmdPayCancel 
                  Caption         =   "Cancel"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   16380
                  TabIndex        =   250
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdPayAddNew 
                  Caption         =   "Add New"
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   11205
                  TabIndex        =   242
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdPaySave 
                  Caption         =   "&Save"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   13770
                  TabIndex        =   246
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.CommandButton cmdPayEdit 
                  Caption         =   "&Edit"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Myriad Web"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   12510
                  TabIndex        =   244
                  Top             =   3240
                  Width           =   1215
               End
               Begin VB.TextBox txtStopDate 
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
                  Left            =   15750
                  Locked          =   -1  'True
                  MaxLength       =   12
                  TabIndex        =   234
                  Top             =   495
                  Width           =   1455
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayable 
                  Height          =   2265
                  Left            =   120
                  TabIndex        =   202
                  Top             =   855
                  Width           =   19050
                  _ExtentX        =   33602
                  _ExtentY        =   3995
                  _Version        =   393216
                  ForeColor       =   0
                  Cols            =   6
                  FixedCols       =   0
                  BackColorFixed  =   12632256
                  ForeColorFixed  =   0
                  BackColorSel    =   12648447
                  ForeColorSel    =   4210752
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
                  _Band(0).Cols   =   6
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   4
                  Left            =   7425
                  TabIndex        =   464
                  Top             =   270
                  Width           =   810
                  VariousPropertyBits=   276824083
                  Caption         =   "Payee Type"
                  Size            =   "1429;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   3
                  Left            =   9450
                  TabIndex        =   213
                  Top             =   270
                  Width           =   1155
                  VariousPropertyBits=   276824083
                  Caption         =   "Client/LandLord"
                  Size            =   "2037;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label7 
                  Height          =   300
                  Index           =   14
                  Left            =   90
                  TabIndex        =   208
                  Top             =   3240
                  Width           =   5790
                  ForeColor       =   4194368
                  BackColor       =   -2147483638
                  VariousPropertyBits=   8388627
                  Size            =   "10213;529"
                  BorderColor     =   -2147483636
                  BorderStyle     =   1
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   9
                  Left            =   15750
                  TabIndex        =   207
                  Top             =   270
                  Width           =   720
                  VariousPropertyBits=   276824083
                  Caption         =   "Stop Date"
                  Size            =   "1270;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   2
                  Left            =   4500
                  TabIndex        =   206
                  Top             =   270
                  Width           =   375
                  VariousPropertyBits=   276824083
                  Caption         =   "Fund"
                  Size            =   "661;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   8
                  Left            =   14310
                  TabIndex        =   205
                  Top             =   270
                  Width           =   825
                  VariousPropertyBits=   276824083
                  Caption         =   "Percentage"
                  Size            =   "1455;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   7
                  Left            =   11520
                  TabIndex        =   204
                  Top             =   270
                  Width           =   945
                  VariousPropertyBits=   276824083
                  Caption         =   "Payable Basis"
                  Size            =   "1667;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
               Begin MSForms.Label Label4 
                  Height          =   195
                  Index           =   0
                  Left            =   150
                  TabIndex        =   203
                  Top             =   270
                  Width           =   930
                  VariousPropertyBits=   276824083
                  Caption         =   "Payable Type"
                  Size            =   "1640;344"
                  FontName        =   "Myriad Web"
                  FontHeight      =   165
                  FontCharSet     =   0
                  FontPitchAndFamily=   2
               End
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H80000002&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H8000000F&
               BorderStyle     =   6  'Inside Solid
               DrawMode        =   9  'Not Mask Pen
               FillColor       =   &H8000000F&
               FillStyle       =   0  'Solid
               Height          =   210
               Left            =   0
               Top             =   375
               Width           =   12855
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPropertySelection1 
            Height          =   2400
            Left            =   225
            TabIndex        =   191
            Top             =   765
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   4233
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   12648447
            ForeColorSel    =   4210752
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            BandDisplay     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Condensed Web"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSForms.Label Label4 
            Height          =   195
            Index           =   6
            Left            =   2340
            TabIndex        =   466
            Top             =   450
            Width           =   1170
            VariousPropertyBits=   276824083
            Caption         =   "Property Search:"
            Size            =   "2064;344"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   225
            Index           =   5
            Left            =   360
            TabIndex        =   465
            Top             =   450
            Width           =   1515
            VariousPropertyBits=   276824083
            Caption         =   "Property List:"
            Size            =   "2672;397"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Memo"
         Height          =   3660
         Left            =   -74910
         TabIndex        =   122
         Top             =   360
         Width           =   12705
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10455
            TabIndex        =   134
            Top             =   3180
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8145
            TabIndex        =   133
            Top             =   3180
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   6960
            TabIndex        =   132
            Top             =   3180
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   5940
            TabIndex        =   130
            Top             =   3180
            Width           =   975
         End
         Begin VB.TextBox txtMemoID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11385
            TabIndex        =   129
            Top             =   135
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdVAMemo 
            Caption         =   "&View All Memo"
            Height          =   315
            Left            =   4410
            TabIndex        =   128
            Top             =   3180
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9315
            TabIndex        =   127
            Top             =   3180
            Width           =   1125
         End
         Begin VB.PictureBox fraAllMemo 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   2880
            Left            =   11700
            ScaleHeight     =   2880
            ScaleWidth      =   12555
            TabIndex        =   123
            Top             =   90
            Width           =   12555
            Begin VB.CommandButton cmdCloseMemo 
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
               Left            =   12150
               Style           =   1  'Graphical
               TabIndex        =   124
               Top             =   0
               Width           =   390
            End
            Begin VB.TextBox txtMemoAll 
               Height          =   2550
               Left            =   45
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   125
               Top             =   315
               Width           =   12510
            End
            Begin VB.Shape Shape4 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   240
               Index           =   3
               Left            =   45
               Top             =   30
               Width           =   12105
            End
            Begin MSForms.Label lblSea 
               Height          =   195
               Left            =   180
               TabIndex        =   126
               Top             =   45
               Visible         =   0   'False
               Width           =   1905
               VariousPropertyBits=   8388627
               Caption         =   "Details"
               Size            =   "3360;344"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMemoDetails 
            Height          =   1260
            Left            =   90
            TabIndex        =   135
            Top             =   1845
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   2223
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   -2147483640
            BackColorSel    =   15329508
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   8421504
            WordWrap        =   -1  'True
            HighLight       =   2
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
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
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.TextBox txtUnitMemo 
            Height          =   1335
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   131
            Top             =   210
            Width           =   12510
         End
         Begin VB.Label Label6 
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
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   11
            Left            =   1755
            TabIndex        =   168
            Top             =   1575
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Index           =   10
            Left            =   855
            TabIndex        =   167
            Top             =   1575
            Width           =   345
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No"
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
            Index           =   9
            Left            =   225
            TabIndex        =   166
            Top             =   1575
            Width           =   210
         End
         Begin MSForms.Label Label5 
            Height          =   195
            Left            =   135
            TabIndex        =   137
            Top             =   1575
            Width           =   600
            VariousPropertyBits=   8388627
            Size            =   "1058;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label8 
            Height          =   195
            Left            =   10620
            TabIndex        =   136
            Top             =   1575
            Width           =   1095
            VariousPropertyBits=   8388627
            Caption         =   "User"
            Size            =   "1931;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   -73920
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":0FE0
               Key             =   ""
               Object.Tag             =   "Client"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":18BA
               Key             =   ""
               Object.Tag             =   "Property"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":2194
               Key             =   ""
               Object.Tag             =   "Unit"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":2A6E
               Key             =   ""
               Object.Tag             =   "Lessee"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":38C0
               Key             =   ""
               Object.Tag             =   "Tenant"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew4.frx":3BDA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraBank 
         Caption         =   "Bank Name and Address"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   0
         Left            =   -74880
         TabIndex        =   72
         Top             =   400
         Width           =   7035
         Begin VB.CheckBox chkShowFundBankAccount 
            Appearance      =   0  'Flat
            Caption         =   "View/Select Bank Funds"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4095
            TabIndex        =   488
            Top             =   225
            Width           =   2865
         End
         Begin VB.TextBox txtBANK_POST_CODE 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   1920
            Width           =   1000
         End
         Begin VB.TextBox txtBANK_ADDRESS3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   1560
            Width           =   2715
         End
         Begin VB.TextBox txtBANK_ADDRESS2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   1260
            Width           =   2715
         End
         Begin VB.TextBox txtBANK_ADDRESS1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   960
            Width           =   2715
         End
         Begin VB.TextBox txtBANK_NAME 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   600
            Width           =   2715
         End
         Begin MSAdodcLib.Adodc adoBank 
            Height          =   330
            Left            =   2025
            Top             =   2655
            Visible         =   0   'False
            Width           =   1920
            _ExtentX        =   3387
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
            Caption         =   "Adodc1"
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBankAccountFund 
            Height          =   2490
            Left            =   4050
            TabIndex        =   487
            Top             =   540
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   4392
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
         Begin MSForms.ComboBox cboBank_ID 
            Height          =   285
            Left            =   1200
            TabIndex        =   55
            Top             =   270
            Width           =   2715
            VariousPropertyBits=   1820346399
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4789;503"
            TextColumn      =   1
            ColumnCount     =   6
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "705"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID:"
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
            Index           =   18
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
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
            Index           =   22
            Left            =   120
            TabIndex        =   75
            Top             =   1920
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
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
            Index           =   20
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name:"
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
            Index           =   19
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdUploadImageAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66360
         TabIndex        =   80
         ToolTipText     =   "Add new image"
         Top             =   4800
         Width           =   1035
      End
      Begin VB.CommandButton cmdImgDelete 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -63345
         TabIndex        =   79
         ToolTipText     =   "Delete current image"
         Top             =   4800
         Width           =   1035
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   -74880
         TabIndex        =   77
         Top             =   4110
         Width           =   14985
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOtherBankDetails 
            Height          =   1920
            Left            =   90
            TabIndex        =   78
            Top             =   480
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   3387
            _Version        =   393216
            FixedCols       =   0
            BackColorFixed  =   12632256
            BackColorSel    =   16777215
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
            _Band(0).Cols   =   2
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overdraft Limit"
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
            Index           =   8
            Left            =   12285
            TabIndex        =   120
            Top             =   165
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overdraft"
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
            Index           =   7
            Left            =   11445
            TabIndex        =   119
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID"
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
            Index           =   0
            Left            =   135
            TabIndex        =   90
            Top             =   165
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund Code"
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
            Index           =   2
            Left            =   2385
            TabIndex        =   89
            Top             =   165
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   88
            Top             =   165
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code"
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
            Index           =   5
            Left            =   9525
            TabIndex        =   87
            Top             =   165
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Index           =   3
            Left            =   5040
            TabIndex        =   86
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number"
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
            Index           =   4
            Left            =   7815
            TabIndex        =   85
            Top             =   165
            Width           =   1185
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Ac"
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
            Index           =   6
            Left            =   10605
            TabIndex        =   84
            Top             =   165
            Width           =   735
         End
         Begin VB.Label lblCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000F&
            Height          =   315
            Left            =   75
            TabIndex        =   91
            Top             =   120
            Width           =   14745
         End
      End
      Begin VB.Frame fraType 
         BackColor       =   &H80000016&
         Caption         =   "CLIENT"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -70305
         TabIndex        =   71
         Top             =   360
         Width           =   3720
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   3
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   1
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   920
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   1890
            Width           =   1455
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   2
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1240
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   0
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   28
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
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
            Index           =   76
            Left            =   90
            TabIndex        =   398
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
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
            Index           =   75
            Left            =   90
            TabIndex        =   397
            Top             =   225
            Width           =   435
         End
      End
      Begin VB.Frame fraOccupied 
         BackColor       =   &H80000016&
         Caption         =   "Lease Details:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -70320
         TabIndex        =   70
         Top             =   2640
         Width           =   3735
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   30
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   450
            Width           =   1815
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   29
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   135
            Width           =   1815
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   31
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   765
            Width           =   1815
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   32
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   33
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
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
            Index           =   34
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
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
            Index           =   82
            Left            =   135
            TabIndex        =   404
            Top             =   1845
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lessee Type:"
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
            Index           =   81
            Left            =   135
            TabIndex        =   403
            Top             =   1485
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
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
            Index           =   80
            Left            =   135
            TabIndex        =   402
            Top             =   1125
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
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
            Index           =   79
            Left            =   135
            TabIndex        =   401
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lessee Name:"
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
            Index           =   78
            Left            =   135
            TabIndex        =   400
            Top             =   495
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lessee ID:"
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
            Index           =   77
            Left            =   135
            TabIndex        =   399
            Top             =   225
            Width           =   720
         End
      End
      Begin MSComctlLib.TreeView tvwLandLord 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8493
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistory 
         Height          =   3945
         Left            =   -74925
         TabIndex        =   92
         Top             =   1170
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   6959
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxACHistorySplit 
         Height          =   1965
         Left            =   -74910
         TabIndex        =   101
         Top             =   5475
         Width           =   15570
         _ExtentX        =   27464
         _ExtentY        =   3466
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
         ScrollBars      =   2
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
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account ID"
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
         Index           =   22
         Left            =   -73740
         TabIndex        =   515
         Top             =   945
         Width           =   780
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Download To CSV"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   0
         Left            =   -66225
         TabIndex        =   512
         Top             =   900
         Width           =   1770
      End
      Begin MSForms.TextBox txtFilterClient 
         Height          =   315
         Left            =   -73785
         TabIndex        =   498
         Tag             =   "Client"
         Top             =   450
         Width           =   1950
         VariousPropertyBits=   746604575
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "3440;556"
         Value           =   "Clients"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By Type"
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
         Index           =   92
         Left            =   -74865
         TabIndex        =   497
         Top             =   495
         Width           =   960
      End
      Begin MSForms.TextBox txtACBalanceByCl 
         Height          =   315
         Left            =   -62400
         TabIndex        =   496
         Top             =   450
         Width           =   1515
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2672;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance by Client"
         Height          =   210
         Index           =   91
         Left            =   -64695
         TabIndex        =   495
         Top             =   495
         Width           =   2235
      End
      Begin MSForms.TextBox txtSupplierFilter 
         Height          =   315
         Left            =   -71310
         TabIndex        =   494
         Top             =   450
         Width           =   1590
         VariousPropertyBits=   746604575
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "2805;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchRef1 
         Height          =   315
         Left            =   -68880
         TabIndex        =   493
         TabStop         =   0   'False
         Top             =   450
         Width           =   1365
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "2408;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref"
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
         Index           =   90
         Left            =   -69285
         TabIndex        =   492
         Top             =   495
         Width           =   315
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
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
         Index           =   5
         Left            =   -68190
         TabIndex        =   115
         Top             =   945
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
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
         Index           =   10
         Left            =   -74835
         TabIndex        =   113
         Top             =   5235
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   11
         Left            =   -73710
         TabIndex        =   112
         Top             =   5235
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   12
         Left            =   -72750
         TabIndex        =   111
         Top             =   5235
         Width           =   345
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
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
         Index           =   13
         Left            =   -71790
         TabIndex        =   110
         Top             =   5235
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Index           =   21
         Left            =   -61500
         TabIndex        =   109
         Top             =   5235
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
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
         Height          =   195
         Index           =   19
         Left            =   -63660
         TabIndex        =   108
         Top             =   5235
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
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
         Index           =   14
         Left            =   -70590
         TabIndex        =   107
         Top             =   5235
         Width           =   825
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
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
         Index           =   20
         Left            =   -62580
         TabIndex        =   106
         Top             =   5235
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No"
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
         Index           =   15
         Left            =   -69750
         TabIndex        =   105
         Top             =   5235
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop No"
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
         Index           =   16
         Left            =   -69030
         TabIndex        =   104
         Top             =   5235
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
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
         Height          =   195
         Index           =   17
         Left            =   -68070
         TabIndex        =   103
         Top             =   5235
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Index           =   18
         Left            =   -67110
         TabIndex        =   102
         Top             =   5235
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
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
         Left            =   -74700
         TabIndex        =   100
         Top             =   945
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   -72510
         TabIndex        =   99
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
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
         Left            =   -69990
         TabIndex        =   98
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Index           =   6
         Left            =   -64395
         TabIndex        =   97
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Index           =   7
         Left            =   -63195
         TabIndex        =   96
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
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
         Index           =   8
         Left            =   -61995
         TabIndex        =   95
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Index           =   9
         Left            =   -60795
         TabIndex        =   94
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   -71070
         TabIndex        =   93
         Top             =   960
         Width           =   345
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000015&
         Height          =   3675
         Index           =   4
         Left            =   -74940
         Top             =   360
         Width           =   12825
      End
      Begin VB.Image imgPremises 
         BorderStyle     =   1  'Fixed Single
         Height          =   3930
         Left            =   -66360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   4050
      End
      Begin MSForms.CommandButton cmdImgLeftMove 
         Height          =   420
         Left            =   -64853
         TabIndex        =   83
         ToolTipText     =   "Next image"
         Top             =   4800
         Width           =   1035
         Caption         =   "Next"
         PicturePosition =   196613
         Size            =   "1826;741"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label lblImageName 
         Height          =   195
         Left            =   -66360
         TabIndex        =   82
         Top             =   360
         Width           =   1080
         Caption         =   "Image Name:"
         Size            =   "1905;344"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblGridCaption 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -74925
         TabIndex        =   114
         Top             =   5145
         Width           =   15615
      End
   End
   Begin VB.PictureBox picMain 
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
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   45
      ScaleHeight     =   915
      ScaleWidth      =   22905
      TabIndex        =   61
      Top             =   90
      Width           =   22935
      Begin VB.CheckBox chkOptedtoTax 
         BackColor       =   &H80000009&
         Caption         =   "Check1"
         Height          =   225
         Left            =   14940
         TabIndex        =   7
         Top             =   90
         Width           =   195
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   15165
         Locked          =   -1  'True
         TabIndex        =   169
         Top             =   90
         Width           =   765
      End
      Begin VB.CommandButton cmdCTSec 
         Caption         =   "..."
         Enabled         =   0   'False
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
         Left            =   13035
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   345
      End
      Begin VB.ListBox lstCT 
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
         Height          =   420
         ItemData        =   "frmClientNew4.frx":3EF4
         Left            =   11640
         List            =   "frmClientNew4.frx":3EFE
         TabIndex        =   117
         Top             =   960
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtCT 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   480
         Width           =   3045
      End
      Begin VB.CommandButton cmdVAT 
         Caption         =   "..."
         Enabled         =   0   'False
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
         Left            =   15975
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   345
      End
      Begin VB.CommandButton cmdClient 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3165
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   345
      End
      Begin VB.TextBox txtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   885
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   2355
      End
      Begin VB.TextBox txtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   885
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   90
         Width           =   2620
      End
      Begin VB.TextBox txtYearEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtVATReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5895
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   3
         Top             =   135
         Width           =   1695
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Opted to Tax:"
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
         Index           =   12
         Left            =   13545
         TabIndex        =   170
         Top             =   135
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Management Type:"
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
         Left            =   8625
         TabIndex        =   118
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year End:"
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
         Index           =   7
         Left            =   4395
         TabIndex        =   66
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX/VAT Number:"
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
         Index           =   6
         Left            =   4365
         TabIndex        =   65
         Top             =   120
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
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
         Left            =   8625
         TabIndex        =   64
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
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
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   660
      End
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
      Left            =   4403
      ScaleHeight     =   390
      ScaleWidth      =   3255
      TabIndex        =   67
      Top             =   3420
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   68
         Top             =   90
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmClientNew4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private shlShell As shell32.Shell
Private shlFolder As shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1
'How to fix the bug on / when you type a date field 1. txtEND_DATE_Change event add conditional exit sub  2. txtEND_DATE_KeyDown e all codes has be written  3. declare bBackSp variable for this form
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Const SB_VERT = 1

Public LOAD_CLINT_CLIENTID As String

Private bDefaultAccount As Boolean, iTotalBankAC As Integer, lDefaultBankID As Long
'Private szPropertyID As String
Private iRecharge As Integer, iSlectedRow As Integer
Private bGlobalData As Boolean
Private bNewEdit As Boolean, bBankNewEdit As Boolean
Private IMAGE_FILE_NAME_ As String
Private szaPremisisIDType() As String
'Dim sText As String
Dim szaSupplierBalanceCL() As String
Private ADD_NEW_CLIENT As Boolean
'Private AGREEMENT_EDIT_MODE As Boolean
Private AGREEMENT_ADDNEW_MODE As Boolean
Private PAYABLE_ADDNEW_MODE As Boolean
'Private PAYABLE_EDIT_MODE As Boolean
Private NEW_TYPE As String
Dim szaClientBal()      As String      'Client     balance
Private bOverdraftWarning As Boolean
Dim Memo_Save_mode As Boolean
Dim strCommandSource  As String
Dim szPropertySelection1 As String
Dim szPropertySelection2 As String
Dim strtlbPayableID As Integer
Dim strtlbAgreementID As Integer
Dim dtFDD      As Date
Dim strSelectedFundName As String
Dim bBackSp As Boolean
Dim intConsolidatedBankID As Integer
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single
Dim sessionID As String
Dim reportingDate As String
Dim rRow As Integer 'Keeps global  flxmgtfee selected row number
Private Sub LoadFlxACHistory(adoConn As ADODB.Connection, Filter As String)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset
   Dim iLoop As Long
   Dim strWhere As String
   Dim rsReportClientHistory As New ADODB.Recordset
   Dim rsDiffClients As New ADODB.Recordset
   Dim tempstr As String
   Dim ifNeedBuildingBal1 As Boolean
   Dim ifNeedBuildingBal2 As Boolean
   fmeLoading.Visible = True
   fmeLoading.Refresh
   ConfigFlxACHistory
   Dim StrWhere2 As String
   Dim strWhere3 As String
   Dim amtBalance As Double
   Dim amtcr As Double
   'The ralationship between payment and invoice is One Batch payment can be spread over many Invoice to allocate
   'below method need to be updated, I cannot get the outstading amount from the table as an invoice can be paid of from diferrent client
   'Then need to to consider that payment in case user wants to see the filtered PI and Payment only from this client

   'I did not find any partialy allocated payments:( SELECT *FROM tlbPayment where type in (7,8,9) and osamount>0 and amount>osamount;) IN WPM
   ' I have found invoices which is partially paid IN WPM

'   szSQL = "delete from  ReportClientHistory ;" 'WHERE SessionID = '" & sessionID & "'
'   adoConn.Execute szSQL

   szSQL = "delete from  ReportClientHistory;"
   adoConn.Execute szSQL

   If Filter = "1" Then
        If txtSearchNo.text <> "" Then
            tempstr = Replace(UCase(txtSearchNo.text), "'", "''")
            strWhere = " AND PF Like '%" & tempstr & "%'" 'PF=Inovice number/Receipt number
        End If
    End If
    If Filter = "2" Then
         If txtSearchRef.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef.text), "'", "''")
            strWhere = " AND Extref Like '%" & tempstr & "%'"
        End If
    End If
    If Filter = "3" Or Filter = "4" Then
         If txtSearchFromD.text <> "" And txtSearchToD.text <> "" Then
            strWhere = " AND RDate >=#" & Format(txtSearchFromD.text, "dd/mmm/yyyy") & "# AND RDate <=#" & Format(txtSearchToD.text, "dd/mmm/yyyy") & "# "
            If Len(txtSearchFromD.text) > 0 And Len(txtSearchToD.text) > 0 Then
                 cmdSearch.Caption = "Clear Sea&rch"
            Else
                 cmdSearch.Caption = "Sea&rch"
            End If
        End If
   End If
   If Filter = "5" Then
         If txtSearchRef1.text <> "" Then
            tempstr = Replace(UCase(txtSearchRef1.text), "'", "''")
            strWhere = " AND Extref Like '%" & tempstr & "%'"
        End If
    End If
     '* This is for lessse history
     If txtFilterClient.Tag = "Lessee" Then
                 szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, '',R.transactionID,R.Type,TT.DESCRIPTION ," & _
                             "(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)) & R.SlNumber AS INVno,Rdate as PDate,Details ,extref,amount,osamount,0,0,R.ClientID,R.SageAccountNumber  FROM " & _
                             "(tlbReceipt AS R INNER JOIN tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID) INNER JOIN Tenants ON Tenants.SageAccountNumber=R.Sageaccountnumber " & _
                             "where ClientID= '" & txtClientID.text & "' ORDER BY R.TransactionID;" 'we do not need supplier type because we are making join with tenant which shall filter itself type
                             
                 szSQL = "insert into ReportClientHistory (reportingDate,SessionID,SIGN,transactionID,Type,Type_desc,PF,Pdate,Details ,extref ,amount,Osamount, flag ,isMaster,ClientID,SageAccountNumber) " & _
                      szSQL
                adoConn.Execute szSQL
                
                adoConn.Execute "Update ReportClientHistory A INNER JOIN RptTransactions R ON A.transactionID=R.FromTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (2,3,4)"
                adoConn.Execute "Update ReportClientHistory A INNER JOIN RptTransactions R ON A.transactionID=R.ToTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (1,2)"
                
                'Now allocation slave for invoices
                'We are not using serial number later on anywhere in report table
                szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
                        "'-', R.TransactionID, R.Type,R.PF,RT.Allocdate,RT.ReceiptAmount,'1',R.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO,R.SageAccountNumber" & _
                        " FROM ((ReportClientHistory R  INNER JOIN RptTransactions  RT ON R.transactionID = RT.ToTran)" & _
                        " INNER JOIN tlbReceipt P ON RT.FromTran = P.TransactionID) INNER JOIN tlbTransactionTypes T ON P.Type = T.TYPE_ID where RT.DeleteFlag=false AND SessionID = '" & sessionID & "' AND R.Type In(1,2)" & strWhere3
            
                szSQL = "insert into ReportClientHistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,PDate,amount,isMaster,ClientID,ActualINV,SageAccountNumber) " & _
                        szSQL
                adoConn.Execute szSQL
              'allocation slave for Receipts
            '   Exit Sub
                szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
                        "'-', R.TransactionID, P.Type, R.PF,PT.Allocdate,PT.ReceiptAmount,'1',P.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO,R.SageAccountNumber " & _
                        "FROM (ReportClientHistory AS R INNER JOIN RptTransactions AS PT ON R.transactionID = PT.FromTran)" & _
                        "INNER JOIN (tlbReceipt AS P INNER JOIN tlbTransactionTypes AS T ON P.Type = T.TYPE_ID) ON PT.ToTran = P.TransactionID where PT.DeleteFlag=false AND SessionID = '" & sessionID & "' AND R.Type In(2,3,4)" & strWhere3
                 szSQL = "insert into ReportClientHistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,Pdate,amount,isMaster,ClientID,ActualINV,SageAccountNumber) " & _
                        szSQL
                adoConn.Execute szSQL
    
    
    Else
    
    
    
    '* Usual tlbpayment table using starts here for normal client history
               'ActualINV is only usefull to show to the slave records
                'If txtFilterClient.Tag = "ALL" Then
                    szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, '',P.transactionID,P.Type,TT.DESCRIPTION ," & _
                             "(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)) & P.SlNumber AS INVno,Pdate,Details ,extref,amount,osamount,0,0,P.ClientID,P.SageAccountNumber  FROM " & _
                             "(tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON P.Type = TT.TYPE_ID) INNER JOIN Supplier ON Supplier.SupplierID=P.Sageaccountnumber " & _
                             "where ClientID= '" & txtClientID.text & "' AND Supplier.Type='" & txtFilterClient.Tag & "' ORDER BY P.TransactionID;" 'AND Supplier.Type='" & txtFilterClient.Tag & "'
            
            '    Else
            '             # specail case 1
            '
            '             we had faced a problem finding outstanding balance when there was two different clients, .i.e
            '             PI in one client that has been paid from another client. So balance needs to be build again.
            '             I have following SQL to check that condition if we need to build balance again or not.
            '             if you select all client then there is no problem. Other you need to check this problem and call this section
            '             szSQL = "SELECT P.ClientID, P1.ClientID FROM (tlbPayment P INNER JOIN PayTransactions T ON P.TransactionID = T.FromTran) INNER JOIN tlbPayment " & _
            '                     "AS P1 ON T.ToTran = P1.TransactionID where P.ClientID<> P1.ClientID and P.SageAccountNumber = '" & txtFilterClient.Tag & "' AND  P1.ClientID = '" & txtClientID.Tag & "';"
            '             rsDiffClients.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
            '             If Not rsDiffClients.EOF Then
            '                 ifNeedBuildingBal1 = True
            '             End If
            '             rsDiffClients.Close
            '             szSQL = "SELECT P.ClientID, P1.ClientID FROM (tlbPayment P INNER JOIN PayTransactions T ON P.TransactionID = T.FromTran) INNER JOIN tlbPayment " & _
            '                     "AS P1 ON T.ToTran = P1.TransactionID where P.ClientID<> P1.ClientID and P1.SageAccountNumber = '" & txtFilterClient.Tag & "' AND  P.ClientID = '" & txtClientID.Tag & "';"
            '             rsDiffClients.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
            '             If Not rsDiffClients.EOF Then
            '                 ifNeedBuildingBal2 = True
            '             End If
            '             rsDiffClients.Close
            '    Dim strcorrection As String
            '        If txtFilterClient.text <> "ALL" Then
            '                strcorrection = "P.SageAccountNumber = '" & txtFilterClient.Tag & "' AND"
            '        End If
            '
            '        szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, '',P.transactionID,P.Type,TT.DESCRIPTION ," & _
            '                 "(MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3)) & P.SlNumber AS INVno,Pdate,Details ,extref,amount,osamount,0,0,P.ClientID  FROM " & _
            '                 "(tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON P.Type = TT.TYPE_ID) " & _
            '                 "WHERE " & strcorrection & "  P.ClientID = '" & txtClientID.Tag & "' ORDER BY P.TransactionID;"
            '    End If
                If txtFilterClient.Tag <> "ALL" Then
                    strWhere3 = " AND  P.ClientID = '" & txtClientID.text & "'"
                End If
                szSQL = "insert into ReportClientHistory (reportingDate,SessionID,SIGN,transactionID,Type,Type_desc,PF,Pdate,Details ,extref ,amount,Osamount, flag ,isMaster,ClientID,SageAccountNumber) " & _
                      szSQL
                adoConn.Execute szSQL
                'Exit Sub
            '    and P.Type in (7,8,9)
             'Update allocation sign
            
                'If txtFilterClient.Tag = "ALL" Then
                        adoConn.Execute "Update ReportClientHistory A INNER JOIN PayTransactions R ON A.transactionID=R.FromTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (7,8,9)"
                        adoConn.Execute "Update ReportClientHistory A INNER JOIN PayTransactions R ON A.transactionID=R.ToTran SET SIGN='+' where SessionID = '" & sessionID & "' AND Type in (6,24)"
            '    Else
            '            adoconn.Execute "UPDATE (ReportClientHistory AS A INNER JOIN PayTransactions AS R ON A.transactionID = R.FromTran) INNER JOIN tlbPayment P ON R.ToTran = P.TransactionID SET A.SIGN = '+' WHERE SessionID = '" & sessionID & "' AND (A.Type In (7,8,9))" & strWhere3
            '            adoconn.Execute "UPDATE tlbPayment P INNER JOIN (ReportClientHistory AS A INNER JOIN PayTransactions AS R ON A.transactionID = R.ToTran) ON P.TransactionID = R.FromTran SET A.SIGN = '+'  WHERE SessionID = '" & sessionID & "' AND (A.Type) In (6,24)" & strWhere3
            '    End If
                'Insertion of Parent records has been completed
            
                'Now allocation slave for invoices
                'We are not using serial number later on anywhere in report table
                szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
                        "'-', R.TransactionID, P.Type,R.PF,PT.Allocdate,PT.PaymentAmount,'1',P.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO,R.SageAccountNumber" & _
                        " FROM ((ReportClientHistory R  INNER JOIN PayTransactions  PT ON R.transactionID = PT.ToTran)" & _
                        " INNER JOIN tlbPayment P ON PT.FromTran = P.TransactionID) INNER JOIN tlbTransactionTypes T ON P.Type = T.TYPE_ID where SessionID = '" & sessionID & "' AND R.Type In(6,24)" & strWhere3
            
                szSQL = "insert into ReportClientHistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,PDate,amount,isMaster,ClientID,ActualINV,SageAccountNumber) " & _
                        szSQL
                adoConn.Execute szSQL
              'allocation slave for Receipts
            '   Exit Sub
                szSQL = "SELECT '" & reportingDate & "' AS ReportingDate, '" & sessionID & "' AS SessionID, " & _
                        "'-', R.TransactionID, P.Type, R.PF,PT.Allocdate,PT.PaymentAmount,'1',P.ClientID,(Mid(T.CONSTANT,4,Len(T.CONSTANT)-3) & P.SlNumber) AS INVNO,R.SageAccountNumber " & _
                        "FROM (ReportClientHistory AS R INNER JOIN PayTransactions AS PT ON R.transactionID = PT.FromTran)" & _
                        "INNER JOIN (tlbPayment AS P INNER JOIN tlbTransactionTypes AS T ON P.Type = T.TYPE_ID) ON PT.ToTran = P.TransactionID where SessionID = '" & sessionID & "' AND R.Type In(7,8,9)" & strWhere3
                 szSQL = "insert into ReportClientHistory(reportingDate,SessionID,SIGN,transactionID,Type,PF,Pdate,amount,isMaster,ClientID,ActualINV,SageAccountNumber) " & _
                        szSQL
                adoConn.Execute szSQL
            
            
                Dim PaymentBalance As Double
            
            
                If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
                        adoPty.Open "Select * from ReportClientHistory where SessionID = '" & sessionID & "'  order by transactionID desc,ismaster desc", adoConn, adOpenDynamic, adLockOptimistic
                        Debug.Print time
                        While Not adoPty.EOF
                            If adoPty("Sign").Value = "-" Then
                              PaymentBalance = PaymentBalance + adoPty("Amount").Value
                           End If
                           If adoPty("Sign").Value = "+" Then
                                adoPty("Balance") = adoPty("Amount") - PaymentBalance 'Balance
                                adoPty("Balance") = Round(adoPty("OSAmount"), 2)
                                PaymentBalance = 0
            
                           End If
                           If adoPty("Sign").Value = "" Then
                                 adoPty("Balance") = adoPty("Amount")
                                 PaymentBalance = 0
                           End If
                           adoPty.Update
                           adoPty.MoveNext
                       Wend
                    '    adoPty.UpdateBatch adAffectAllChapters
                        adoPty.Close
                End If
    End If
    If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
            StrWhere2 = " AND Balance>0 "
    Else
            StrWhere2 = " AND OSAmount>0 "
    End If
    If chkShowOutstanding.Value = 0 Then
        'rsReportClientHistory.Open "Select * from ReportLAChistory where 1=1 " & strWhere & "  order by transactionID,ismaster", adoconn, adOpenStatic, adLockReadOnly
        adoConn.Execute "Update  ReportClientHistory A, (Select transactionID from ReportClientHistory where  SessionID= '" & sessionID & "'" & _
                         strWhere & " order by transactionID,ismaster) As B Set flag=1 where A.transactionID=B.transactionID"
    Else
        adoConn.Execute "Update  ReportClientHistory A, (Select transactionID from ReportClientHistory where SessionID= '" & sessionID & "'" & StrWhere2 & _
                         strWhere & " order by transactionID,ismaster) As B Set flag=1 where A.transactionID=B.transactionID"

    End If
    If Filter = "6" And txtSupplierFilter.text <> "ALL" Then 'when there is multiple filter first by client and second by client
            adoConn.Execute "Delete from ReportClientHistory where sageaccountnumber<>'" & txtSupplierFilter.text & "' and  SessionID= '" & sessionID & "'"
    End If
    rsReportClientHistory.Open "Select * from ReportClientHistory where flag=1 and SessionID= '" & sessionID & "' order by transactionID ,ismaster ", adoConn, adOpenKeyset, adLockReadOnly
    'rsReportClientHistory.Close
    
    If rsReportClientHistory.RecordCount = 0 Then
         flxACHistory.Rows = 2
    Else
         flxACHistory.Rows = rsReportClientHistory.RecordCount + 1
    End If
    iKount = 1
    With flxACHistory
    While Not rsReportClientHistory.EOF
        .TextMatrix(iKount, 0) = rsReportClientHistory("SIGN").Value
        .TextMatrix(iKount, 2) = rsReportClientHistory("SageAccountNumber").Value
        .TextMatrix(iKount, 3) = IIf(IsNull(rsReportClientHistory("Type_desc").Value), "", rsReportClientHistory("Type_desc").Value)
        If InStr(.TextMatrix(iKount, 3), "Purchase") > 0 Then .TextMatrix(iKount, 3) = Mid(.TextMatrix(iKount, 3), 10)
        If InStr(.TextMatrix(iKount, 3), "Payment") > 0 And InStr(.TextMatrix(iKount, 3), "Account") = 0 Then .TextMatrix(iKount, 3) = "Payment"
        If InStr(.TextMatrix(iKount, 3), "Account") > 0 Then .TextMatrix(iKount, 3) = "Payment on A/C"
        If InStr(.TextMatrix(iKount, 3), "Invoice") > 0 Then .TextMatrix(iKount, 3) = "Invoice"
        .TextMatrix(iKount, 4) = IIf(IsNull(rsReportClientHistory("PDate").Value), "", rsReportClientHistory("PDate").Value)
        If rsReportClientHistory("SIGN").Value = "-" Then
            .RowHeight(iKount) = 0
            '.TextMatrix(iKount, 4) = "Receipt to " & rsReportClientHistory("PF").Value
            '.TextMatrix(iKount, 5) = "Receipt From " & rsReportClientHistory("PF").Value
            .TextMatrix(iKount, 1) = rsReportClientHistory("ActualINV").Value
             If rsReportClientHistory("Type").Value = 6 Or rsReportClientHistory("Type").Value = 24 Then
                .TextMatrix(iKount, 6) = "Payment to: " & rsReportClientHistory.Fields.Item("ActualINV").Value 'from i sreveresrse because slave has reverse PF
             Else
                .TextMatrix(iKount, 6) = "Payment From: " & rsReportClientHistory.Fields.Item("ActualINV").Value
             End If
        Else
            .RowHeight(iKount) = 260
            .TextMatrix(iKount, 1) = rsReportClientHistory("PF").Value
            .TextMatrix(iKount, 5) = IIf(IsNull(rsReportClientHistory("Extref").Value), "", rsReportClientHistory("Extref").Value)
            .TextMatrix(iKount, 6) = IIf(IsNull(rsReportClientHistory("Details").Value), "", rsReportClientHistory("Details").Value)

        End If
        .TextMatrix(iKount, 7) = IIf(rsReportClientHistory("Amount").Value = 0, "", Format(rsReportClientHistory("Amount").Value, "0.00"))
        If ifNeedBuildingBal1 = True Or ifNeedBuildingBal2 = True Then
            .TextMatrix(iKount, 8) = IIf(rsReportClientHistory("Balance").Value = 0, "", Format(rsReportClientHistory("Balance").Value, "0.00"))
        Else
            .TextMatrix(iKount, 8) = IIf(rsReportClientHistory("OSamount").Value = 0, "", Format(rsReportClientHistory("OSamount").Value, "0.00"))
        End If
        If rsReportClientHistory("SIGN").Value <> "-" Then 'for the allocated amount you don't need debit or credit row
            If rsReportClientHistory("Type").Value = 6 Or rsReportClientHistory("Type").Value = 24 Then
                .TextMatrix(iKount, 10) = Format(rsReportClientHistory("Amount").Value, "0.00")
            Else
                .TextMatrix(iKount, 9) = Format(rsReportClientHistory("Amount").Value, "0.00")
            End If
        End If
        .TextMatrix(iKount, 11) = rsReportClientHistory("transactionID").Value
'        iKount = iKount + 1
        rsReportClientHistory.MoveNext
        amtBalance = amtBalance + Val(flxACHistory.TextMatrix(iKount, 10))
        amtcr = amtcr + Val(flxACHistory.TextMatrix(iKount, 9))
        iKount = iKount + 1
   Wend
   End With
   rsReportClientHistory.Close
   Set rsReportClientHistory = Nothing
   fmeLoading.Visible = False
   fmeLoading.Refresh
  ' MsgBox "Debit :" & amtBalance & "Credit :" & amtcr
    'adoConn.Execute "Delete from ReportClientHistory where SessionID= '" & sessionID & "'"
   ' MsgBox flxACHistory.Rows
  End Sub
Private Sub SaveSizes() ' Save the form's and controls' dimensions.
    Dim i As Integer
    Dim ctl As Control
    On Error Resume Next
    'If m_FormWid = 0 Then Exit Sub
    ' Save the controls' positions and sizes.
    ReDim m_ControlPositions(1 To Controls.Count)
    i = 1
    For Each ctl In Controls
        With m_ControlPositions(i)
            If TypeOf ctl Is Line Then
                .Left = ctl.x1
                .Top = ctl.Y1
                .Width = ctl.x2 - ctl.x1
                .Height = ctl.Y2 - ctl.Y1
            Else
                .Left = ctl.Left
                .Top = ctl.Top
                .Width = ctl.Width
                .Height = ctl.Height
                On Error Resume Next
                .FontSize = ctl.Font.Size
                On Error GoTo 0
            End If
        End With
        i = i + 1
Next
    ' Save the form's size.
    m_FormWid = ScaleWidth
    m_FormHgt = ScaleHeight
End Sub
Private Sub ResizeControls()
        ' Arrange the controls for the new size.
        If m_FormWid = 0 Then Exit Sub
        Dim i As Integer
        Dim ctl As Control
        Dim X_Scale As Single
        Dim Y_Scale As Single
        
        ' Don't bother if we are minimized.
        If WindowState = vbMinimized Then Exit Sub
        
        ' Get the form's current scale factors.
        X_Scale = ScaleWidth / m_FormWid
        Y_Scale = ScaleHeight / m_FormHgt
        
        ' Position the controls.
        i = 1
        For Each ctl In Controls
            With m_ControlPositions(i)
                If TypeOf ctl Is Line Then
                    ctl.x1 = X_Scale * .Left
                    ctl.Y1 = Y_Scale * .Top
                    ctl.x2 = ctl.x1 + X_Scale * .Width
                    ctl.Y2 = ctl.Y1 + Y_Scale * .Height
                Else
                    ctl.Left = X_Scale * .Left
                    ctl.Top = Y_Scale * .Top
                    ctl.Width = X_Scale * .Width
                    If Not (TypeOf ctl Is ComboBox) Then
                        ' Cannot change height of ComboBoxes.
                        ctl.Height = Y_Scale * .Height
                    End If
                    On Error Resume Next
                    ctl.Font.Size = Y_Scale * .FontSize
                    On Error GoTo 0
                End If
            End With
            i = i + 1
        Next
End Sub

Private Sub chkShowFundBankAccount_Click()
    If chkShowFundBankAccount.Value = 1 Then
        flxBankAccountFund.Visible = True
        Call LoadflxBankAccountFund
    Else
        flxBankAccountFund.Visible = False
    End If
End Sub

Private Sub chkShowOutstanding_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    'Debug.Print time
    Call LoadFlxACHistory(adoConn, "")
    'Debug.Print time
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdAdvanceProgr_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    adoConn.Execute "Update tlbagreement T,ClientProAgr P set T.ReportClientID=P.ClientID,T.ReportPropertyID=P.PropertyID where T.CPA_ID=P.CPA_ID"
    adoConn.Close
    MsgBox "ReportClientID,ReportPropertyID has been updated in tlbagreement table"
End Sub

Private Sub cmdBrowsFile_Click(Index As Integer)
   Dim ofn As OPENFILENAME
   Dim lHwnd As Long
   Const HKEY_LOCAL_MACHINE As Long = &H80000002
   Dim szOldFile_PathName As String
   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
   Dim fso As Object

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "All Files (*.rpt)" + Chr$(0) + "*.rpt" + Chr$(0)
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir & "\CompanyReports"
   ofn.lpstrTitle = "Select a Report file"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   'If Index = 2 Then txtComments2(1).text = JustifyFilePath(ofn.lpstrFileTitle)
   cmdClientAddAtch(10).Enabled = True
   If Index = 2 Then
        txtComments2(1).text = JustifyFilePath(ofn.lpstrFileTitle)
        FocusControl cmdClientAddAtch(10)
   End If
   If Index = 0 Then
        txtComments2(2).text = JustifyFilePath(ofn.lpstrFileTitle)
        FocusControl cmdClientAddAtch(13)
   End If
   If Index = 1 Then
        txtComments2(3).text = JustifyFilePath(ofn.lpstrFileTitle)
        FocusControl cmdClientAddAtch(16)
   End If
   If Index = 3 Then
        txtComments2(4).text = JustifyFilePath(ofn.lpstrFileTitle)
        FocusControl cmdClientAddAtch(19)
   End If
End Sub
Private Sub UpdateBalanceCL()
   Dim i As Integer, j As Integer

   For i = 1 To flxClientList.Rows - 1
      For j = 0 To UBound(szaSupplierBalanceCL, 2) - 1
         If flxClientList.TextMatrix(i, 1) = szaSupplierBalanceCL(0, j) Then
            flxClientList.TextMatrix(i, 3) = Format(szaSupplierBalanceCL(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalanceCL, 2) Then flxClientList.TextMatrix(i, 3) = "0.00"
   Next i
End Sub
Private Function SupplierAccountBalanceALLClient(adoConn As ADODB.Connection, szSuppID As String) As Currency
'I am using this function when Showing supplier balance for all client
'function written by anol 20170724
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

 

    szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE (Type = 6 OR Type = 24) and SageAccountNumber='" & szSuppID & "' " & _
           "GROUP BY SageAccountNumber;"


   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

  
   While Not adoPayDr.EOF
      
      SupplierAccountBalanceALLClient = IIf(IsNull(adoPayDr.Fields.Item("Dr").Value), 0, adoPayDr.Fields.Item("Dr").Value)
      
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

  szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 and SageAccountNumber='" & szSuppID & "'" & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      
        SupplierAccountBalanceALLClient = SupplierAccountBalanceALLClient - IIf(IsNull(adoPayCr.Fields.Item("Cr").Value), 0, adoPayCr.Fields.Item("Cr").Value)
      
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Function

Private Function SupplierAccountBalanceByClient2(adoConn As ADODB.Connection, szSuppID As String) As Currency
    
'Build Supplier AC balance by client by anol 20180913 this balance is for only one supplier

   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT ClientID AS SageAccountNumber " & _
           "FROM Client;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSupplierBalanceCL(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalanceCL(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      If "MARIACED" = adoPayDr.Fields.Item("SageAccountNumber").Value Then
'                    MsgBox adoPayDr.Fields.Item("SageAccountNumber").Value
'       End If
      szaSupplierBalanceCL(1, iIndex) = 0
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close
   '6
   'New section1
   szSQL = "SELECT Type, ClientID,Type,Sum(Amount) AS Amt " & _
                 "FROM tlbPayment where SageAccountNumber='" & szSuppID & "' group by  Type, ClientID ;"

'Debug.Print szSQL
'adoPayCr.Close
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalanceCL(0, i) = adoPayCr.Fields.Item("ClientID").Value Then
'            If "N19PROPE" = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'                    Debug.Print adoPayCr.Fields.Item("SageAccountNumber").Value
'            End If
            If adoPayCr.Fields.Item("Type").Value = 6 Or adoPayCr.Fields.Item("Type").Value = 24 Then
                 szaSupplierBalanceCL(1, i) = Val(szaSupplierBalanceCL(1, i)) + adoPayCr.Fields.Item("Amt").Value
            End If
            If adoPayCr.Fields.Item("Type").Value = 8 Or adoPayCr.Fields.Item("Type").Value = 7 Or adoPayCr.Fields.Item("Type").Value = 9 Then
                 szaSupplierBalanceCL(1, i) = Val(szaSupplierBalanceCL(1, i)) - adoPayCr.Fields.Item("Amt").Value
            End If
         End If
      Next i
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

End Function
Private Sub LoadClients()

'you just change label position then searchbox and grid column will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString
   
   
   'SupplierAccountBalanceByClient2 Conn2, txtSupplierFilter.text  'load client wise balance with single supplier
'   If strCommandSource = "ClientFiltet" Then
'        If Len(txtSearchClientName.text) > 0 Then
'             szSQL = "SELECT  ClientID, ClientName FROM Client where  ClientName like'%" & txtSearchClientName.text & "%' ORDER BY ClientID ;"
'        ElseIf Len(txtSearchClientID.text) > 0 Then
'             szSQL = "SELECT  ClientID, ClientName FROM Client where ClientID like'%" & txtSearchClientID.text & "%' ORDER BY ClientID ;"
'        Else
'             szSQL = "SELECT ClientID, ClientName " & _
'                "FROM Client order by ClientID;"
'        End If
'   Else
        szSQL = "SELECT ClientID, ClientName " & _
           "FROM Client order by ClientID;"
'   End If
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
   
   If Not rstRec.EOF Then
      flxClientList.Rows = 2
      flxClientList.RowHeight(0) = 0
      rstRec.MoveFirst
      flxClientList.ColAlignment(1) = vbRightJustify
      flxClientList.TextMatrix(0, 1) = "Client ID"
      flxClientList.TextMatrix(0, 2) = "Client Name"
      flxClientList.TextMatrix(1, 1) = "ALL"
      flxClientList.TextMatrix(1, 2) = "ALL Clients"
'      txtACBalanceByCl.text = Format(GetSupplierBalance(txtSupplierID.text), "0.00")
      rRow = 2
      flxClientList.AddItem ""
      While Not rstRec.EOF
         flxClientList.TextMatrix(rRow, 1) = rstRec!ClientID
         flxClientList.TextMatrix(rRow, 2) = rstRec!ClientName
'         flxClientList.TextMatrix(rRow, 3) = GetSupplierBalanceByClient(rstRec!clientID)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxClientList.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close

   Set rstRec = Nothing
   Set Conn2 = Nothing
   
End Sub
Private Sub LoadSupplierTypes()

'you just change label position then searchbox and grid column will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   'Conn2.Open getConnectionString
   
   
   'SupplierAccountBalanceByClient2 Conn2, txtSupplierFilter.text  'load client wise balance with single supplier
   
'   szSQL = "SELECT ClientID, ClientName " & _
'           "FROM Client order by ClientID;"
'   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly
'
'   If Not rstRec.EOF Then
     
     
      flxClientList.Rows = 2
      flxClientList.RowHeight(0) = 0

        '      rstRec.MoveFirst
        flxClientList.ColAlignment(1) = vbRightJustify
        
        flxClientList.TextMatrix(0, 1) = "Client ID"
        flxClientList.TextMatrix(0, 2) = "Client Name"
        flxClientList.TextMatrix(1, 1) = "Client"
        flxClientList.TextMatrix(1, 2) = "Clients"
        
        '      txtACBalanceByCl.text = Format(GetSupplierBalance(txtSupplierID.text), "0.00")
        rRow = 2
        flxClientList.AddItem ""
        flxClientList.TextMatrix(2, 1) = "Supplier"
        flxClientList.TextMatrix(2, 2) = "Supplier"
        flxClientList.AddItem ""
        flxClientList.TextMatrix(3, 1) = "Agent"
        flxClientList.TextMatrix(3, 2) = "Managing Agent"
        
        flxClientList.AddItem ""
        flxClientList.TextMatrix(4, 1) = "LLORD"
        flxClientList.TextMatrix(4, 2) = "Landlord"
        
        flxClientList.AddItem ""
        flxClientList.TextMatrix(5, 1) = "Lessee"
        flxClientList.TextMatrix(5, 2) = "Lessee"
      
      
'      While Not rstRec.EOF
'         flxClientList.TextMatrix(rRow, 1) = rstRec!clientID
''          flxClientList.Cols = 3
'         flxClientList.TextMatrix(rRow, 2) = rstRec!ClientName
''         flxClientList.TextMatrix(rRow, 3) = GetSupplierBalanceByClient(rstRec!clientID)
'         rstRec.MoveNext
'         If Not rstRec.EOF Then flxClientList.AddItem ""
'         rRow = rRow + 1
'      Wend
''   End If
'
'   rstRec.Close
'   Conn2.Close
'
'   Set rstRec = Nothing
   Set Conn2 = Nothing
   
End Sub

Private Sub cmdClientAddAtch_Click(Index As Integer)
    Dim adoConn As New ADODB.Connection
    Dim rsComment As New ADODB.Recordset
    If Index = 0 Then
       If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub

       If (NEW_TYPE = "Landlord") Then
          AddNewAttachmentInCombo cmbFiles, "Landlord", txtClientID.text
       Else
          AddNewAttachmentInCombo cmbFiles, "Client", txtClientID.text
       End If
       ShowMsgInTaskBar "File has been saved successfully."
     End If
     If Index = 1 Then '2nd button on  1st
            If cmbFiles.text = "" Then Exit Sub
               If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
                  MsgBox "File has been moved from original location.", vbExclamation
               MousePointer = vbDefault
     End If
     If Index = 2 Then '3rd button on  1st
             If cmbFiles.text = "" Then Exit Sub
            If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
            If (NEW_TYPE = "Landlord") Then
               DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtClientID.text, "Landlord"
            Else
               DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtClientID.text, "Client"
            End If
            MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
     End If
     If Index = 3 Then '1st  button on  2nd row
            txtComments1.Locked = False
            cmdClientAddAtch(3).Enabled = False
            cmdClientAddAtch(4).Enabled = True
            cmdClientAddAtch(5).Enabled = True
            FocusControl txtComments1
     End If
     If Index = 4 Then '2nd button on  2nd row
        If SaveComments("Client", "Comments1", txtComments1.text, "ClientID", txtClientID.text) Then
           ShowMsgInTaskBar "The comments have been saved successfully."
        End If
        txtComments1.Locked = True
        cmdClientAddAtch(3).Enabled = True
        cmdClientAddAtch(4).Enabled = False
        cmdClientAddAtch(5).Enabled = False
     End If
     If Index = 5 Then '3rd button on  2nd row
             If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
                txtComments1.Locked = True
                cmdClientAddAtch(3).Enabled = True
                cmdClientAddAtch(4).Enabled = False
                cmdClientAddAtch(5).Enabled = False
                adoConn.Open getConnectionString
                rsComment.Open "Select Comments1 from Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
                If Not rsComment.EOF Then
                    txtComments1.text = IIf(IsNull(rsComment("Comments1").Value), "", rsComment("Comments1").Value)
                End If
                rsComment.Close
                Set rsComment = Nothing
                adoConn.Close
                Set adoConn = Nothing
                FocusControl txtComments1
     End If
     If Index = 6 Then '1st button on  3rd row
            txtComments2(0).Locked = False
            cmdClientAddAtch(6).Enabled = False
            cmdClientAddAtch(7).Enabled = True
            cmdClientAddAtch(8).Enabled = True
            FocusControl txtComments2(0)
     End If
     If Index = 7 Then '2nd button on  3rd row
           If SaveComments("Client", "Comments2", txtComments2(0).text, "ClientID", txtClientID.text) Then
                ShowMsgInTaskBar "The comments2 have been saved successfully."
            End If
            txtComments2(0).Locked = True
            cmdClientAddAtch(6).Enabled = True
            cmdClientAddAtch(7).Enabled = False
            cmdClientAddAtch(8).Enabled = False
     End If
     If Index = 8 Then 'Cancel button 3rd row
          If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
            txtComments2(0).Locked = True
            cmdClientAddAtch(6).Enabled = True
            cmdClientAddAtch(7).Enabled = False
            cmdClientAddAtch(8).Enabled = False
            adoConn.Open getConnectionString
            rsComment.Open "Select Comments2 from Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
            If Not rsComment.EOF Then
                txtComments2(0).text = IIf(IsNull(rsComment("Comments2").Value), "", rsComment("Comments2").Value)
            End If
            rsComment.Close
            Set rsComment = Nothing
            adoConn.Close
            Set adoConn = Nothing
            FocusControl txtComments2(0)
     End If
     If Index = 9 Then 'This is lessee statement edit click '4th row Edit button
            txtComments2(1).Locked = False
            cmdClientAddAtch(9).Enabled = False
            cmdClientAddAtch(10).Enabled = True
            cmdClientAddAtch(11).Enabled = True
            cmdBrowsFile(2).Enabled = True
            FocusControl txtComments2(1)
     End If
     If Index = 10 Then  '4th row save button
            adoConn.Open getConnectionString
            adoConn.Execute "Update Client set LesseeTemplate='" & txtComments2(1).text & "' where ClientID='" & txtClientID.text & "'"
            adoConn.Close
            MsgBox "The template have been saved successfully."
           cmdClientAddAtch(10).Enabled = False
            cmdClientAddAtch(9).Enabled = True
            cmdBrowsFile(2).Enabled = False
     End If
     If Index = 11 Then 'This is cancel lesse statement button clilck '4th row cancel button
           If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
              txtComments2(1).Locked = True
              cmdClientAddAtch(9).Enabled = True
             cmdClientAddAtch(10).Enabled = False
              cmdClientAddAtch(11).Enabled = False
              cmdBrowsFile(2).Enabled = True
              adoConn.Open getConnectionString
              rsComment.Open "Select LesseeTemplate from  Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
              If Not rsComment.EOF Then
                  txtComments2(1).text = IIf(IsNull(rsComment("LesseeTemplate").Value), "", rsComment("LesseeTemplate").Value)
              End If
              rsComment.Close
              Set rsComment = Nothing
              adoConn.Close
              Set adoConn = Nothing
              FocusControl txtComments2(1)
     End If
     If Index = 12 Then 'This is lessee statement edit click '5th row Edit button
            txtComments2(2).Locked = False
            cmdClientAddAtch(12).Enabled = False
            cmdClientAddAtch(13).Enabled = True
            cmdClientAddAtch(14).Enabled = True
            cmdBrowsFile(0).Enabled = True
            FocusControl txtComments2(2)
     End If
     If Index = 13 Then  '5th row save button
            adoConn.Open getConnectionString
            adoConn.Execute "Update Client set LesseeAccTemplate='" & txtComments2(2).text & "' where ClientID='" & txtClientID.text & "'"
            adoConn.Close
            MsgBox "The template have been saved successfully."
            cmdClientAddAtch(12).Enabled = False
            cmdClientAddAtch(13).Enabled = True
            cmdBrowsFile(0).Enabled = False
     End If
     If Index = 14 Then 'This is cancel lesse statement button clilck '4th row cancel button
           If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
              txtComments2(2).Locked = True
              cmdClientAddAtch(12).Enabled = True
              cmdClientAddAtch(13).Enabled = False
              cmdClientAddAtch(14).Enabled = False
              cmdBrowsFile(0).Enabled = True
              adoConn.Open getConnectionString
              rsComment.Open "Select LesseeAccTemplate from  Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
              If Not rsComment.EOF Then
                  txtComments2(1).text = IIf(IsNull(rsComment("LesseeAccTemplate").Value), "", rsComment("LesseeAccTemplate").Value)
              End If
              rsComment.Close
              Set rsComment = Nothing
              adoConn.Close
              Set adoConn = Nothing
              FocusControl txtComments2(2)
     End If
     If Index = 15 Then 'This is lessee statement edit click '5th row Edit button
            txtComments2(3).Locked = False
            cmdClientAddAtch(15).Enabled = False
            cmdClientAddAtch(16).Enabled = True
            cmdClientAddAtch(17).Enabled = True
            cmdBrowsFile(1).Enabled = True
            FocusControl txtComments2(3)
     End If
      If Index = 16 Then  '5th row save button
            adoConn.Open getConnectionString
            adoConn.Execute "Update Client set CSPreviewTemplate='" & txtComments2(3).text & "' where ClientID='" & txtClientID.text & "'"
            adoConn.Close
            MsgBox "The template have been saved successfully."
            cmdClientAddAtch(15).Enabled = False
            cmdClientAddAtch(16).Enabled = True
            cmdBrowsFile(1).Enabled = False
     End If
     If Index = 17 Then 'This is cancel lesse statement button clilck '4th row cancel button
           If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
              txtComments2(3).Locked = True
              cmdClientAddAtch(15).Enabled = True
              cmdClientAddAtch(16).Enabled = False
              cmdClientAddAtch(17).Enabled = False
              cmdBrowsFile(1).Enabled = True
              adoConn.Open getConnectionString
              rsComment.Open "Select CSPreviewTemplate from  Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
              If Not rsComment.EOF Then
                  txtComments2(3).text = IIf(IsNull(rsComment("CSPreviewTemplate").Value), "", rsComment("CSPreviewTemplate").Value)
              End If
              rsComment.Close
              Set rsComment = Nothing
              adoConn.Close
              Set adoConn = Nothing
              FocusControl txtComments2(3)
     End If
     If Index = 18 Then 'This is lessee statement edit click '5th row Edit button
            txtComments2(4).Locked = False
            cmdClientAddAtch(18).Enabled = False
            cmdClientAddAtch(19).Enabled = True
            cmdClientAddAtch(20).Enabled = True
            cmdBrowsFile(3).Enabled = True
            FocusControl txtComments2(4)
     End If
     If Index = 19 Then  '5th row save button
            adoConn.Open getConnectionString
            adoConn.Execute "Update Client set CSTemplate='" & txtComments2(4).text & "' where ClientID='" & txtClientID.text & "'"
            adoConn.Close
            MsgBox "The template have been saved successfully."
            cmdClientAddAtch(18).Enabled = False
            cmdClientAddAtch(19).Enabled = True
            cmdBrowsFile(3).Enabled = False
     End If
     If Index = 20 Then 'This is cancel lesse statement button clilck '4th row cancel button
           If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
              txtComments2(4).Locked = True
              cmdClientAddAtch(18).Enabled = True
              cmdClientAddAtch(19).Enabled = False
              cmdClientAddAtch(20).Enabled = False
              cmdBrowsFile(3).Enabled = True
              adoConn.Open getConnectionString
              rsComment.Open "Select CSTemplate from  Client where ClientID='" & txtClientID.text & "'", adoConn, adOpenStatic, adLockReadOnly
              If Not rsComment.EOF Then
                  txtComments2(4).text = IIf(IsNull(rsComment("CSTemplate").Value), "", rsComment("CSTemplate").Value)
              End If
              rsComment.Close
              Set rsComment = Nothing
              adoConn.Close
              Set adoConn = Nothing
              FocusControl txtComments2(4)
     End If
End Sub

Private Sub cmdClientFilter_Click()
    
    strCommandSource = "SupplierTypesFilter"
    
    LoadSupplierTypes '
    
    
    Frame5.Top = cmdClientFilter.Top + 700
    Frame5.Left = cmdClientFilter.Left - 500
    Frame5.Visible = True
    Frame5.ZOrder 0
    FocusControl txtSearchClientID
End Sub



Private Sub cmdconsolidatedAccountName_Click()
    Frame5.Top = 3710
    Frame5.Left = 5715
    strCommandSource = "ConsolidatedBank"
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadConsolidatedBank(adoConn)
    Frame5.Visible = True
    FocusControl txtSearchClientID
    adoConn.Close
    Set adoConn = Nothing
End Sub



Private Sub cmdFix_Click()
    Dim adoConn As New ADODB.Connection
    Dim rsCheck As New ADODB.Recordset
    Dim szSQL As String
    adoConn.Open getConnectionString
    adoConn.Execute "Update tlbAgreement T,ClientProAgr P set T.ReportClientID=P.ClientID,T.ReportPropertyID=P.PropertyID where T.CPA_ID=P.CPA_ID"
    
    szSQL = "SELECT CL_ID,max(Cdate(duedate)) as DDate from tblPurInv where isManagementFee=true group by CL_ID"
    rsCheck.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    While Not rsCheck.EOF
         szSQL = "Update tlbAgreement A,tblPurInv P Set A.LastChargeDate='" & rsCheck("DDate").Value & "'  where isManagementFee=true AND A.ReportClientID='" & rsCheck("CL_ID").Value & "'"
         adoConn.Execute szSQL
         rsCheck.MoveNext
    Wend
    rsCheck.Close
   
    
    
    adoConn.Close
    MsgBox "Last Charge date has been updated"
End Sub

Private Sub cmdPaymentTypeNew_Click(Index As Integer)
    If Index = 0 Or Index = 2 Then
        frmSecondaryCode.PRIMARY_CODE_SHOW = "RAT"
        Load frmSecondaryCode
        frmSecondaryCode.Show 1
    End If
    If Index = 1 Then
        Frame5.Top = 1710
        Frame5.Left = 8715
        strCommandSource = "PaymentType1"
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Call LoadflxPaymentMethod(adoConn)
        Frame5.Visible = True
        FocusControl txtSearchClientID
        adoConn.Close
        Set adoConn = Nothing
    End If
End Sub

Private Sub cmdPrintAgreement_Click()
    frmPrintAgreement.Show
End Sub

Private Sub cmdRCCSave2_Click(Index As Integer)
   If Index = 0 Then
        If SaveComments("Client", "Comments2", txtComments2(0).text, "ClientID", txtClientID.text) Then
       ShowMsgInTaskBar "The comments2 have been saved successfully."
        End If
        txtComments2(0).Locked = True
        cmdClientAddAtch(6).Enabled = True
        cmdClientAddAtch(7).Enabled = False
        cmdClientAddAtch(8).Enabled = False
   End If
   If Index = 1 Then 'this is save lessee statement click button
       Dim adoConn As New ADODB.Connection
       adoConn.Open getConnectionString
       adoConn.Execute "Update Client set LesseeTemplate='" & txtComments2(1).text & "' where ClientID='" & txtClientID.text & "'"
       adoConn.Close
       MsgBox "The template have been saved successfully."
      cmdClientAddAtch(10).Enabled = False
       cmdClientAddAtch(9).Enabled = True
       cmdBrowsFile(2).Enabled = False
   End If
End Sub

Private Sub cmdSearch_Click()
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        fraSearch.Left = 9404
        fraSearch.Top = 4140
        
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        If cmdSearch.Caption = "Clear Sea&rch" Then
             txtSearchNo.text = ""
             txtSearchRef.text = ""
             fmeLoading.Visible = False
             cmdSearch.Caption = "Sea&rch"
             fraSearch.Visible = False
             Call LoadFlxACHistory(adoConn, "")
        Else
            If fraSearch.Visible = False Then
                fraSearch.Visible = True
                txtSearchNo.SetFocus
            Else
                fraSearch.Visible = False
            End If
        End If
        adoConn.Close
        Set adoConn = Nothing
End Sub

Private Sub cmdSupplierFilter_Click()
'    strCommandSource = "SupplierFilter"
'    Call LoadSupplier
'    Frame5.Top = cmdClientFilter.Top + 700
'    Frame5.Left = cmdClientFilter.Left - 500
'    Frame5.Visible = True
'    Frame5.ZOrder 0
'    FocusControl txtSearchClientID
'    Exit Sub
    If txtFilterClient.text = "Clients" Then
        strCommandSource = "ClientFilter"
        LoadClients 'load client and also the array of balance
    End If
    If txtFilterClient.text = "Supplier" Then
        strCommandSource = "SupplierFilter"
        Call LoadSupplier("Supplier")
    End If
    If txtFilterClient.text = "Agent" Then
        strCommandSource = "AgentFilter"
        Call LoadSupplier("Agent")
    End If
    If txtFilterClient.text = "Landlord" Then
        strCommandSource = "LandlordFilter"
        Call LoadSupplier("LLORD")
    End If
    If txtFilterClient.text = "Lessee" Then
        strCommandSource = "LesseeFilter"
        Call LoadLessee
    End If
    
'    UpdateBalanceCL 'this function fills the grid with balalnce from array
'    Dim Conn2 As New ADODB.Connection
'    Conn2.Open getConnectionString
'    flxClientList.TextMatrix(1, 3) = Format(SupplierAccountBalanceALLClient(Conn2, txtClientID.text), "0.00")
'    txtACBalanceByCl.text = flxClientList.TextMatrix(1, 3)
'    Conn2.Close
    Frame5.Top = cmdClientFilter.Top + 700
    Frame5.Left = cmdClientFilter.Left - 500
    Frame5.Visible = True
    Frame5.ZOrder 0
    FocusControl txtSearchClientID
End Sub

Private Sub flxBankAccountFund_Click()
    If flxBankAccountFund.TextMatrix(flxBankAccountFund.row, 0) = "X" Then
        flxBankAccountFund.TextMatrix(flxBankAccountFund.row, 0) = ""
        flxBankAccountFund.CellBackColor = vbWhite
        Exit Sub
    Else
        flxBankAccountFund.TextMatrix(flxBankAccountFund.row, 0) = "X"
    End If
End Sub

Private Sub flxManagementFee_Click()
    If cmdAgmntSave.Enabled = False Then
        cmdAgmntEdit.Enabled = True
        cmdDeleteMgtFee.Enabled = True
    End If
End Sub

Private Sub flxPayable_Click()
    cmdPayEdit.Enabled = True
End Sub

Private Sub Form_Resize()
'            SaveSizes
'        ResizeControls
End Sub
'in tlbAgreement table we save manangement fee each line for One property(Each property)
Private Sub cboBank_ID_LostFocus()
      Dim conBank As New ADODB.Connection
       Dim rstBank As New ADODB.Recordset
       Dim szSQL As String
    If Trim(cboBank_ID.text) = "" Then
        txtBANK_NAME.text = ""
        txtBANK_ADDRESS1.text = ""
        txtBANK_ADDRESS2.text = ""
        txtBANK_ADDRESS3.text = ""
        txtBANK_POST_CODE.text = ""
        txtNCCODE.text = ""
        txtNominal.text = ""
        txtPaymentMethod.text = ""
        txtBank_AC_Name.text = ""
        txtBANK_SC.text = ""
        txtBANK_AC_NUM.text = ""
   Else
'       conBank.Open getConnectionString
'       szSQL = "SELECT * " & _
'               "FROM tlbClientBanks, tlbBank " & _
'               "WHERE CLIENT_ID = '" & txtClientID.text & "' And " & _
'                   "tlbBank.BANK_ID = '" & cboBank_ID.text & "' " & _
'               "ORDER BY Bank_AC_Name;"
'
'       rstBank.Open szSQL, conBank, adOpenStatic, adLockReadOnly
'       If rstBank.EOF Then
'            MsgBox "Please select the corrent bank ID", vbInformation, "Wrong selection"
'            rstBank.Close
'            conBank.Close
'            FocusControl cboBank_ID
'            Exit Sub
'       End If
'       rstBank.Close
'       conBank.Close
       
   End If
End Sub

'Private Sub cboDmdPropertyList_LostFocus()
'        Dim rstBank As New ADODB.Recordset
'        Dim conBank As New ADODB.Connection
'        Dim szSQL As String
'        'If cboDmdPropertyList.text <> "" Then
'            conBank.Open getConnectionString
'            szSQL = "SELECT * " & _
'               "FROM Property " & _
'               "WHERE " & _
'                   "PropertyName = '" & cboDmdPropertyList.text & "' " & _
'               "ORDER BY PropertyName;"
'
'            rstBank.Open szSQL, conBank, adOpenStatic, adLockReadOnly
'            If rstBank.EOF Then
'                 MsgBox "Please select the correct property", vbInformation, "Wrong selection"
'                 rstBank.Close
'                 conBank.Close
'                 FocusControl cboDmdPropertyList
'                 cboDmdPropertyList.text = ""
'                 Exit Sub
'            End If
'            rstBank.Close
'            conBank.Close
'         'End If
'End Sub

'Private Sub cboconsolidatedAccountName_Click()
'    cboconsolidatedAccountName_LostFocus
'End Sub

'Private Sub cboconsolidatedAccountName_LostFocus()
'       Dim adoConn As New ADODB.Connection
'       Dim szSQL As String
'       Dim adoRst As New ADODB.Recordset
'       If Trim(cboconsolidatedAccountName.text) <> "" Then
'        adoConn.Open getConnectionString
'        szSQL = "SELECT * " & _
'                "FROM ConsolidatedBankList " & _
'                "where BankName='" & cboconsolidatedAccountName.text & "';"
'        adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'        If adoRst.EOF Then
'
'            MsgBox "Please select a valid Conslidated Bank account", vbInformation, "Select a Conslidated Bank account"
'            FocusControl cboconsolidatedAccountName
'            txtBANK_SC.text = ""
'            txtBANK_AC_NUM.text = ""
'        Else
'            txtBANK_SC.text = adoRst("SortCode").Value
'            txtBANK_AC_NUM.text = adoRst("BankACNumber").Value
'        End If
'        adoRst.Close
'        adoConn.Close
'        Set adoConn = Nothing
'    End If
'End Sub

Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtREVIEW_DATE
    End If
End Sub

'Private Sub cboProperty_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl txtREVIEW_DATE
'    End If
'End Sub

Private Sub cboProperty_LostFocus()
'    Dim rstBank As New ADODB.Recordset
'        Dim conBank As New ADODB.Connection
'        Dim szSQL As String
'
'            conBank.Open getConnectionString
'            szSQL = "SELECT * " & _
'               "FROM Property " & _
'               "WHERE " & _
'                   "PropertyName = '" & cboProperty.text & "' " & _
'               "ORDER BY PropertyName;"
'
'            rstBank.Open szSQL, conBank, adOpenStatic, adLockReadOnly
'            If rstBank.EOF Then
'                 MsgBox "Please select the correct property", vbInformation, "Wrong selection"
'                 rstBank.Close
'                 conBank.Close
'                 FocusControl cboProperty
'                 cboProperty.text = ""
'                 Exit Sub
'            End If
'            rstBank.Close
'            conBank.Close
End Sub

Private Sub chkConsolidated_Click()
'     If bBankNewEdit = False And cmdEdit.Enabled = False Then 'this means this is in edit mode
'        Dim adoconn As New ADODB.Connection
'        Dim rsIsReadonly As New ADODB.Recordset
'        adoconn.Open getConnectionString
'        Dim bPreviousState As Integer
'        bPreviousState = chkConsolidated.Value
'        Dim bResult As Boolean
'        rsIsReadonly.Open "Select * from tlbClientBanks where Client_ID='" & txtClientID.text & "' and NominalCode='" & txtNCCODE.text & "'", adoconn, adOpenStatic, adLockReadOnly
'        If Not rsIsReadonly.EOF Then
'                bResult = IIf(IsNull(rsIsReadonly("conBankReadOnly").Value), 0, rsIsReadonly("conBankReadOnly").Value)
'                If bResult = True Then
'                    MsgBox "This consolidated bank Account has been marked as ready only as there are some transaction already exists to this accounts.", vbOKOnly, "Warning"
'                    'chkConsolidated.Value = bPreviousState
'                End If
'        End If
'        rsIsReadonly.Close
'        Set rsIsReadonly = Nothing
'        adoconn.Close
'        Set adoconn = Nothing
'        If bResult = True Then
'
'            Exit Sub
'        End If
'    End If
    
    If chkConsolidated.Value = 1 Then
'        cboconsolidatedAccountName.Visible = True
        'cboconsolidatedAccountName.Enabled = True
        'cboconsolidatedAccountName.Locked = False
'        txtBank_AC_Name.Visible = False
'        cboconsolidatedAccountName.Enabled = True
    Else
        txtBank_AC_Name.Visible = True
        txtconsolidatedAccountName.text = ""
        txtconsolidatedAccountName.Tag = ""
'        cboconsolidatedAccountName.Visible = False
'        cboconsolidatedAccountName.Enabled = False
    End If
End Sub

Private Sub chkOptedtoTax_Click()
    If chkOptedtoTax.Value = 0 Then
        txtAcBalance(1).text = ""
        txtAcBalance(1).Tag = ""
    End If
End Sub

Private Sub cmdBrowseTemplate_Click()
   Dim ofn As OPENFILENAME
   Dim lHwnd As Long
   Const HKEY_LOCAL_MACHINE As Long = &H80000002
   Dim szOldFile_PathName As String
   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
   Dim fso As Object

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "All Files (*.rpt)" + Chr$(0) + "*.rpt" + Chr$(0)
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir & "\CompanyReports"
   ofn.lpstrTitle = "Select a Report file"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   txtRenSummaryStatement.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 1 Then txtEmailTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 2 Then txtStatementTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
'   If Index = 0 Then
'        FocusControl cmdBrowsFile(1)
'   ElseIf Index = 1 Then
'        FocusControl cmdBrowsFile(2)
'   ElseIf Index = 2 Then
'        FocusControl cmdSaveNew
'   End If
End Sub

Private Sub cmdCancelFeenCharge_Click()
    cmdEditFeeNChargePaydates.Enabled = True
    cmdSaveFEEnCharge.Enabled = False
    cmdCancelFeenCharge.Enabled = False
End Sub

Private Sub cmdCanelAgree_Click()
'    cboProperty.Enabled = False
    cmdAgrTopEdit.Enabled = True
    cmdAgrTopSave.Enabled = False
    cmdCanelAgree.Enabled = False
    FocusControl flxPropertySelection1
    txtAgreementStartDate.Enabled = False
            txtAgreementEndDate.Enabled = False
            txtREVIEW_DATE.Enabled = False
             Call LoadflxAgreement(szPropertySelection1)
End Sub

Private Sub cmdClose2_Click()
    Unload Me
End Sub

Private Sub cmdClose3_Click()
    Unload Me
End Sub

Private Sub cmdCommandArray_Click(Index As Integer)
    tabMain.Enabled = False
    picMain.Enabled = False
    If Index = 0 Then
            strCommandSource = "PAYABLETYPES"
            Call loadpayableTypes
            Frame5.Top = 3915
            Frame5.Left = 900
            Frame5.Visible = True
     ElseIf Index = 1 Then
'            strCommandSource = "DEMANDTYPES"
'            Call LoadDemandTypes
'            Frame5.Top = 3915
'            Frame5.Left = 1400
'            Frame5.Visible = True
            strCommandSource = "PayeeTYPES"
            Call LoadPayeeTypes
            Frame5.Top = 3915
            Frame5.Left = 1400
            Frame5.Visible = True

     ElseIf Index = 2 Then
            strCommandSource = "FUND"
            Call LoadFunds
            Frame5.Top = 3915
            Frame5.Left = 2400
            Frame5.Visible = True
     ElseIf Index = 3 Then
            If txtPayeeType.text = "" Then
                MsgBox "Please select a payee type"
                tabMain.Enabled = True
                picMain.Enabled = True
            Else
                strCommandSource = "ClientLandlord"
                Call loadClientLandlord
                Frame5.Top = 3915
                Frame5.Left = 3500
                Frame5.Visible = True
             End If
     ElseIf Index = 6 Then
            strCommandSource = "ManagingAgent"
            Call LoadManagingAgent
            Frame5.Top = 3915
            Frame5.Left = 3500
            Frame5.Visible = True
    
     ElseIf Index = 4 Then
            strCommandSource = "Frequencies"
            Call loadFrequencies
            Frame5.Top = 3915
            Frame5.Left = 4500
            Frame5.Visible = True
     ElseIf Index = 5 Then
            strCommandSource = "PayableBasis"
            Call loadPayableBasis
            Frame5.Top = 3915
            Frame5.Left = 5500
            Frame5.Visible = True
     ElseIf Index = 9 Then
            strCommandSource = "chargeTypes"
            Call loadchargeTypes
            Frame5.Top = 3915
            Frame5.Left = 900
            Frame5.Visible = True
     ElseIf Index = 8 Then
            strCommandSource = "DEMANDTYPESMangFee"
            Call LoadDemandTypes
            Frame5.Top = 3915
            Frame5.Left = 1400
            Frame5.Visible = True
     ElseIf Index = 7 Then
            strCommandSource = "FUNDMangFee"
            Call LoadFunds
            Frame5.Top = 3915
            Frame5.Left = 2400
            Frame5.Visible = True
     ElseIf Index = 6 Then
            strCommandSource = "ClientLandlordMangFee"
            Call loadClientLandlord
            Frame5.Top = 3915
            Frame5.Left = 3500
            Frame5.Visible = True
     ElseIf Index = 10 Then
            strCommandSource = "ChargingMethod"
            Call LoadChargingMethod
            Frame5.Top = 3915
            Frame5.Left = 3500
            Frame5.Visible = True
     ElseIf Index = 11 Then
            strCommandSource = "ChargingBasis"
            Call LoadChargeBasis
            Frame5.Top = 3915
            Frame5.Left = 3500
            Frame5.Visible = True
     ElseIf Index = 12 Then
            strCommandSource = "FrequenciesMngtFee"
            Call loadFrequencies
            Frame5.Top = 3915
            Frame5.Left = 4500
            Frame5.Visible = True
    ElseIf Index = 13 Then
            strCommandSource = "BankFund"
            Call LoadFunds
            Frame5.Top = 3915
            Frame5.Left = 2400
            Frame5.Visible = True
    End If
    FocusControl txtSearchClientID
End Sub
Private Sub LoadChargingMethod()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
'This is Ideal gridview Popup which is written by anol 2020-18-12
'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   'lblClientID(1).Visible = False
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = 0 'lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "CODE"
   lblClientID(1).Caption = "Charging Method"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
   adoConn.Open getConnectionString
   szSQL = "SELECT CODE, VALUE FROM SECONDARYCODE WHERE PRIMARYCODE = 'CRGBS' Order by VALUE;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   rRow = 1
              
   While Not rstRec.EOF
        flxClientList.row = 1
        flxClientList.RowSel = 1
        flxClientList.ColSel = 1
        flxClientList.TextMatrix(rRow, 0) = ""
        flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("CODE").Value
        flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("VALUE").Value
       
        flxClientList.RowHeight(rRow) = 280
        rstRec.MoveNext
        If Not rstRec.EOF Then flxClientList.AddItem ""
        rRow = rRow + 1
    Wend
    rstRec.Close
    Set rstRec = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub loadchargeTypes()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Charge Type"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
   
   

   adoConn.Open getConnectionString
   szSQL = "SELECT ID, FeeType FROM ChargeTypes where PropertyID='" & szPropertySelection1 & "';"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FeeType").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdDeleteMgtFee_Click()
    Dim adoConn As New ADODB.Connection
    If flxManagementFee.row = 0 Then
        MsgBox "Please select a row to Delete", vbInformation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete this line?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
        adoConn.Open getConnectionString
        adoConn.Execute "Delete from tlbagreement where AGREEMENT_ID=" & flxManagementFee.TextMatrix(flxManagementFee.row, 1) & ""
        Dim strLastChargeDate As String
        strLastChargeDate = findLastChargeDate(szPropertySelection1, flxManagementFee.TextMatrix(flxManagementFee.row, 7), adoConn)
        Call loadflxManagementFee(adoConn)
        adoConn.Close
        Set adoConn = Nothing
        AgreementButtonMode DefaultMode
        AgreementClearMode ClearOnlyTextBoxes
End Sub
Private Function findLastChargeDate(strPropertyID As String, fundID As Long, adoConn As ADODB.Connection)
   ' Dim adoConn As New ADODB.Connection
    Dim rsChargedate As New ADODB.Recordset
    'adoConn.Open getConnectionString
    adoConn.Execute "Update tlbAgreement T,ClientProAgr P set T.ReportClientID=P.ClientID,T.ReportPropertyID=P.PropertyID where T.CPA_ID=P.CPA_ID"
    rsChargedate.Open "Select max(S.LastChargeDate) as chrgDate from tlbAgreement S where  S.ReportPropertyID='" & _
    strPropertyID & "' And Fund = " & fundID & "", adoConn, adOpenStatic, adLockReadOnly
    
    
'    rsChargedate.Open "Select max(R.ChargeDate) as chrgDate from tlbReceiptSplit S,tlbreceipt R,Units U where U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'    strPropertyID & "'", adoconn, adOpenStatic, adLockReadOnly
    If Not rsChargedate.EOF Then
        findLastChargeDate = IIf(IsNull(rsChargedate("chrgDate").Value) = True, "", rsChargedate("chrgDate").Value)
    End If
    rsChargedate.Close
    Set rsChargedate = Nothing
    'adoConn.Close
    'Set adoConn = Nothing

End Function
Private Sub cmdDeleteRentPayable_Click()
    Dim adoConn As New ADODB.Connection
    If flxPayable.row = 0 Then
        MsgBox "Please select a row to Delete", vbInformation, "Warning"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete this line?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
    adoConn.Open getConnectionString
    adoConn.Execute "Delete from tlbPayable where PAYABLE_ID=" & flxPayable.TextMatrix(flxPayable.row, 1) & ""
    Call loadflxPayable(adoConn)
    txtClientLandlord.text = ""
    txtClientLandlord.Tag = ""
    txtPayableBasis.text = ""
    txtPayableBasis.Tag = ""
    txtPercentage.text = ""
    txtPercentage.Tag = ""
    'new code by anol 20210823
        txtClientLandlord.text = ""
        txtClientLandlord.Tag = ""
        txtPayeeType.text = ""
        txtPayeeType.Tag = ""
        txtPayableBasis.text = ""
        txtPayableBasis.Tag = ""
        txtPercentage.text = ""
        txtPercentage.Tag = ""
    adoConn.Close
    Set adoConn = Nothing
    PayableButtonMode DefaultMode
    PayableClearMode ClearOnlyTextBoxes
End Sub

Private Sub cmdEdit_Click()
    If cmdSaveClient.Enabled = True Then
        MsgBox "Please save the header section first before proceeding with update details", vbInformation, "Warning"
        FocusControl cmdSaveClient
        Exit Sub
   End If
   PopulateBank
   If iTotalBankAC = 0 Then Exit Sub      'there are no bank details has been inputed yet
   cmdClient.Enabled = False
   cmdEdit.Enabled = False
   bBankNewEdit = False
   bOverdraftWarning = True
   EnableDisableAcText False
   LockAcText False
   'txtBANK_SC.Locked = True

   CommandButtonEnabled False
   flxOtherBankDetails.Enabled = True
   cmdAddEditBankCode.Enabled = True
   cboBank_ID.Locked = False
   cmdSetDefaultAC.Enabled = True
   cmdBACS.Enabled = True
   cmdDeleteBank.Enabled = False
   Frame14.Enabled = False
   chkConsolidated.Enabled = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 22) = "1", False, True)
   
End Sub

Private Sub cmdEditFeeNChargePaydates_Click()
    fraPaymentDate(14).Enabled = True
    fraPaymentDate(13).Enabled = True
    fraPaymentDate(12).Enabled = True
    fraPaymentDate(11).Enabled = True
    fraPaymentDate(10).Enabled = True
    fraPaymentDate(9).Enabled = True
    cmdAutoSetup(1).Enabled = True
    cmdSaveFEEnCharge.Enabled = True
    txtNoOfDaysToSendMFB4Due.Locked = False
    cmdEditFeeNChargePaydates.Enabled = False
    FocusControl txtNoOfDaysToSendMFB4Due
End Sub

Private Sub cmdNC_Click()
    Frame5.Top = 1710
    Frame5.Left = 5715
    strCommandSource = "NominalCode"
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadCmbNC(adoConn)
    Frame5.Visible = True
    FocusControl txtSearchClientID
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdPaymentType_Click()
    Frame5.Top = 1710
    Frame5.Left = 8715
    strCommandSource = "PaymentType"
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadflxPaymentMethod(adoConn)
    Frame5.Visible = True
    FocusControl txtSearchClientID
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cmdPaymentTypeCancel_Click()
    cmdPaymentType.Enabled = False
    cmdPaymentTypeNew(0).Enabled = False
    cmdBrowseTemplate.Enabled = False
    txtPaymentTerms.Locked = True
    cmdSavePaymentDetails.Enabled = False
    cmdPaymentTypeUpdate.Enabled = True
    cmdPaymentTypeCancel.Enabled = False
End Sub

Private Sub cmdPaymentTypeUpdate_Click()
    cmdPaymentType.Enabled = True
    cmdPaymentTypeNew(0).Enabled = True
    cmdBrowseTemplate.Enabled = True
    txtPaymentTerms.Locked = False
    cmdSavePaymentDetails.Enabled = True
    cmdPaymentTypeUpdate.Enabled = False
    txtPaymentType.Locked = False
    txtClientHomeTel(11).Locked = False
    txtClientHomeTel(12).Locked = False
    txtClientHomeTel(13).Locked = False
    txtClientHomeTel(14).Locked = False
    chkUsePayableTemplate.Enabled = True
    cmdPaymentTypeCancel.Enabled = True
End Sub





'Private Sub cmdClientAddAtch(3)_Click()
'  txtComments1.Locked = False
'   cmdClientAddAtch(3).Enabled = False
'   cmdClientAddAtch(4).Enabled = True
'   cmdClientAddAtch(5).Enabled = True
'   FocusControl txtComments1
'End Sub



'Private Sub cmdClientAddAtch(4)_Click()
'
'   If SaveComments("Client", "Comments1", txtComments1.text, "ClientID", txtClientID.text) Then
'      ShowMsgInTaskBar "The comments have been saved successfully."
'   End If
'   txtComments1.Locked = True
'   cmdClientAddAtch(3).Enabled = True
'   cmdClientAddAtch(4).Enabled = False
'   cmdClientAddAtch(5).Enabled = False
'End Sub

'Private Sub cmdRCCSave2_Click()
'  If SaveComments("Client", "Comments2", txtComments2.text, "ClientID", txtClientID.text) Then
'      ShowMsgInTaskBar "The comments2 have been saved successfully."
'   End If
'   txtComments2.Locked = True
'   cmdClientAddAtch(6).Enabled = True
'   cmdRCCSave2.Enabled = False
'   cmdClientAddAtch(8).Enabled = False
'End Sub

Private Sub cmdRemittanceTemplate_Click()
       Dim ofn As OPENFILENAME
   Dim lHwnd As Long
   Const HKEY_LOCAL_MACHINE As Long = &H80000002
   Dim szOldFile_PathName As String
   Dim szNewFile_Path As String, szNewFile_Name As String, szNewFile_PathName As String
   Dim fso As Object

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "All Files (*.rpt)" + Chr$(0) + "*.rpt" + Chr$(0)
   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir & "\CompanyReports"
   ofn.lpstrTitle = "Select a Report file"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   txtRemittanceTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
End Sub

Public Function ValidDate(text As String) As Boolean
   ValidDate = True
   If IsDate(text) = False Then
       MsgBox "Invalid Date Selected.", vbOKOnly + vbCritical, "Invalid Date"
       ValidDate = False
   End If
End Function
Private Sub cmdSaveFEEnCharge_Click()
    Dim tempdate As String
   If Trim(txtNoOfDaysToSendMFB4Due.text) = "" Then
        MsgBox "Please enter number of days to send Management Fee", vbInformation, "Warning"
        FocusControl txtNoOfDaysToSendMFB4Due
        Exit Sub
   End If
   'make sure all payment dates are entered.
   If MissingDate(cboDay(31)) = True Then Exit Sub
   If MissingDate(cboDay(32)) = True Then Exit Sub
   If MissingDate(cboDay(33)) = True Then Exit Sub
   If MissingDate(cboDay(34)) = True Then Exit Sub
   If MissingDate(cboDay(35)) = True Then Exit Sub
   If MissingDate(cboDay(36)) = True Then Exit Sub
   If MissingDate(cboDay(37)) = True Then Exit Sub
   If MissingDate(cboDay(38)) = True Then Exit Sub
   If MissingDate(cboDay(39)) = True Then Exit Sub
   If MissingDate(cboDay(40)) = True Then Exit Sub
   If MissingDate(cboDay(41)) = True Then Exit Sub
   If MissingDate(cboDay(42)) = True Then Exit Sub

   If MissingDate(cboQDay(11)) = True Then Exit Sub
   If MissingDate(cboQDay(12)) = True Then Exit Sub
   If MissingDate(cboQDay(13)) = True Then Exit Sub
   If MissingDate(cboQDay(14)) = True Then Exit Sub
   If MissingDate(cboQMth(11)) = True Then Exit Sub
   If MissingDate(cboQMth(12)) = True Then Exit Sub
   If MissingDate(cboQMth(13)) = True Then Exit Sub
   If MissingDate(cboQMth(14)) = True Then Exit Sub
   
   If MissingDate(cboHDay(5)) = True Then Exit Sub
   If MissingDate(cboHDay(6)) = True Then Exit Sub
   If MissingDate(cboHMth(5)) = True Then Exit Sub
   If MissingDate(cboHMth(6)) = True Then Exit Sub
   If MissingDate(cboYDay(2)) = True Then Exit Sub
   If MissingDate(cboYMth(2)) = True Then Exit Sub
 
   
   'validate the dates.
   tempdate = Format("January " & cboDay(31).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(31).SetFocus
      Exit Sub
   End If

   tempdate = Format("February " & cboDay(32).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
       cboDay(32).SetFocus
      Exit Sub
   End If

   tempdate = Format("March " & cboDay(33).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(33).SetFocus
      Exit Sub
   End If

   tempdate = Format("April " & cboDay(34).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(34).SetFocus
      Exit Sub
   End If

   tempdate = Format("May " & cboDay(35).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(35).SetFocus
      Exit Sub
   End If

   tempdate = Format("June " & cboDay(36).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(36).SetFocus
      Exit Sub
   End If

   tempdate = Format("July " & cboDay(37).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(37).SetFocus
      Exit Sub
   End If

   tempdate = Format("August " & cboDay(38).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(38).SetFocus
      Exit Sub
   End If

   tempdate = Format("September " & cboDay(39).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(39).SetFocus
      Exit Sub
   End If

   tempdate = Format("October " & cboDay(40).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(40).SetFocus
      Exit Sub
   End If

   tempdate = Format("November " & cboDay(41).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(41).SetFocus
      Exit Sub
   End If

   tempdate = Format("December " & cboDay(42).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboDay(42).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboQMth(11).text & " " & cboQDay(11).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboQMth(11).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboQMth(12).text & " " & cboQDay(12).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboQMth(12).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboQMth(13).text & " " & cboQDay(13).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboQDay(13).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboQMth(14).text & " " & cboQDay(14).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboQMth(14).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboHDay(5).text & " " & cboHMth(5).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboHDay(5).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboHDay(6).text & " " & cboHMth(6).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboHDay(6).SetFocus
      Exit Sub
   End If

   tempdate = Format(cboYDay(2).text & " " & cboYMth(2).text & ", 20" & Right(Date, 2), "MMMM DD, YYYY")
   If Not ValidDate(tempdate) Then
      cboYDay(2).SetFocus
      Exit Sub
   End If

   Dim i As Integer

   If txtNoOfDaysToSendMFB4Due.text = "" Then
      MsgBox "Please enter the number of days of Fees and charges to send before due.", vbCritical + vbOKOnly, "Fees and charges Notice Period"
      txtNoOfDaysToSendMFB4Due.SetFocus
      Exit Sub
   End If



   Dim conn3 As New ADODB.Connection

'* Save records in the database
   conn3.Open getConnectionString
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT * " & _
            "FROM GlobalData " & _
            "WHERE PropertyID = '" & szPropertySelection2 & "' ", _
                    conn3, adOpenDynamic, adLockOptimistic
   If rst.EOF Then
       MsgBox "Please enter global Data first for property" & szPropertySelection2, vbInformation, "Warning"
   Else
       rst.MoveFirst
   End If

   If szPropertySelection2 <> "" Then
       rst!propertyID = szPropertySelection2
   Else
       MsgBox "Please select a property to continue", vbInformation, "Save Payment Dates"
       rst.Close
       conn3.Close
       Exit Sub
   End If


   'If Rst!MDueDate1 <> cboDay(31).text & " January" Then
        rst!MDueDate1 = cboDay(31).text & " January"
   'End If
   'If Rst!MDueDate2 <> cboDay(32).text & " February" Then
        rst!MDueDate2 = cboDay(32).text & " February"
   'End If
   'If Rst!MDueDate3 <> cboDay(33).text & " March" Then
         rst!MDueDate3 = cboDay(33).text & " March"
   'End If
   'If Rst!MDueDate4 <> cboDay(34).text & " April" Then
        rst!MDueDate4 = cboDay(34).text & " April"
   'End If
   'If Rst!MDueDate5 <> cboDay(35).text & " May" Then
        rst!MDueDate5 = cboDay(35).text & " May"
   'End If
   'If Rst!MDueDate6 <> cboDay(36).text & " June" Then
        rst!MDueDate6 = cboDay(36).text & " June"
'   End If
'   If Rst!MDueDate7 <> cboDay(37).text & " July" Then
        rst!MDueDate7 = cboDay(37).text & " July"
'   End If
'   If Rst!MDueDate8 <> cboDay(38).text & " August" Then
        rst!MDueDate8 = cboDay(38).text & " August"
'   End If
'   If Rst!MDueDate9 <> cboDay(39).text & " September" Then
        rst!MDueDate9 = cboDay(39).text & " September"
'   End If
'   If Rst!MDueDate10 <> cboDay(40).text & " October" Then
        rst!MDueDate10 = cboDay(40).text & " October"
'   End If
'   If Rst!MDueDate11 <> cboDay(41).text & " November" Then
        rst!MDueDate11 = cboDay(41).text & " November"
'   End If
'   If Rst!MDueDate12 <> cboDay(42).text & " December" Then
        rst!MDueDate12 = cboDay(42).text & " December"
'   End If
'   If Rst!QDueDate1 <> cboQDay(11).text & " " & cboQMth(11).text Then
        rst!QDueDate1 = cboQDay(11).text & " " & cboQMth(11).text
'   End If
'   If Rst!QDueDate2 <> cboQDay(12).text & " " & cboQMth(12).text Then
        rst!QDueDate2 = cboQDay(12).text & " " & cboQMth(12).text
'   End If
'   If Rst!QDueDate3 <> cboQDay(13).text & " " & cboQMth(13).text Then
        rst!QDueDate3 = cboQDay(13).text & " " & cboQMth(13).text
'   End If
'   If Rst!QDueDate4 <> cboQDay(14).text & " " & cboQMth(14).text Then
       rst!QDueDate4 = cboQDay(14).text & " " & cboQMth(14).text
'   End If
'   If Rst!HYDueDate1 <> cboHDay(5).text & " " & cboHMth(5).text Then
        rst!HYDueDate1 = cboHDay(5).text & " " & cboHMth(5).text
'   End If
'   If Rst!HYDueDate2 <> cboHDay(6).text & " " & cboHMth(6).text Then
        rst!HYDueDate2 = cboHDay(6).text & " " & cboHMth(6).text
'   End If
'   If Rst!YDueDate <> cboYDay(2).text & " " & cboYMth(2).text Then
        rst!YDueDate = cboYDay(2).text & " " & cboYMth(2).text
'   End If

   rst!NoOfDaysToSendMFB4Due = CInt(IIf(txtNoOfDaysToSendMFB4Due.text = "", 0, txtNoOfDaysToSendMFB4Due.text))
   rst.Update

   rst.Close

   MsgBox "Your changes have been saved.", vbInformation, "Saved"

   Call DisableBoxes

   conn3.Close
   txtNoOfDaysToSendMFB4Due.Locked = True
   cmdEditFeeNChargePaydates.Enabled = True
   cmdSaveFEEnCharge.Enabled = False
   cmdCancelFeenCharge.Enabled = False
End Sub

Private Sub DisableBoxes()
    fraPaymentDate(14).Enabled = False
    fraPaymentDate(13).Enabled = False
    fraPaymentDate(12).Enabled = False
    fraPaymentDate(11).Enabled = False
    fraPaymentDate(10).Enabled = False
    fraPaymentDate(9).Enabled = False
End Sub
Public Function MissingDate(ctr As Control) As Boolean
   MissingDate = False
   If ctr.text = "" Then
       MsgBox "You must select all the payment dates", vbOKOnly + vbCritical, "Missing Payment Date"
       FocusControl ctr
       MissingDate = True
   End If
End Function

Private Sub cmdSavePaymentDetails_Click()
    If txtPaymentType.Tag = "" Then
            MsgBox "Please enter a payment type to save", vbInformation, "Warning"
            FocusControl cmdPaymentType
            Exit Sub
    End If
    cmdPaymentType.Enabled = False
    cmdPaymentTypeNew(0).Enabled = False
    cmdBrowseTemplate.Enabled = False
    txtPaymentTerms.Locked = True
    Dim adoConn As New ADODB.Connection
    Dim rsClient As New ADODB.Recordset
    Dim rsPayments As New ADODB.Recordset
    
    adoConn.Open getConnectionString
    rsClient.Open "Select * from Client where clientID='" & txtClientID.text & "'", adoConn, adOpenKeyset, adLockOptimistic
            rsPayments.Open "Select * from Supplier where SupplierID='" & txtClientID.text & "'", adoConn, adOpenKeyset, adLockOptimistic
            If Trim(txtPaymentType.text) <> "" Then
                        rsPayments!PaymentType = txtPaymentType.Tag
            Else
                        rsPayments!PaymentType = ""
            End If
            rsPayments.Update
    
    rsClient!PaymentTerms = IIf(txtPaymentTerms.text = "", 0, txtPaymentTerms.text)
    If Trim(txtPaymentType.text) <> "" Then
            rsClient!PaymentType = txtPaymentType.Tag
    Else
                rsClient!PaymentType = ""
    End If
    rsClient!RemittanceTemplate = txtRemittanceTemplate.text
    rsClient!RentSummaryTemplate = txtRenSummaryStatement.text
    'txtRenSummaryStatement.text = IIf(IsNull(rsClient!RentSummaryTemplate), "", rsClient!RentSummaryTemplate) 'rsClient!Comments2
    
    rsClient!AccountName = txtClientHomeTel(11).text
    rsClient!SortCode = txtClientHomeTel(12).text
    rsClient!AccountNumber = txtClientHomeTel(13).text
    rsClient!BankPaymentRef = txtClientHomeTel(14).text
    rsClient!UsePayableTemplate = chkUsePayableTemplate.Value
    
    rsClient.Update
    adoConn.Close
    Set adoConn = Nothing
    MsgBox "Your Payment Details have been saved successfully.", vbInformation, "Saved"
    cmdPaymentTypeUpdate.Enabled = True
    cmdSavePaymentDetails.Enabled = False
    txtClientHomeTel(11).Locked = True
    txtClientHomeTel(12).Locked = True
    txtClientHomeTel(13).Locked = True
    txtClientHomeTel(14).Locked = True
    chkUsePayableTemplate.Enabled = False
End Sub

Private Sub cmdVAT_Click()
    strCommandSource = "VAT"
    Call LoadVAT
End Sub
Private Sub LoadVAT()
   Frame5.Top = 225
   Frame5.Left = 6255
  
   flxClientList.ColWidth(0) = 1000
   flxClientList.ColWidth(1) = 2000
   flxClientList.TextMatrix(0, 0) = "CODE"
   flxClientList.TextMatrix(0, 1) = "RATE"
   lblClientID(2).Visible = False
   TextBox1.Visible = False
   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblClientID(0).Width = 900
   lblClientID(0).Left = 50
   lblClientID(1).Width = 1900
   lblClientID(1).Left = lblClientID(0).Left + flxClientList.ColWidth(0)
   
   txtSearchClientID.Width = 900
   txtSearchClientID.Left = 40
   txtSearchClientName.Left = 1000
   
   TextBox1.Width = 1900
   TextBox1.Left = txtSearchClientID.Left + flxClientList.ColWidth(0)
   
'   txtSearch3(0).Visible = False
   
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblClientID(0).Caption = "CODE"
   lblClientID(1).Caption = "RATE"
'   lblSearch2(0).Visible = False
'   lblSearch3(0).Visible = False
'   lblSearch4(0).Visible = False
   
   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString
'
   szSQL = "SELECT VAT_CODE, VAT_RATE,VAT_ID " & _
           "FROM tlbVatCode where IN_USE;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxClientList.Clear
      flxClientList.Cols = 3
      flxClientList.Rows = 2
      flxClientList.ColWidth(2) = 0
      flxClientList.RowHeight(0) = 0

      rstRec.MoveFirst
      flxClientList.ColAlignment(1) = vbRightJustify

      flxClientList.TextMatrix(0, 0) = "VAT Code"
      flxClientList.TextMatrix(0, 1) = "VAT Rate"

      rRow = 1
      flxClientList.AddItem ""
      While Not rstRec.EOF
         flxClientList.TextMatrix(rRow, 0) = rstRec!VAT_CODE
         flxClientList.TextMatrix(rRow, 1) = rstRec!VAT_RATE
         flxClientList.TextMatrix(rRow, 2) = rstRec!VAT_ID
         rstRec.MoveNext
         If Not rstRec.EOF Then flxClientList.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close

   Set rstRec = Nothing
   Set Conn2 = Nothing
   Frame5.Visible = True
End Sub

Private Sub Command1_Click()
    Call LoadFlxACHistory_Old1
End Sub

Private Sub Command2_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call LoadFlxACHistory(adoConn, "")
    adoConn.Close
End Sub

Private Sub Command3_Click()
    
End Sub

Private Sub flxManagementFee_displaytextboxes()
    
    rRow = flxManagementFee.row
    If rRow = 0 Then Exit Sub
    If flxManagementFee.TextMatrix(rRow, 1) = "" Then Exit Sub
    strtlbAgreementID = flxManagementFee.TextMatrix(rRow, 1) ' = rsPayable.Fields.Item("PAYABLE_ID").Value
    txtChargeType.Tag = flxManagementFee.TextMatrix(rRow, 25) ' Charge type ID
    txtChargeType.text = flxManagementFee.TextMatrix(rRow, 4)        'Charge type Description
'    txtDemandTypemngtFee.Tag = flxManagementFee.TextMatrix(rRow, 5)  '= rsPayable.Fields.Item("FundID").Value
'    txtDemandTypemngtFee.text = flxManagementFee.TextMatrix(rRow, 6)  '= rsPayable.Fields.Item("FundName").Value
    txtFundMngtFee.Tag = flxManagementFee.TextMatrix(rRow, 7) ' = rsPayable.Fields.Item("clientLandlordID").Value
    txtFundMngtFee.text = flxManagementFee.TextMatrix(rRow, 8)   '= IIf(rsPayable.Fields.Item("PAYABLE_BASIS_ ").Value = "TA", "Total Amount", "Percentage")
    txtManagingAgentAC.Tag = flxManagementFee.TextMatrix(rRow, 9)   'Percentage
    txtManagingAgentAC.text = flxManagementFee.TextMatrix(rRow, 9)
    txtChargingMethod.text = flxManagementFee.TextMatrix(rRow, 11)
    txtChargingMethod.Tag = flxManagementFee.TextMatrix(rRow, 10)
    txtChargeBasis.Tag = IIf(flxManagementFee.TextMatrix(rRow, 12) = "Percentage", "PC", "AN")
    txtChargeBasis.text = flxManagementFee.TextMatrix(rRow, 12) 'IIf(flxManagementFee.TextMatrix(rRow, 12) = "PC", "Percentage", "Annual")
    txtSTART_DATE = flxManagementFee.TextMatrix(rRow, 13)
    txtFrequecymngtFee.Tag = flxManagementFee.TextMatrix(rRow, 14)
    txtFrequecymngtFee.text = flxManagementFee.TextMatrix(rRow, 15)
    txtNtDueDate.text = flxManagementFee.TextMatrix(rRow, 17)
    txtAmount.text = flxManagementFee.TextMatrix(rRow, 18)
    txtTotalAmountPerYear.text = flxManagementFee.TextMatrix(rRow, 19)
    txtPeriod.text = flxManagementFee.TextMatrix(rRow, 20)
    txtLastChargeDate.text = flxManagementFee.TextMatrix(rRow, 21)
    txtStopDatemngtFee.text = flxManagementFee.TextMatrix(rRow, 22) 'Need to last charge date before this
    txtCapAmount.text = flxManagementFee.TextMatrix(rRow, 23)
    txtEND_DATE.text = flxManagementFee.TextMatrix(rRow, 24)

    cmdAgmntEdit.Enabled = True
    cmdDeleteMgtFee.Enabled = True
    cmdAgmntCancel.Enabled = True
    FocusControl cmdCommandArray(9)
End Sub

'  Build up CLIENTs' Account BALANCE
Private Sub flxMemoDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     flxMemoDetails.ToolTipText = flxMemoDetails.TextMatrix(flxMemoDetails.MouseRow, flxMemoDetails.MouseCol)
End Sub
Private Sub flxMemoDetails_Click()
      txtMemoID.text = flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)
      txtUnitMemo.text = flxMemoDetails.TextMatrix(flxMemoDetails.row, 5)
      cmdUnitMemoEdit.Enabled = True
      cmdDelete.Enabled = True
End Sub
Private Sub cmdCloseMemo_Click()
    fraAllMemo.Visible = False
    txtUnitMemo.SetFocus
    cmdVAMemo.Visible = True
End Sub
Private Sub cmdVAMemo_Click()
        fraAllMemo.Visible = True
        txtMemoAll.text = ""
        Call ViewMemo
        cmdVAMemo.Visible = False
        txtMemoAll.SetFocus
End Sub
Private Sub cmdUnitMemoNew_Click()
    Memo_Save_mode = True
'   fmeTenant.Enabled = False
   cmdUnitMemoNew.Enabled = False
   cmdUnitMemoEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdUnitMemoSave.Enabled = True
   cmdUnitMemoCancel.Enabled = True
   flxMemoDetails.Enabled = False
   txtUnitMemo.Locked = False
   fraAllMemo.Visible = False
   txtUnitMemo.text = ""
   txtUnitMemo.SetFocus
End Sub
Private Sub cmdUnitMemoEdit_Click()
   'Modified by Anol 02 Nov 2014
   'Issue 488  Memo and attachment not saving a record of the date of the memo entry
      fraAllMemo.Visible = False
       If txtMemoID.text = "" Then
            cmdVAMemo.Enabled = True
      Else
            cmdVAMemo.Enabled = False
      End If
      cmdVAMemo.Visible = True
      If txtMemoID.text = "" Then
          ShowMsgInTaskBar "Please select the memo you would like to edit", "Y"
          If flxMemoDetails.Enabled = True Then
               flxMemoDetails.SetFocus
          End If
          Exit Sub
      End If
      
      cmdUnitMemoNew.Enabled = False
      cmdUnitMemoEdit.Enabled = False
      cmdUnitMemoSave.Enabled = True
      cmdUnitMemoCancel.Enabled = True

'      fmeTenant.Enabled = False
      txtUnitMemo.Locked = False
      Memo_Save_mode = False
   If txtUnitMemo.Enabled = True Then
      txtUnitMemo.SetFocus
   End If
End Sub
Private Sub cmdUnitMemoSave_Click()
   cmdVAMemo.Visible = False
   If Len(txtUnitMemo.text) = 0 Then
      ShowMsgInTaskBar "Please enter description of memo", "Y"
      If txtUnitMemo.Enabled = True Then
         txtUnitMemo.SetFocus
      End If
      Exit Sub
   End If
   If flxMemoDetails.row = 0 And Memo_Save_mode = False Then
       ShowMsgInTaskBar "Please select a memo from list", "Y"
       Exit Sub
   End If
   
   If SaveMemo Then
      ShowMsgInTaskBar "The memo has been saved successfully."
      LoadGridMemo
   Else
      ShowMsgInTaskBar "Could not save lease analysis", , "N"
   End If
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdUnitMemoCancel.Enabled = False
   flxMemoDetails.Enabled = True
   flxMemoDetails.row = 0
   txtUnitMemo.text = ""
   txtMemoID.text = ""
   txtUnitMemo.Locked = True
'   fmeTenant.Enabled = True
   cmdDelete.Enabled = False
   cmdVAMemo.Enabled = True
   fraAllMemo.Visible = True
   txtMemoAll.text = ""
   Call ViewMemo
   txtMemoAll.SetFocus
End Sub
Private Sub cmdDelete_Click()
      fraAllMemo.Visible = False
      If flxMemoDetails.row = 0 Then
          ShowMsgInTaskBar "Please select a memo from the list", "Y"
          If flxMemoDetails.Enabled = True Then
               flxMemoDetails.SetFocus
          End If
          Exit Sub
      End If
      If MsgBox("Are you sure to delete memo?", vbQuestion + vbYesNo, "Delete Memo") = vbNo Then Exit Sub
      Dim adoConn As New ADODB.Connection
      adoConn.Open getConnectionString
      adoConn.Execute "DELETE from MemoDetails where MemoID=" & Val(flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)) & " and sageaccountNumber='" & txtClientID.text & "'"
      
      adoConn.Close
      MsgBox "Memo has been deleted successfully", vbInformation + vbOKOnly, "Delete Memo"
      
      LoadGridMemo
      
      cmdUnitMemoNew.Enabled = True
      cmdUnitMemoEdit.Enabled = True
      cmdUnitMemoSave.Enabled = False
      cmdUnitMemoCancel.Enabled = False
      flxMemoDetails.Enabled = True
      flxMemoDetails.row = 0
      txtUnitMemo.text = ""
      txtMemoID.text = ""
      txtUnitMemo.Locked = True
'      fmeTenant.Enabled = True
      cmdDelete.Enabled = False
      fraAllMemo.Visible = True
      txtMemoAll.text = ""
      Call ViewMemo
      txtMemoAll.SetFocus
End Sub
Private Sub cmdUnitMemoCancel_Click()
   'Issue 488
   'Modified by anol 04 Oct 2014
'   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   'MemoButtonEnable False
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdDelete.Enabled = False
   txtUnitMemo.Locked = True
   flxMemoDetails.Enabled = True
   txtUnitMemo.text = ""
   fraAllMemo.Visible = True
   cmdVAMemo.Enabled = True
   cmdVAMemo.Visible = False
   txtMemoAll.SetFocus
End Sub
Private Function NewMemoID() As Integer
   Dim conMemo As New ADODB.Connection
   conMemo.Open getConnectionString
   Dim szSQL As String
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT MAX(MemoID) AS x   " & _
                 "FROM MemoDetails;"
   rstSet.Open szSQL, conMemo, adOpenStatic, adLockReadOnly

   NewMemoID = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
   rstSet.Close
   Set rstSet = Nothing
   conMemo.Close
End Function

Private Function SaveMemo() As Boolean
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim conMemo As New ADODB.Connection
   Dim rstLease_ As New ADODB.Recordset
   conMemo.Open getConnectionString
   Dim sSQLQuery_ As String
   Dim sSQLFilter As String
   If Not Memo_Save_mode Then
       sSQLFilter = "WHERE MemoID = " & Val(flxMemoDetails.TextMatrix(flxMemoDetails.row, 1)) & " AND Memotype='Client' AND SageAccountNumber = '" & txtClientID.text & "'"
   Else
       sSQLFilter = ""
   End If
   sSQLQuery_ = "SELECT * " & _
                "FROM MemoDetails " & sSQLFilter
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenDynamic, adLockOptimistic
   If Memo_Save_mode Then rstLeaseAnalysis_.AddNew
   If Memo_Save_mode = False Then
      rstLeaseAnalysis_!MemoID = txtMemoID.text
   Else
      rstLeaseAnalysis_!MemoID = NewMemoID()
   End If
   
   rstLeaseAnalysis_!MemoType = "Client"
   rstLeaseAnalysis_!SageAccountNumber = txtClientID.text
   rstLeaseAnalysis_!MemoDescription = IIf(txtUnitMemo.text <> "", txtUnitMemo.text, "")
   rstLeaseAnalysis_!UpdateTime = Now
   rstLeaseAnalysis_!UserName = frmMMain.SystemUserName
   rstLeaseAnalysis_.Update
   rstLeaseAnalysis_.Close
   Set rstLease_ = Nothing
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   SaveMemo = True
End Function
Public Sub LoadGridMemo()
   'Issue 488
   'Added by anol 03 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtClientID.text & "' And  MemoType='Client' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
   Dim iRow As Integer
   iRow = 1

   flxMemoDetails.Clear
   flxMemoDetails.Rows = 1
   flxMemoDetails.Cols = 7
   flxMemoDetails.ColWidth(0) = 0 'Label12.Left - Label12.Left   'Serial No
   flxMemoDetails.ColWidth(1) = 0
   flxMemoDetails.ColWidth(2) = 0
   flxMemoDetails.ColWidth(3) = 0
   flxMemoDetails.ColWidth(4) = Label6(10).Left - Label6(9).Left    'UpdateTime
   flxMemoDetails.ColWidth(5) = Label8.Left - Label6(10).Left    'MemoDescription
   flxMemoDetails.ColWidth(6) = 2000                         'UserName
   flxMemoDetails.RowHeight(0) = 0
   If rstLeaseAnalysis_.EOF = True Then
       flxMemoDetails.Rows = 2
   End If
   
   While Not rstLeaseAnalysis_.EOF
      flxMemoDetails.AddItem ""
      flxMemoDetails.TextMatrix(iRow, 0) = iRow
      flxMemoDetails.TextMatrix(iRow, 1) = rstLeaseAnalysis_!MemoID
      flxMemoDetails.TextMatrix(iRow, 2) = rstLeaseAnalysis_!MemoType 'col size 0
      flxMemoDetails.TextMatrix(iRow, 3) = rstLeaseAnalysis_!SageAccountNumber 'col size 0
      flxMemoDetails.TextMatrix(iRow, 4) = rstLeaseAnalysis_!UpdateTime
      flxMemoDetails.TextMatrix(iRow, 5) = rstLeaseAnalysis_!MemoDescription
      flxMemoDetails.TextMatrix(iRow, 6) = rstLeaseAnalysis_!UserName
      rstLeaseAnalysis_.MoveNext
      iRow = iRow + 1
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   If iRow > 0 Then
      flxMemoDetails.row = 0
   End If
End Sub
Private Sub ViewMemo()
   'Issue 488
   'Added by anol 04 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtClientID.text & "' And  MemoType='Client' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
  Dim strTemp As String
   While Not rstLeaseAnalysis_.EOF
         If Len(rstLeaseAnalysis_!UpdateTime) > 0 Then
               strTemp = " -  "
         Else
               strTemp = ""
         End If
         If Len(txtMemoAll.text) > 0 Then txtMemoAll.text = txtMemoAll.text & vbCrLf & vbCrLf
         txtMemoAll.text = txtMemoAll.text & Left(rstLeaseAnalysis_!UpdateTime, 11) & strTemp & rstLeaseAnalysis_!UserName & vbCrLf & vbCrLf & IIf(IsNull(rstLeaseAnalysis_!MemoDescription) = True, "", rstLeaseAnalysis_!MemoDescription)
         rstLeaseAnalysis_.MoveNext
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   cmdCloseMemo.Refresh
End Sub
Private Function GetClientBalance(szClientID As String) As Currency
   Dim j As Integer

   For j = 0 To UBound(szaClientBal, 2) - 1
      If szClientID = szaClientBal(0, j) Then
         GetClientBalance = Format(szaClientBal(1, j), "0.00")
         Exit For
      End If
   Next j
   If j = UBound(szaClientBal, 2) Then GetClientBalance = 0
End Function
Private Sub ClientAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim szSqlPI As String
   Dim szSQLSI As String
   Dim i       As Integer
   Dim iSI     As Integer
   Dim iPI     As Integer
   Dim iIndex  As Integer

   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset
   Dim adoRptDr As New ADODB.Recordset, adoRptCr As New ADODB.Recordset

'-----------------      Purchase Side    -----------------------------------
   szSqlPI = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "FROM   tlbPayment, Client " & _
             "WHERE  tlbPayment.SageAccountNumber = Client.ClientID " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSqlPI, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaClientBal(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
           "FROM tlbPayment AS P, Client " & _
           "WHERE (P.Type = 6 OR P.Type = 24) AND P.SageAccountNumber = Client.ClientID " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaClientBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaClientBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
           "FROM tlbPayment AS P, Client " & _
           "WHERE P.Type <> 6 AND P.Type <> 24 AND P.SageAccountNumber = Client.ClientID " & _
           "GROUP BY P.SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaClientBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaClientBal(1, i) = szaClientBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaClientBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaClientBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoRptDr = Nothing
   Set adoRptCr = Nothing
End Sub
Private Sub cboBank_ID_Click()
   'Resolved by BOSL
   'Error while clicking exit button
   'Modified by anol 11 Mar 2015
   If cboBank_ID.ListIndex <> -1 Then
      txtBANK_NAME.text = cboBank_ID.Column(1)
      txtBANK_ADDRESS1.text = cboBank_ID.Column(3)
      txtBANK_ADDRESS2.text = cboBank_ID.Column(5)
      txtBANK_ADDRESS3.text = cboBank_ID.Column(6)
      txtBANK_POST_CODE.text = cboBank_ID.Column(4)
      txtBANK_SC.text = cboBank_ID.Column(2)
   End If
End Sub

'Private Sub cboCHARGE_BASIS_Click()
'   If cboCHARGE_BASIS.text = "%" And cboCHARGE_METHOD.text = "FIXED" Then
'      MsgBox "Charge Method is ''Fixed'' in this case % is not valid choice."
'      cboCHARGE_BASIS.text = ""
'   End If
'End Sub

'Private Sub cboDmdPropertyList_Click()
'   Label19(3).ForeColor = vbBlack
'
'   LoadGlobalData cboDmdPropertyList.Value
'End Sub

'Private Sub cboFrequency_LostFocus()
'   If txtSTART_DATE.text = "" Or txtEND_DATE.text = "" Or cboFrequency.text = "" Then Exit Sub
'
'   Dim adoconn As New ADODB.Connection
'   Dim adoRst As ADODB.Recordset
'
'   adoconn.Open getConnectionString
'
'   '****** I am adding a validation when global data isnt set then exit current function
'
'   Dim i As Integer, iDateSet As Integer
'   Dim Rst As ADODB.Recordset
'   Dim SQLStr As String
'
'   Set Rst = New ADODB.Recordset
'
'   SQLStr = "SELECT PaymentDates FROM ChargeTypes WHERE ID = " & CInt(cboCHARGE_TYPE.Value) & ";"
'   Rst.Open SQLStr, adoconn, adOpenDynamic, adLockPessimistic
'
'   iDateSet = Rst.Fields(0).Value
'
'   Rst.Close
'
'   If iDateSet = 0 Then
'      SQLStr = "SELECT Record_ID, ClientID, QuarterlyDueDate1, QuarterlyDueDate2, " & _
'                  "QuarterlyDueDate3, QuarterlyDueDate4, HalfYearlyDueDate1, " & _
'                  "HalfYearlyDueDate2, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
'                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, " & _
'                  "MonthlyDueDate8, MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, " & _
'                  "MonthlyDueDate12, YearlyDueDate, FeeIsuDays, PayIsuDays, LettingFee, " & _
'                  "LettingAM, LettingFreq, LettingNtDueDt, LettingStDt, LettingChrgType, " & _
'                  "MngFee, MngAM, MngFreq, MngNtDueDt, MngStDt, MngChrgType, RentPayble, " & _
'                  "RentAM, RentFreq, RentNtDueDt, RentStDt, RentChrgType " & _
'               "FROM ClientGlobalData " & _
'               "WHERE ClientGlobalData.ClientID = '" & txtClientID.text & "';"
'   Else
'      SQLStr = "SELECT DateSetID, NameOfSet, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
'                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, " & _
'                  "MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, " & _
'                  "QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
'                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate " & _
'               "FROM PaymentDates WHERE DateSetID = " & iDateSet & ";"
'   End If
'   Rst.Open SQLStr, adoconn, adOpenDynamic, adLockPessimistic
'
'   If Rst.EOF Then
'       MsgBox "You Need to Enter the Client Global Data.", vbOKOnly + vbInformation, "Global Data"
'       Rst.Close
'       Exit Sub
'   End If
'
'
'   'dtEndDate As Date, dtNtDueDate As Date, iFreq As Integer, ByVal adoconn As ADODB.Connection, iCTId As Integer, szChrPayTypes As String
'   txtNtDueDate.text = FindNextDueDate(CDate(txtSTART_DATE.text), CInt(cboFrequency.Value), adoconn, CInt(cboCHARGE_TYPE.Value))
''DateDiff("d", "04/01/2020", "05/01/2020") results : 1
'   'So is if dtEndDate is less than FindNextDueDate this due date will calculate with end date
'   If IsDate(txtEND_DATE.text) Then
'        If DateDiff("d", txtEND_DATE.text, txtNtDueDate.text) > 0 Then
'           txtNtDueDate.text = txtEND_DATE.text
'        End If
'   End If
'
'   adoconn.Close
'   Set adoconn = Nothing
'End Sub

Private Function FindNextDueDate(dtNtDueDate As Date, iFreq As Integer, ByVal adoConn As ADODB.Connection, iCTId As Integer) As Date
   'iCTId is payable type ID
   GetClientGlobalDatabyProperty txtClientID, adoConn, iCTId

   Select Case iFreq
      Case 1:                               'Weekly in advance
         FindNextDueDate = dtNtDueDate
      Case 2:                               'Weekly in arrears
         FindNextDueDate = DateAdd("d", 7, dtNtDueDate)
      Case 3:                               'Fortnightly in advance
         FindNextDueDate = dtNtDueDate
      Case 4:                               'Fortnightly in arrears
         FindNextDueDate = DateAdd("d", 14, dtNtDueDate)
      Case 5:                               'Monthly in advance
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InAdv, Pay_Monthly)
      Case 6:                               'Monthly in arrears
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InArr, Pay_Monthly)
      Case 7:                               'Quarterly in advance
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InAdv, Pay_Quarterly)
      Case 8:                               'Quarterly in arrears
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InArr, Pay_Quarterly)
      Case 9:                               'Half yearly in advance
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InAdv, Pay_Half_Yearly)
      Case 10:                               'Half yearly in arrears
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InArr, Pay_Half_Yearly)
      Case 11:                              'yearly in advance
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InAdv, Pay_Yearly)
      Case 12:                              'yearly in arrears
         FindNextDueDate = ClNextPayingDate(dtNtDueDate, InArr, Pay_Yearly)
   End Select
   
End Function
Private Sub GetClientGlobalDatabyProperty(szClientID As String, Conn As ADODB.Connection, iCTId As Integer)

   Dim i As Integer, iDateSet As Integer
   Dim rst As ADODB.Recordset
   Dim SQLStr As String

   Set rst = New ADODB.Recordset

   SQLStr = "SELECT Record_ID, ClientID, QuarterlyDueDate1, QuarterlyDueDate2, " & _
                  "QuarterlyDueDate3, QuarterlyDueDate4, HalfYearlyDueDate1, " & _
                  "HalfYearlyDueDate2, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, " & _
                  "MonthlyDueDate8, MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, " & _
                  "MonthlyDueDate12, YearlyDueDate, FeeIsuDays, PayIsuDays, LettingFee, " & _
                  "LettingAM, LettingFreq, LettingNtDueDt, LettingStDt, LettingChrgType, " & _
                  "MngFee, MngAM, MngFreq, MngNtDueDt, MngStDt, MngChrgType, RentPayble, " & _
                  "RentAM, RentFreq, RentNtDueDt, RentStDt, RentChrgType " & _
               "FROM ClientGlobalData " & _
               "WHERE ClientGlobalData.ClientID = '" & szClientID & "' ;"
   
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
       MsgBox "You Need to Enter the Client Global Data.", vbOKOnly + vbInformation, "Global Data"
       rst.Close
       Exit Sub
   End If

   szClYearly = rst!YearlyDueDate

   szClHalfYearly1 = rst!HalfYearlyDueDate1
   szClHalfYearly2 = rst!HalfYearlyDueDate2

   szClQuarterly1 = rst!QuarterlyDueDate1
   szClQuarterly2 = rst!QuarterlyDueDate2
   szClQuarterly3 = rst!QuarterlyDueDate3
   szClQuarterly4 = rst!QuarterlyDueDate4

   szaMonthlyCl(0) = rst!MonthlyDueDate1
   szaMonthlyCl(1) = rst!MonthlyDueDate2
   szaMonthlyCl(2) = rst!MonthlyDueDate3
   szaMonthlyCl(3) = rst!MonthlyDueDate4
   szaMonthlyCl(4) = rst!MonthlyDueDate5
   szaMonthlyCl(5) = rst!MonthlyDueDate6
   szaMonthlyCl(6) = rst!MonthlyDueDate7
   szaMonthlyCl(7) = rst!MonthlyDueDate8
   szaMonthlyCl(8) = rst!MonthlyDueDate9
   szaMonthlyCl(9) = rst!MonthlyDueDate10
   szaMonthlyCl(10) = rst!MonthlyDueDate11
   szaMonthlyCl(11) = rst!MonthlyDueDate12

   rst.Close
End Sub
'Private Sub cboFund_DropButtonClick()
'   cboFund.ListStyle = fmListStylePlain
'   cboFund.ListRows = 5
'End Sub

Private Sub cboNC_LostFocus()
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   If txtNCCODE.text = "" Then Exit Sub

   If bBankNewEdit Then                                                'Adding new bank account
      adoConn.Open getConnectionString

      szSQL = "SELECT * " & _
              "FROM tlbClientBanks " & _
              "WHERE CLIENT_ID = '" & txtClientID.text & "' AND " & _
                  "NominalCode = '" & txtNCCODE.text & "';"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      If Not adoRST.EOF Then
         MsgBox "This nominal code is already in use against another bank account." & vbNewLine & _
               vbTab & "Please select another code.", vbCritical + vbOKOnly, "Client Bank"
         txtNCCODE.text = ""
         txtNominal.text = ""
         FocusControl cmdNC
      End If

      adoRST.Close
      adoConn.Close
      Set adoRST = Nothing
      Set adoConn = Nothing
   End If
   If Not bBankNewEdit Then                  'Adding new bank account
      adoConn.Open getConnectionString

      szSQL = "SELECT * " & _
              "FROM tlbClientBanks " & _
              "WHERE CLIENT_ID = '" & txtClientID.text & "' AND " & _
                  "NominalCode = '" & txtNCCODE.text & "' AND " & _
                  "MY_ID <> " & flxOtherBankDetails.TextMatrix(iSlectedRow, 13) & ";"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      If Not adoRST.EOF Then
         MsgBox "There is a bank details has been booked for this Nominal Code." & vbNewLine & _
               vbTab & "Please select another code.", vbCritical + vbOKOnly, "Client Bank"
         txtNCCODE.text = ""
         txtNominal.text = ""
         FocusControl cmdNC
      End If

      adoRST.Close
      adoConn.Close
      Set adoRST = Nothing
      Set adoConn = Nothing
   End If
End Sub

Private Sub FillPayDueDate() 'Call this function when you click Gridview  this fucntions hall populate next due date
        Exit Sub
   'If txtPAY_START_DATE.text = "" Or txtPayFrequency.text = "" Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoRST As ADODB.Recordset
   adoConn.Open getConnectionString
   Dim i As Integer, iDateSet As Integer
   Dim rst As ADODB.Recordset
   Dim SQLStr As String

   Set rst = New ADODB.Recordset

'   SQLStr = "SELECT PaymentDates FROM PayableTypes WHERE ID =" & CInt(txtPayableType.Tag) & ""
'   Rst.Open SQLStr, adoConn, adOpenDynamic, adLockPessimistic
'
'   iDateSet = Rst.Fields(0).Value
'
'   Rst.Close
'
'   If iDateSet = 0 Then
      SQLStr = "SELECT Record_ID, ClientID, QuarterlyDueDate1, QuarterlyDueDate2, " & _
                  "QuarterlyDueDate3, QuarterlyDueDate4, HalfYearlyDueDate1, " & _
                  "HalfYearlyDueDate2, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, " & _
                  "MonthlyDueDate8, MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, " & _
                  "MonthlyDueDate12, YearlyDueDate, FeeIsuDays, PayIsuDays, LettingFee, " & _
                  "LettingAM, LettingFreq, LettingNtDueDt, LettingStDt, LettingChrgType, " & _
                  "MngFee, MngAM, MngFreq, MngNtDueDt, MngStDt, MngChrgType, RentPayble, " & _
                  "RentAM, RentFreq, RentNtDueDt, RentStDt, RentChrgType " & _
               "FROM ClientGlobalData " & _
               "WHERE ClientGlobalData.ClientID = '" & txtClientID.text & "';"
'   Else
'      SQLStr = "SELECT DateSetID, NameOfSet, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
'                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, " & _
'                  "MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, " & _
'                  "QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
'                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate " & _
'               "FROM PaymentDates WHERE DateSetID = " & iDateSet & ";"
'   End If
   rst.Open SQLStr, adoConn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
       MsgBox "You Need to Enter the Client Global Data.", vbOKOnly + vbInformation, "Global Data"
       rst.Close
       Exit Sub
   End If

    'dtEndDate As Date, dtNtDueDate As Date, iFreq As Integer, ByVal adoconn As ADODB.Connection, iCTId As Integer, szChrPayTypes As String
   'txtPAY_NtDueDate.text = FindNextDueDate(CDate(txtPAY_START_DATE.text), CInt(txtPayFrequency.Tag), adoconn, CInt(txtPayableType.Tag))
 'So is if dtEndDate is less than FindNextDueDate this due date will calculate with end date
'    If IsDate(txtPAY_END_DATE.text) Then
'        If DateDiff("d", txtPAY_END_DATE.text, txtPAY_NtDueDate.text) > 0 Then
'           txtPAY_NtDueDate.text = txtPAY_END_DATE.text
'        End If
'   End If
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadflxAgreement(szPropertySel As String)
'  table ClientProAgr is a header level table and tlbagreement is a child level table

    Dim sSQLQuery_ As String, sFilter As String
    Dim szaPropertyID() As String
    
    Dim adoConn As New ADODB.Connection
    Dim rstAgreement As New ADODB.Recordset
    adoConn.Open getConnectionString
    
    
    '*****************************************agreement details*********************************************
    'one property per client so we dont need to put a Property filter here. rather Select property from table and show it in the combobox
    
    sSQLQuery_ = "SELECT *,ClientProAgr.PropertyID as propID " & _
                 "FROM ClientProAgr WHERE ClientID = '" & txtClientID.text & "' and PropertyID='" & szPropertySel & "';"
    rstAgreement.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
    
    If rstAgreement.EOF Then
         cmdAgrTopEdit.Caption = "Add Agreement"
    Else
         cmdAgrTopEdit.Caption = "Edit"
         'Because it can save only one property so show first property
    
    End If
    txtAgreementStartDate.Enabled = False
    txtAgreementEndDate.Enabled = False
    txtREVIEW_DATE.Enabled = False
    With rstAgreement
    If Not .EOF Then
          txtREVIEW_DATE.text = IIf(IsNull(!REVIEW_DATE) = True, "", !REVIEW_DATE)
          txtAgreementStartDate.text = IIf(IsNull(!agreementStartDate) = True, "", !agreementStartDate)
          txtAgreementEndDate.text = IIf(IsNull(!agreementEndDate) = True, "", !agreementEndDate) '!agreementEndDate
          tabAgreement.Enabled = True
          cmdAgrTopEdit.Enabled = True
          cmdAgrTopSave.Enabled = False
          cmdAgmntAddNew.Enabled = True
    Else
          cmdAgrTopEdit.Enabled = True
          cmdAgrTopSave.Enabled = False
          cmdAgmntAddNew.Enabled = False
          txtREVIEW_DATE.text = ""
          txtNoOfDaysToSendMFB4Due.text = ""
          txtAgreementStartDate.text = ""
          txtAgreementEndDate.text = ""
          tabAgreement.Enabled = False
    End If
    .Close
    End With
   Set rstAgreement = Nothing
'*****************************************Management Fees*********************************************
 'Dim strLastChargeDate As String
    'strLastChargeDate = findLastChargeDate(szPropertySelection1, adoconn)
   Call loadflxManagementFee(adoConn)   'main loading procedure

   '******************************************** Rent Payable ***************************************
   Call loadflxPayable(adoConn)   'main loading procedure
   

   adoConn.Close
   Set adoConn = Nothing

End Sub
Private Sub loadflxManagementFee(adoConn As ADODB.Connection)
    Dim rRow As Integer
    Dim rsManagementFee As New ADODB.Recordset
    Call ConfigflxManagementFee
    Dim szSQL As String
'    Dim strLastChargeDate As String
    szSQL = "SELECT agr.*,C.FeeType ,C.ID as Charge_Type_ID, " & _
            "F.FundID,F.FundName,agr.Frequency as FREQID,(Select FC.Frequency from Frequencies FC " & _
            "where  FC.ID=agr.Frequency) as Frequency ,SC.CODE as chBASISCODE, SC.VALUE as CHRBASIS  " & _
            "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
            "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND SC.CODE=agr.CHARGE_METHOD AND " & _
            "CPA.ClientID = '" & txtClientID.text & "' And C.ID = agr.CHARGE_TYPE And " & _
            "CPA.PropertyID = '" & szPropertySelection1 & "';"

'charging method SELECT CODE, VALUE FROM SECONDARYCODE WHERE PRIMARYCODE = 'CRGBS'
   rsManagementFee.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If rsManagementFee.EOF Then
        Label7(13).Caption = "No Management Fees has been setup."
        rsManagementFee.Close
        Set rsManagementFee = Nothing
        Exit Sub
   Else
        Label7(13).Caption = ""
        rRow = 1

        While Not rsManagementFee.EOF
            flxManagementFee.TextMatrix(rRow, 0) = "" 'selection column
            flxManagementFee.TextMatrix(rRow, 1) = rsManagementFee.Fields.Item("AGREEMENT_ID").Value 'ManagementFee_ID
            flxManagementFee.TextMatrix(rRow, 2) = rsManagementFee.Fields.Item("CPA_ID").Value
            flxManagementFee.TextMatrix(rRow, 3) = rsManagementFee.Fields.Item("AGREEMENT_ID").Value 'ManagementFee_ID
            flxManagementFee.TextMatrix(rRow, 4) = rsManagementFee.Fields.Item("FeeType").Value 'Charge Description /CHARGE_TYPE
'            flxManagementFee.TextMatrix(rRow, 5) = rsManagementFee.Fields.Item("DemandID").Value 'Demand Type ID
'            flxManagementFee.TextMatrix(rRow, 6) = rsManagementFee.Fields.Item("DEMAND_TYPE").Value 'Demand Type Description
            flxManagementFee.TextMatrix(rRow, 7) = rsManagementFee.Fields.Item("Fund").Value 'Fund ID
            flxManagementFee.TextMatrix(rRow, 8) = rsManagementFee.Fields.Item("FundName").Value ' Fund Description
            flxManagementFee.TextMatrix(rRow, 9) = rsManagementFee.Fields.Item("ManagingAgentID").Value
            flxManagementFee.TextMatrix(rRow, 10) = rsManagementFee.Fields.Item("CHARGE_METHOD").Value
            flxManagementFee.TextMatrix(rRow, 11) = rsManagementFee.Fields.Item("CHRBASIS").Value
            flxManagementFee.TextMatrix(rRow, 12) = IIf(rsManagementFee.Fields.Item("CHARGE_BASIS").Value = "PC", "Percentage", "ANNUAL")
            flxManagementFee.TextMatrix(rRow, 13) = Format(rsManagementFee.Fields.Item("START_DATE").Value, "dd/mm/yyyy")
            If flxManagementFee.TextMatrix(rRow, 13) = "" Then
                flxManagementFee.TextMatrix(rRow, 13) = "N/A"
            End If
            flxManagementFee.TextMatrix(rRow, 14) = IIf(IsNull(rsManagementFee.Fields.Item("FREQID").Value), "N/A", rsManagementFee.Fields.Item("FREQID").Value)
            flxManagementFee.TextMatrix(rRow, 15) = IIf(IsNull(rsManagementFee.Fields.Item("Frequency").Value), "N/A", rsManagementFee.Fields.Item("Frequency").Value)
            flxManagementFee.TextMatrix(rRow, 16) = rsManagementFee.Fields.Item("CHBASISCODE").Value 'CHRBASIS Value 'IIf(rsManagementFee.Fields.Item("PAYABLE_BASIS_").Value = "TA", "Total Amount", "Percentage")
            flxManagementFee.TextMatrix(rRow, 17) = IIf(IsNull(rsManagementFee.Fields.Item("NtDueDate").Value), "N/A", rsManagementFee.Fields.Item("NtDueDate").Value) 'rsManagementFee.Fields.Item("NtDueDate").Value
            flxManagementFee.TextMatrix(rRow, 18) = Format(rsManagementFee.Fields.Item("Amount").Value, "0.00")  'amount
            If flxManagementFee.TextMatrix(rRow, 18) = "0.00" And flxManagementFee.TextMatrix(rRow, 12) = "ANNUAL" Then
                 flxManagementFee.TextMatrix(rRow, 18) = "N/A"
            End If
            flxManagementFee.TextMatrix(rRow, 19) = Format(rsManagementFee.Fields.Item("TotalAmount").Value, "0.00")  'Total amount
             If flxManagementFee.TextMatrix(rRow, 19) = "0.00" And flxManagementFee.TextMatrix(rRow, 12) = "ANNUAL" Then
                 flxManagementFee.TextMatrix(rRow, 19) = "N/A"
            End If
            flxManagementFee.TextMatrix(rRow, 20) = Format(rsManagementFee.Fields.Item("EachPeriod").Value, "0.00") 'Each Period
             If flxManagementFee.TextMatrix(rRow, 20) = "0.00" And flxManagementFee.TextMatrix(rRow, 12) = "ANNUAL" Then
                 flxManagementFee.TextMatrix(rRow, 20) = "N/A"
            End If
            
            flxManagementFee.TextMatrix(rRow, 21) = IIf(IsNull(rsManagementFee.Fields.Item("LastChargeDate").Value), "", rsManagementFee.Fields.Item("LastChargeDate").Value) 'Display last charge date here
            flxManagementFee.TextMatrix(rRow, 22) = IIf(IsNull(rsManagementFee.Fields.Item("StopDate").Value), "", rsManagementFee.Fields.Item("StopDate").Value)  'StopDate
            flxManagementFee.TextMatrix(rRow, 23) = Format(rsManagementFee.Fields.Item("CapAmount").Value, "0.00") 'Cap Amount
            flxManagementFee.TextMatrix(rRow, 24) = IIf(IsNull(rsManagementFee.Fields.Item("END_DATE").Value), "", rsManagementFee.Fields.Item("END_DATE").Value) 'End Date
                
            If flxManagementFee.TextMatrix(rRow, 24) <> "" Then
                  flxManagementFee.TextMatrix(rRow, 24) = Format(flxManagementFee.TextMatrix(rRow, 24), "dd/mm/yyyy")
            End If
            flxManagementFee.TextMatrix(rRow, 25) = rsManagementFee.Fields.Item("Charge_Type_ID").Value 'Charge Type ID
            flxManagementFee.RowHeight(rRow) = 280
            rsManagementFee.MoveNext
            If Not rsManagementFee.EOF Then flxManagementFee.AddItem ""
            rRow = rRow + 1
         Wend
 
   End If

   rsManagementFee.Close
   Set rsManagementFee = Nothing
   
   Dim DicFundID As New Dictionary
   Dim iRow As Integer
   For iRow = 1 To flxManagementFee.Rows - 1
        If UCase(flxManagementFee.TextMatrix(iRow, 10)) = UCase("RE_ED") Then
            If Not DicFundID.Exists(flxManagementFee.TextMatrix(iRow, 7)) Then
                    DicFundID.Add flxManagementFee.TextMatrix(iRow, 7), flxManagementFee.TextMatrix(iRow, 8)
            Else
                    Label7(13).Caption = "Warning: There is a duplicate " & flxManagementFee.TextMatrix(iRow, 8) & " fund entered in this agreement. Please delete the agreement line with the duplicated fund"
            End If
        End If
   Next
   
End Sub
Private Sub loadflxPayable(adoConn As ADODB.Connection)
    Dim rRow As Integer
    Dim rsPayable As New ADODB.Recordset
    Call ConfigFlxPayable
    Dim szSQL As String
    szSQL = "SELECT P.PAYABLE_ID,P.CPA_ID, T.PayType , PAYABLE_TYPE, P.PayeeType, F.FundID,F.FundName,clientLandlordID, " & _
                    "PAY_START_DATE, PAY_END_DATE,   P.ONDD,P.PAYABLE_BASIS_,PAY_NtDueDate,Percentage,StopDate,PAY_END_DATE " & _
                    "FROM tlbPayable AS P, ClientProAgr AS C, PayableTypes AS T, FUND as F " & _
                    "WHERE  F.FundID=P.PAY_Fund AND P.CPA_ID = C.CPA_ID And C.ClientID = '" & txtClientID.text & "' And " & _
                    "T.ID = P.PAYABLE_TYPE And C.PropertyID = '" & szPropertySelection1 & "';"
'Debug.Print szSQL
   rsPayable.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rsPayable.EOF Then
      Label7(14).Caption = "No Rent Payable record has been setup."
      rsPayable.Close
      Set rsPayable = Nothing
      Exit Sub
   Else
      Label7(14).Caption = ""
        rRow = 1
'         cmdPayEdit.Enabled = False
        While Not rsPayable.EOF
'            flxPayable.row = 1
'            cmdPayEdit.Enabled = True
            flxPayable.TextMatrix(rRow, 0) = ""
            flxPayable.TextMatrix(rRow, 1) = rsPayable.Fields.Item("PAYABLE_ID").Value
            flxPayable.TextMatrix(rRow, 2) = rsPayable.Fields.Item("CPA_ID").Value
            flxPayable.TextMatrix(rRow, 3) = rsPayable.Fields.Item("PAYABLE_TYPE").Value 'Payable ID
            flxPayable.TextMatrix(rRow, 4) = rsPayable.Fields.Item("PayType").Value 'Payable Description
''            flxPayable.TextMatrix(rRow, 5) = rsPayable.Fields.Item("DemandID").Value
'            flxPayable.TextMatrix(rRow, 6) = rsPayable.Fields.Item("PAY_DEMAND_TYPE").Value
            flxPayable.TextMatrix(rRow, 7) = rsPayable.Fields.Item("FundID").Value
            flxPayable.TextMatrix(rRow, 8) = rsPayable.Fields.Item("FundName").Value
            'move clientLandlordID from index 9 to 10 and in 9 insert payable type 2021/08/23
            flxPayable.TextMatrix(rRow, 9) = IIf(IsNull(rsPayable.Fields.Item("PayeeType").Value), "", rsPayable.Fields.Item("PayeeType").Value)
            flxPayable.TextMatrix(rRow, 10) = rsPayable.Fields.Item("clientLandlordID").Value
'            flxPayable.TextMatrix(rRow, 10) = rsPayable.Fields.Item("PAY_START_DATE").Value
'            flxPayable.TextMatrix(rRow, 11) = rsPayable.Fields.Item("FREQID").Value
'            flxPayable.TextMatrix(rRow, 12) = rsPayable.Fields.Item("Frequency").Value
            flxPayable.TextMatrix(rRow, 13) = rsPayable.Fields.Item("ONDD").Value
            flxPayable.TextMatrix(rRow, 14) = rsPayable.Fields.Item("PAYABLE_BASIS_").Value
            flxPayable.TextMatrix(rRow, 15) = IIf(rsPayable.Fields.Item("PAYABLE_BASIS_").Value = "FA", "Full Amount", "Percentage")
            If flxPayable.TextMatrix(rRow, 15) = "Full Amount" Then
                flxPayable.TextMatrix(rRow, 16) = "N/A"
            Else
                flxPayable.TextMatrix(rRow, 16) = IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00"))
            End If
            flxPayable.TextMatrix(rRow, 17) = IIf(IsNull(rsPayable.Fields.Item("StopDate").Value) = True, "", rsPayable.Fields.Item("StopDate").Value)
            flxPayable.TextMatrix(rRow, 18) = IIf(IsNull(rsPayable.Fields.Item("PAY_END_DATE").Value) = True, "", rsPayable.Fields.Item("PAY_END_DATE").Value)
            flxPayable.RowHeight(rRow) = 280
            rsPayable.MoveNext
            If Not rsPayable.EOF Then flxPayable.AddItem ""
            rRow = rRow + 1
         Wend
      
      'SetFlxPayableHeader flxPayable, rsPayable
   End If

   rsPayable.Close
   Set rsPayable = Nothing
End Sub
Private Sub chkOverDraft_Click()
    'you are seeing after status of click here
   If chkOverDraft.Value = 0 Then
      txtOverDraft.text = ""
      If bOverdraftWarning Then
'        MsgBox "Overdraft warnign enabled"
        If balanceisNegative Then
              chkOverDraft.Value = 1
              MsgBox "It is not possible to disable the 'Allow Overdraft' setting currently," & _
              " because the bank balance for the selected bank account is currently showing an overdraft", vbInformation, "Warning!"
        End If
      End If
   Else
      FocusControl txtOverDraft

   End If
End Sub
Private Function balanceisNegative() As Boolean
    Dim dblAmount As Double
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    dblAmount = Format(BankAccBalance(adoConn, flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14), txtClientID.text), "0.00")
    If dblAmount < 0 Then
        balanceisNegative = True
    End If
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub cmdAddEditBankCode_Click()
'   Load frmNominalLedger
'   frmNominalLedger.CALLER_FORM = "frmClientNew4"
'   frmNominalLedger.txtClientList.Tag = txtClientID.text
'   frmNominalLedger.Show
'   'Me.Hide
'
'   Exit Sub
'
'   Load frmNominalLedger1
'   frmNominalLedger1.CALLER_FORM = "frmClientNew4"
'   frmNominalLedger1.Show
'   Me.Hide
End Sub

Private Sub cmdAddNewBD_Click()
   Load frmBank
   frmBank.CALLER_FORM = "frmClientNew4"
   frmBank.Show
   frmBank.ZOrder 0
   'FocusControl frmBank
   'Me.Hide
End Sub

Private Sub cmdAgmntAddNew_Click()
   Dim conAgr As New ADODB.Connection
   Dim rstCPA_ID As New ADODB.Recordset

   Dim szSQL As String

   'Set the RDO Connections to the dataset
'   conAgr.Open getConnectionString
'
'   szSQL = "SELECT CPA_ID " & _
'           "FROM ClientProAgr " & _
'           "WHERE " & _
'               "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
'               "ClientProAgr.PropertyID = '" & szPropertyID & "';"
'   rstCPA_ID.Open szSQL, conAgr, adOpenDynamic, adLockOptimistic
'
'   If rstCPA_ID.EOF Then
'      MsgBox "Please enter and save Global Settings.", vbCritical + vbOKOnly, "Main Agreement"
'      'cmdAgrTopEdit_Click
'      rstCPA_ID.Close
'      Set rstCPA_ID = Nothing
'      conAgr.Close
'      Set conAgr = Nothing
'      Exit Sub
'   End If
'
'   rstCPA_ID.Close
'   Set rstCPA_ID = Nothing
'   conAgr.Close
'   Set conAgr = Nothing
'Do you wish to add a new Management Fee to this agreement?
   If MsgBox("Do you wish to add a new Management Fee to this agreement?", vbYesNo, "Add New Agreement") = vbNo Then Exit Sub
   AgreementButtonMode NewEntryMode
   AgreementClearMode ClearOnlyTextBoxes
   AGREEMENT_ADDNEW_MODE = True
   FocusControl cmdCommandArray(9)
   'When you click add new you need to show last charge date for the selected property
'   Dim lastchargeDate As String
'   Dim adoconn As New ADODB.Connection
'   adoconn.Open getConnectionString
'   lastchargeDate = findLastChargeDate(szPropertySelection1, flxManagementFee.TextMatrix(flxManagementFee.row, 7), adoconn)
'   adoconn.Close
'   txtLastChargeDate.text = Format(lastchargeDate, "dd/MM/yyyy")
'   cboCHARGE_TYPE.SetFocus
End Sub

Private Sub cmdAgmntCancel_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   AgreementButtonMode DefaultMode
   AgreementClearMode ClearOnlyTextBoxes
   FocusControl cmdAgmntAddNew
End Sub

Private Sub cmdAgrTopEdit_Click()
   CPAButtonMode EditMode
'   cboProperty.Enabled = True
'   cboProperty.Locked = False
   cmdAgrTopEdit.Enabled = False
   cmdAgrTopSave.Enabled = True
   cmdCanelAgree.Enabled = True
   FocusControl txtAgreementStartDate
End Sub

Private Sub cmdAgrTopSave_Click()
   'anol remember only one property can be saved in the ClientProAgr table is saved as per rule.
   'And once you have saved there is no option for delete the record 2020-07-03
   Dim nChoice As Integer
'   If txtREVIEW_DATE.text = "" Then
'      MsgBox "Please type the rent review date.", vbCritical + vbOKOnly, "Rent Review Date"
'      txtREVIEW_DATE.SetFocus
'      Exit Sub
'   End If
   If txtAgreementStartDate.text = "" Then
      MsgBox "Please type the Agreement Start Date.", vbCritical + vbOKOnly, "Agreement Start Date"
      txtAgreementStartDate.SetFocus
      Exit Sub
   End If
'   If txtAgreementEndDate.text = "" Then
'      MsgBox "Please type the Agreement End Date.", vbCritical + vbOKOnly, "Agreement End Date"
'      txtAgreementEndDate.SetFocus
'      Exit Sub
'   End If
   If szPropertySelection1 = "" Then
      MsgBox "Please Select a property from the list.", vbCritical + vbOKOnly, "Select a property"
      FocusControl flxPropertySelection1
      Exit Sub
   End If

'   If txtNoOfDaysToSendMFB4Due.text = "" Then
'      MsgBox "Please enter the notice period in days.", vbCritical + vbOKOnly, "Notice Period"
'      txtNoOfDaysToSendMFB4Due.SetFocus
'      Exit Sub
'   End If
   
   nChoice = MsgBox("Do you wish to save this agreement?", vbQuestion + vbYesNo, "Save Agreement")
   If nChoice = vbNo Then
      CPAButtonMode DefaultMode
      CPAClearMode ClearOnlyTextBoxes
      Exit Sub
   End If
   If nChoice = vbCancel Then
      Exit Sub
   End If
   
   'MousePointer = vbHourglass

   Dim conAgr As New ADODB.Connection
   Dim rstAgr As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Open getConnectionString

   szSQL = "SELECT * " & _
        "FROM ClientProAgr " & _
        "WHERE " & _
            "ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
            "ClientProAgr.PropertyID = '" & szPropertySelection1 & "';"
   rstAgr.Open szSQL, conAgr, adOpenDynamic, adLockOptimistic

   With rstAgr
      If .EOF Then
         .AddNew
         !ClientID = txtClientID.text
         !propertyID = szPropertySelection1
      End If

      !REVIEW_DATE = txtREVIEW_DATE.text
      !agreementStartDate = txtAgreementStartDate.text
      !agreementEndDate = txtAgreementEndDate.text

      .Update
      .Close
   End With

   conAgr.Close
   Set rstAgr = Nothing
   Set conAgr = Nothing

   CPAButtonMode DefaultMode
   cmdAgrTopEdit.Enabled = True
   cmdAgrTopSave.Enabled = False
   tabAgreement.Enabled = True
'   cboProperty.Enabled = False
   'MousePointer = vbDefault

   MsgBox "Agreement has been added successfully", vbInformation, "Saved"
   FocusControl cmdAgmntAddNew
   cmdAgrTopEdit.Caption = "&Edit"
   cmdAgmntAddNew.Enabled = True
   cmdPayAddNew.Enabled = True
   
   cmdAgrTopEdit.Enabled = True
    cmdAgrTopSave.Enabled = False
    cmdAgmntAddNew.Enabled = True
   Exit Sub
ErrorHandler:
   MsgBox Err.Number & Err.description & " ", vbCritical + vbOK, "PCM Error: 125"
   MousePointer = vbDefault
End Sub

Private Sub cmdAutoSetup_Click(Index As Integer)
   Dim DTdate As Date, var

'   On Error GoTo ErrorHandler

   var = InputBox("Please type the first payment date of the year. (dd/mm/yyyy)", "Frist Payment Date", "01/01/" & Year(Date))
   If var = "" Then Exit Sub

   DTdate = Format(var, "dd mmmm yyyy")

   SetAddDates DTdate

   Exit Sub
ErrorHandler:
   If MsgBox("Please retype the date only.", vbCritical + vbRetryCancel, "Wrong Input") = vbRetry Then
      cmdAutoSetup_Click (0)
   End If
End Sub

Private Sub SetAddDates(DTdate As Date)
   cboDay(31).text = Format(DTdate, "dd")
   cboDay(32).text = Format(DTdate, "dd")
   cboDay(33).text = Format(DTdate, "dd")
   cboDay(34).text = Format(DTdate, "dd")
   cboDay(35).text = Format(DTdate, "dd")
   cboDay(36).text = Format(DTdate, "dd")
   cboDay(37).text = Format(DTdate, "dd")
   cboDay(38).text = Format(DTdate, "dd")
   cboDay(39).text = Format(DTdate, "dd")
   cboDay(40).text = Format(DTdate, "dd")
   cboDay(41).text = Format(DTdate, "dd")
   cboDay(42).text = Format(DTdate, "dd")

  

'Quarterly
   cboQDay(11).text = Format(DTdate, "dd")
   cboQDay(12).text = Format(DTdate, "dd")
   cboQDay(13).text = Format(DTdate, "dd")
   cboQDay(14).text = Format(DTdate, "dd")

   cboQMth(11).text = Format(DTdate, "mmmm")
   cboQMth(12).text = Format(DateAdd("m", 3, DTdate), "mmmm")
   cboQMth(13).text = Format(DateAdd("m", 6, DTdate), "mmmm")
   cboQMth(14).text = Format(DateAdd("m", 9, DTdate), "mmmm")
'
''Half yearly
   cboHDay(5).text = Format(DTdate, "dd")
   cboHDay(6).text = Format(DTdate, "dd")

   cboHMth(5).text = Format(DTdate, "mmmm")
   cboHMth(6).text = Format(DateAdd("m", 6, DTdate), "mmmm")
'
''Yearly
   cboYDay(2).text = Format(DTdate, "dd")
   cboYMth(2).text = Format(DTdate, "mmmm")
End Sub

Private Sub cmdBACS_Click()
    Dim X As Integer
    If flxOtherBankDetails.row = 0 Then MsgBox "Please select a bank to edit", vbInformation, "Warning"
    If txtNCCODE.text <> flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14) Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4) <> txtBank_AC_Name.text Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6) <> txtBANK_SC.text Or _
        UCase(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)) <> UCase(txtPaymentMethod.text) Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5) <> txtBANK_AC_NUM.text Then
        X = MsgBox("Do you want to save the current changes you made?", vbYesNo, "Save?")
        If X = vbYes Then
            FocusControl cmdSaveBank
            Exit Sub
        Else
            Call restoreSelectedValueflxOtherBankDetails
        End If
    End If
    Frame14.Enabled = True
    flxOtherBankDetails.Enabled = True
    bOverdraftWarning = False
    CommandButtonEnabled True
    EnableDisableAcText True
    NewBankText True, True
    flxOtherBankDetails_RowColChange
    cmdSetDefaultAC.Enabled = True
    'added by anol 11 Mar 2015
    cboBank_ID.Locked = True
    cmdSaveBank.Enabled = False
    cmdSetDefaultAC.Enabled = False
'    cmdDeleteBank.Enabled = False
    cmdBACS.Enabled = False
   
    FocusControl cmdEdit
   If iTotalBankAC = 0 Then Exit Sub      'there are no bank details has been inputed yet

   If txtBANK_NAME.text = "" Or flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 13) = "" Then
      MsgBox "Please select a bank from the list.", vbExclamation + vbOKOnly, "Warning"
      Exit Sub
   End If
   
      
      
      
  
   Me.Enabled = False
   frmConfigEB.idBank = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 13)
   Load frmConfigEB
   frmConfigEB.Show
   
End Sub

Private Sub cmdCancelBank_Click()
    If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
    Frame14.Enabled = True
    flxOtherBankDetails.Enabled = True
    bOverdraftWarning = False
    CommandButtonEnabled True
    EnableDisableAcText True
    NewBankText True, True
    flxOtherBankDetails_RowColChange
    cmdSetDefaultAC.Enabled = True
    'added by anol 11 Mar 2015
    cboBank_ID.Locked = True
    cmdSaveBank.Enabled = False
    cmdSetDefaultAC.Enabled = False
'    cmdDeleteBank.Enabled = False
    cmdBACS.Enabled = False
    cmdClient.Enabled = True
    flxBankAccountFund.Visible = False
    chkShowFundBankAccount.Value = 0
    
End Sub

Private Sub cmdCancelChange_Click()
   UnlockMainClientText False
   MainCommandButtonEnable False
   cmdEditClient.Enabled = True
   
'   ComponentInFrameClearMode Me, picMain, ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, Frame1(0), ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, Frame1(1), ClearOnlyTextBoxes
'   ComponentInFrameClearMode Me, Frame1(2), ClearOnlyTextBoxes
  ' tabMain.Tab = 0
'   cboProperty.Clear
   txtREVIEW_DATE.text = ""
   txtNoOfDaysToSendMFB4Due.text = ""
   flxACHistory.Clear
   flxACHistory.Rows = 2
   flxACHistorySplit.Clear
   flxACHistorySplit.Rows = 2
  ' NewBankText True, True
   txtBANK_NAME.Enabled = False
   txtBANK_ADDRESS1.Enabled = False
   txtBANK_ADDRESS2.Enabled = False
   txtBANK_ADDRESS3.Enabled = False
   txtBANK_POST_CODE.Enabled = False
   
'   flxOtherBankDetails.Clear
'   flxOtherBankDetails.Rows = 2
   cmdDeleteClient.Enabled = False
   
   '**control mode by anol
'    cboProperty.Enabled = False
'    txtREVIEW_DATE.Locked = True
    cmdAgrTopSave.Enabled = False
    txtNoOfDaysToSendMFB4Due.Locked = True
    cmdAgrTopSave.Enabled = False
    tabAgreement.Enabled = True
    cmdAgmntAddNew.Enabled = False
    cmdAgrTopEdit.Enabled = False
    
    cmdPayAddNew.Enabled = False
    
'    cmdAddNewBD.Enabled = False
    cmdSetDefaultAC.Enabled = False
'    cmdDeleteBank.Enabled = False
    cmdBACS.Enabled = False
    cmdAddNewBank.Enabled = False
    'cmdEditBank.Enabled = False
    cmdSaveBank.Enabled = False
    chkOverDraft.Enabled = False
    chkConsolidated.Enabled = False
'    cboDmdPropertyList.Enabled = False
'    cmdGSEdit.Enabled = False
    
    
    cmdVAMemo.Enabled = False
    cmdUnitMemoNew.Enabled = False
    cmdUnitMemoEdit.Enabled = False
    cmdUnitMemoSave.Enabled = False
    cmdDelete.Enabled = False
    cmdUnitMemoCancel.Enabled = False
    cmdClientAddAtch(0).Enabled = False
    cmdClientAddAtch(1).Enabled = False
    cmdClientAddAtch(2).Enabled = False
    cmdCloseMemo.Enabled = False
'    cmdClientDetailsEdit.Enabled = False
    cmdImgLeftMove.Enabled = False
    cmdUploadImageAdd.Enabled = False
    cmdImgDelete.Enabled = False
    cmdDeleteClient.Enabled = False
    cmdEditClient.Enabled = False
    cmdCancelBank.Enabled = False
    LockingAllTextClientAddress True
    Call LoadClientAddressLins
 '****
'    cmdClientDetailsEdit.Enabled = True
    cmdEditClient.Enabled = True
     chkOptedtoTax.Enabled = False
     cmdBrowseTemplate.Enabled = False

End Sub

Private Sub cmdClientDetailsCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllTextClientAddress True
   CommandButtonEnable True
   Call LoadClientAddressLins
End Sub

Private Sub CommandButtonEnable(bEnable As Boolean)
'   cmdClientDetailsEdit.Enabled = bEnable
   'cmdClientDetailsSave.Enabled = Not bEnable
   'cmdClientDetailsCancel.Enabled = Not bEnable
End Sub

Private Sub cmdClientDetailsEdit_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

'   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllTextClientAddress False
   txtClientAddressLine1(0).SetFocus
   CommandButtonEnable False
   cmdEditClient.Enabled = False
   cmdSaveClient.Enabled = True
   cmdCancelChange.Enabled = True
End Sub

Private Sub cmdClientDetailsSave_Click(adoConn As ADODB.Connection)
   Dim rstClient As New ADODB.Recordset
   Dim rstSupplier As New ADODB.Recordset
   Dim szSQL As String

  

   If (NEW_TYPE = "Landlord") Then
      szSQL = "SELECT * " & _
              "FROM Landlord " & _
              "WHERE LandlordID = '" & txtClientID.text & "';"
      rstClient.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With rstClient
         !LandlordAddressLine1 = txtClientAddressLine1(0).text
         !LandlordAddressLine2 = txtClientAddressLine1(1).text
         !LandlordAddressLine3 = txtClientAddressLine1(2).text
         !LandlordAddressLine4 = txtClientAddressLine1(3).text
         !LandlordPostCode = txtClientAddressLine1(5).text
         
         !LandlordOfficeEmail = txtClientHomeTel(4).text
         !LandlordPersonalEmail = txtClientHomeTel(3).text
         !LandlordHomeTel = txtClientHomeTel(0).text
         !LandlordMobile = txtClientHomeTel(2).text
         !LandlordOfficeTel = txtClientHomeTel(1).text
         
         !LandlordOfficeAddressLine1 = txtClientHomeTel(15).text
         !LandlordOfficeAddressLine2 = txtClientHomeTel(16).text
         !LandlordOfficeAddressLine3 = txtClientHomeTel(17).text
         !LandlordOfficeAddressLine4 = txtClientHomeTel(18).text
         !LandlordOfficePostCode = txtClientHomeTel(20).text
         
         !CompReg = txtClientHomeTel(21).text
         !RegAdd1 = txtClientHomeTel(22).text
         !RegAdd2 = txtClientHomeTel(23).text
         !RegAdd3 = txtClientHomeTel(24).text
         !RegAdd4 = txtClientHomeTel(25).text
         
         !RegPostCode = txtClientHomeTel(27).text
         .Update
         .Close
         
         szSQL = "SELECT * " & _
                 "FROM Tenants " & _
                 "WHERE SageAccountNumber = '" & txtClientID.text & "';"
         .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
         !HOAddressLine1 = txtClientAddressLine1(0).text
         !HOAddressLine2 = txtClientAddressLine1(1).text
         !HOAddressLine3 = txtClientAddressLine1(2).text
         !HOAddressLine4 = txtClientAddressLine1(3).text
         '!HOAddressLine4 = txtClientAddressLine5.text
         !HOPostCode = txtClientAddressLine1(5).text
         !Email1 = txtClientHomeTel(4).text
         !Email2 = txtClientHomeTel(3).text
         
         !DirectLine2 = txtClientHomeTel(0).text
         !DirectLine1 = txtClientHomeTel(2).text
         
         !BillAddressLine1 = txtClientHomeTel(15).text
         !BillAddressLine2 = txtClientHomeTel(16).text
         !BillAddressLine3 = txtClientHomeTel(17).text
         !BillAddressLine4 = txtClientHomeTel(18).text
         !BillPostCode = txtClientHomeTel(20).text
         !BillTelephone = txtClientHomeTel(1).text
'         If optClientAdd.Value Then
'            !InvoiceTo = "H"
'         Else
'            !InvoiceTo = "B"
'         End If
         
         If chkClientAddress.Value Then
            !StToClientAddress = 1
         Else
            !StToClientAddress = 0
         End If
         If chkStatementAddress.Value Then
            !StToStatementAddress = 1
         Else
            !StToStatementAddress = 0
         End If
         .Update
         .Close
      End With
   Else
      szSQL = "SELECT * " & _
              "FROM Client " & _
              "WHERE ClientID = '" & txtClientID.text & "';"
      rstClient.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      With rstClient
         !ClientAddressLine1 = txtClientAddressLine1(0).text
         !ClientAddressLine2 = txtClientAddressLine1(1).text
         !ClientAddressLine3 = txtClientAddressLine1(2).text
         !ClientAddressLine4 = txtClientAddressLine1(3).text
         !ClientAddressLine5 = txtClientAddressLine1(4).text
         !ConsolidatedStatement = CInt(chkConsolidatedStatement.Value)
         !StToClientAddress = CInt(chkClientAddress.Value)
         !StToStatementAddress = CInt(chkStatementAddress.Value)
         !ClientPostCode = txtClientAddressLine1(5).text
         
        
        !ClientHomeTel = txtClientHomeTel(0).text
        !ClientOfficeTel = txtClientHomeTel(1).text
        !ClientMobile = txtClientHomeTel(2).text
        !ClientPersonalEmail = txtClientHomeTel(3).text
        !ClientOfficeEmail = txtClientHomeTel(4).text
        !groupCode = txtClientHomeTel(5).text
        
        !StClientHomeTel = txtClientHomeTel(6).text
        !StClientOfficeTel = txtClientHomeTel(7).text
        !StClientMobile = txtClientHomeTel(8).text
        !StClientPersonalEmail = txtClientHomeTel(9).text
        !StClientOfficeEmail = txtClientHomeTel(10).text
        
      
        
         
         
         !ClientOfficeAddressLine1 = txtClientHomeTel(15).text
         !ClientOfficeAddressLine2 = txtClientHomeTel(16).text
         !ClientOfficeAddressLine3 = txtClientHomeTel(17).text
         !ClientOfficeAddressLine4 = txtClientHomeTel(18).text
         !ClientOfficeAddressLine5 = txtClientHomeTel(19).text
         !ClientOfficePostCode = txtClientHomeTel(20).text
        
         !CompReg = txtClientHomeTel(21).text
         !RegAdd1 = txtClientHomeTel(22).text
         !RegAdd2 = txtClientHomeTel(23).text
         !RegAdd3 = txtClientHomeTel(24).text
         !RegAdd4 = txtClientHomeTel(25).text
         !RegAdd5 = txtClientHomeTel(26).text
         !RegPostCode = txtClientHomeTel(27).text
         .Update
         .Close
      End With
       szSQL = "SELECT * " & _
              "FROM Supplier " & _
              "WHERE SupplierID = '" & txtClientID.text & "';"
      rstSupplier.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
     If Not rstSupplier.EOF Then
              With rstSupplier
                 !SupplierAddressLine1 = txtClientAddressLine1(0).text
                 !SupplierAddressLine2 = txtClientAddressLine1(1).text
                 !SupplierAddressLine3 = txtClientAddressLine1(2).text
                 !SupplierAddressLine4 = txtClientAddressLine1(3).text
                 !SupplierPostCode = txtClientAddressLine1(5).text
                 !SupplierPersonalEmail = txtClientHomeTel(3).text
                 !SupplierOfficeAddressLine1 = txtClientHomeTel(15).text
                 !SupplierOfficeAddressLine2 = txtClientHomeTel(16).text
                 !SupplierOfficeAddressLine3 = txtClientHomeTel(17).text
                 !SupplierOfficeAddressLine4 = txtClientHomeTel(18).text
                 !SupplierOfficePostCode = txtClientHomeTel(20).text
                 !SupplierOfficeEmail = txtClientHomeTel(1).text
                 !VatCode = txtAcBalance(1).Tag
                 !optedTotax = IIf(chkOptedtoTax.Value = 1, True, False)
                 .Update
                 .Close
              End With
      End If
   End If

   
   Set rstClient = Nothing
   LockingAllTextClientAddress True
   'ShowMsgInTaskBar "Data has been updated successfully", "Y", "P"
   CommandButtonEnable True
End Sub



Private Function CheckBankInDemandRecord(szBankMY_ID As String) As Boolean
   If szBankMY_ID = "" Then Exit Function

   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT COUNT(DemandID) " & _
           "FROM DemandRecords " & _
           "WHERE DemandRecords.spare1 = '" & szBankMY_ID & "' AND " & _
               "(DemandRecords.IsPrinted = FALSE OR " & _
               "DemandRecords.UPDATE_SAGE = FALSE) AND " & _
               "DemandRecords.DemandHistory = FALSE;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Val(adoRST.Fields.Item(0).Value) > 0 Then
      CheckBankInDemandRecord = True
   Else
      CheckBankInDemandRecord = False

      adoConn.Execute "DELETE * FROM tlbClientBanks WHERE MY_ID = " & szBankMY_ID & ";"
   End If
'   If adoRst.EOF Then
'      CheckBankInDemandRecord = True
'   Else
'      CheckBankInDemandRecord = False
'   End If
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function

Private Sub cmdCTSec_Click()
   lstCT.Top = txtCT.Top
   lstCT.Left = txtCT.Left
   lstCT.Width = txtCT.Width
   lstCT.Visible = True
   lstCT.ZOrder 0
   lstCT.SetFocus
   lstCT.ListIndex = 0
End Sub

Private Sub cmdDeleteBank_Click()
    If cmdSaveClient.Enabled = True Then
        MsgBox "Please save the header section first before proceeding with update details", vbInformation, "Warning"
        FocusControl cmdSaveClient
        Exit Sub
   End If
   If flxOtherBankDetails.row = 0 Then
        MsgBox "Please Select a bank account for delete", vbInformation, "Warning"
        Exit Sub
   End If
   If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.RowSel, 7) = "YES" Then
      MsgBox "        You cannot delete a default bank account." & Chr(10) & _
             "Please remove default status to delete this bank account.", vbInformation + vbOKOnly, "Records Deleting Error"
      Exit Sub
   End If
   If flxOtherBankDetails.TextMatrix(1, 1) = "" Then Exit Sub
   If MsgBox("Do you want to delete current account details?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub

   If CheckBankInDemandRecord(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.RowSel, 13)) Then
      MsgBox "You cannot delete this bank account record " & Chr(10) & _
             "because you have unprinted or unposted demands," & Chr(10) & _
             "relating to this bank account record.", vbCritical + vbOKOnly, "Cannot delete bank details"
      Exit Sub
   End If

'   flxOtherBankDetails.RemoveItem (flxOtherBankDetails.row)
'
'   flxOtherBankDetails_RowColChange
    ConfigFlxOtherBankDetails
    LoadFlxOtherBankDetails
   NewBankText True, False
   cmdDeleteBank.Enabled = True
   EnableDisableAcText True
    'Frame14.Enabled = False
   MsgBox "Bank account successfully deleted.", vbInformation + vbOKOnly, "Bank Account Deleted"
End Sub

'Private Sub cmdClientAddAtch(2)_Click()
'   If cmbFiles.text = "" Then Exit Sub
'   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
'   If (NEW_TYPE = "Landlord") Then
'      DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtClientID.text, "Landlord"
'   Else
'      DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtClientID.text, "Client"
'   End If
'   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
'End Sub

Private Sub cmdEditBank_Click()
   If cmdSaveClient.Enabled = True Then
        MsgBox "Please save the header section first before proceeding with update details", vbInformation, "Warning"
        FocusControl cmdSaveClient
        Exit Sub
   End If
   PopulateBank
   If iTotalBankAC = 0 Then Exit Sub      'there are no bank details has been inputed yet
'   If txtBANK_NAME.text = "" Then
'      MsgBox "Please select a bank from the list.", vbExclamation + vbOKOnly, "Warning"
'      Exit Sub
'   End If

   bBankNewEdit = False
   bOverdraftWarning = True
   EnableDisableAcText False

   CommandButtonEnabled False
   flxOtherBankDetails.Enabled = True
   cmdAddEditBankCode.Enabled = True
'   flxOtherBankDetails.row = flxOtherBankDetails.Rows - 1
   'Added by anol 11 Mar 2015
   cboBank_ID.Locked = False
   cmdSetDefaultAC.Enabled = True
   cmdBACS.Enabled = True
End Sub

Private Sub cmdAddNewBank_Click()
    If cmdSaveClient.Enabled = True Then
        MsgBox "Please save the header section first before proceeding with update details", vbInformation, "Warning"
        FocusControl cmdSaveClient
        Exit Sub
   End If
   If txtClientID.text = "" Then Exit Sub
   bOverdraftWarning = False
   If iTotalBankAC = 0 Then
      bDefaultAccount = True
   Else
      bDefaultAccount = False
   End If

   bBankNewEdit = True
   PopulateBank

   'cboBank_ID.SetFocus
    chkOverDraft.Value = False
    txtOverDraft.text = "0.00"
    chkConsolidated.Value = False

   EnableDisableAcText False
   FocusControl cboBank_ID
   NewBankText True, True
   cboBank_ID.Locked = False
   
   CommandButtonEnabled False
   cmdAddEditBankCode.Enabled = True
   flxOtherBankDetails.row = flxOtherBankDetails.Rows - 1
   cmdEdit.Enabled = False
   Frame14.Enabled = False
End Sub

Private Sub CommandButtonEnabled(bEnable As Boolean)
   cmdAddNewBank.Enabled = bEnable
   'cmdEditBank.Enabled = bEnable
   cmdDeleteBank.Enabled = bEnable
   cmdSaveBank.Enabled = Not bEnable
   cmdCancelBank.Enabled = Not bEnable
  ' cmdCommandArray(13).Enabled = Not bEnable
'   flxOtherBankDetails.Enabled = bEnable

   cmdBACS.Enabled = bEnable
End Sub

Public Function PopulateBank()
   Dim sSQLQuery_ As String

   adoBank.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT BANK_ID, BANK_NAME, SORT_CODE, " & _
                     "BANK_ADDRESS1, BANK_POST_CODE, " & _
                     "BANK_ADDRESS2, BANK_ADDRESS3, BANK_ADDRESS4 " & _
                "FROM tlbBank"

   adoBank.RecordSource = sSQLQuery_
   adoBank.CommandType = adCmdText
   adoBank.Refresh

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoBank.Recordset.RecordCount
   TotalCol = adoBank.Recordset.Fields.Count

   Dim Data() As String, i As Integer, j As Integer

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To adoBank.Recordset.RecordCount - 1
       For j = 0 To adoBank.Recordset.Fields.Count - 1
           Data(j, i) = IIf(IsNull(adoBank.Recordset.Fields(j)), "", adoBank.Recordset.Fields(j))
       Next j
       adoBank.Recordset.MoveNext
   Next i

   cboBank_ID.Column() = Data()
End Function

Public Sub RefreshedBankDetails()
   cboBank_ID_Click
   LoadFlxOtherBankDetails
End Sub

Private Sub cmdAddNewClient_Click()
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

   NEW_TYPE = "Client"
'
'   If optClient.Value Then
'      NEW_TYPE = optClient.Caption
'   Else
'      NEW_TYPE = optLandlord.Caption
'   End If

   bNewEdit = True

  ' MousePointer = vbHourglass

   ADD_NEW_CLIENT = True
   UnlockMainClientText True
   MainCommandButtonEnable True
   ComponentInFrameClearMode Me, picMain, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, Frame1(0), ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, Frame1(1), ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, Frame1(2), ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fraType, ClearOnlyTextBoxes
   ComponentInFrameClearMode Me, fraOccupied, ClearOnlyTextBoxes
   txtYearEndDate.Enabled = False
   If UBound(szaChoice) > 0 Then
      txtClientID.Locked = False
      If szaChoice(1) <> "" Then
         If InStr(szaChoice(1), "CL") > 0 Then
               txtClientName.SetFocus
         End If
      Else
         txtClientID.Enabled = True
         txtClientID.SetFocus
      End If
   End If

   tvwLandLord.Nodes.Clear
   chkOptedtoTax.Value = 0
   txtAcBalance(1).text = ""
   txtAcBalance(1).Tag = ""
   chkOptedtoTax.Enabled = True
   cmdBrowseTemplate.Enabled = True
End Sub
'
'Private Sub SageCustomerAccCombo()
'   On Error GoTo Error_Handler
'
'   ' Declare Objects
'   Dim oSDO As SageDataObject120.SDOEngine
'   Dim oWS As SageDataObject120.Workspace
'   Dim oSalesRecord As SageDataObject120.SalesRecord
'
'   ' Declare Variables
'   Dim szDataPath As String
'
'   ' Create the SDOEngine Object
'   Set oSDO = New SageDataObject120.SDOEngine
'
'   ' Create the Workspace
'   Set oWS = oSDO.Workspaces.Add("Prestige")
'
'   'read datapath from registr
'   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
'   If szDataPath = "" Then
'      ' Select Company. The SelectCompany method takes the program install
'      ' folder as a parameter
'      szDataPath = oSDO.SelectCompany(sageDirPath)
'      'Save company name in the registry
'      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
'   End If
'   ' Try to Connect - Will Throw an Exception if it Fails
'   If oWS.Connect(szDataPath, sageUserName, sagePassword, "Prestige") Then
'
'      Set oSalesRecord = oWS.CreateObject("SalesRecord")
'
'      Dim TotalRow, TotalCol As Long
'      Dim Data() As String
'      Dim i As Integer
'
'      TotalRow = oSalesRecord.Count
'      TotalCol = 2
'      cboLandLordSageCustAC.Clear
'
'      ReDim Data(TotalCol, TotalRow) As String
'
'      oSalesRecord.MoveFirst
'      For i = 0 To TotalRow - 1
'         Data(0, i) = CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value)
'         Data(1, i) = CStr(oSalesRecord.Fields.Item("NAME").Value)
'         oSalesRecord.MoveNext
'      Next i
'
'      cboLandLordSageCustAC.Column() = Data()
'      cboLandLordSageCustAC.ColumnCount = TotalCol
'      cboLandLordSageCustAC.BoundColumn = 1
'
'      'Disconnect
'      oWS.Disconnect
'   End If
'
'   ' Destroy Objects
'   Set oSalesRecord = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   MsgBox "(pcm_002) The SDO generated the following error: " & oSDO.LastError.text
'   Set oSalesRecord = Nothing
'   Set oWS = Nothing
'   Set oSDO = Nothing
'End Sub

Private Sub UnlockMainClientText(bUnlock As Boolean)
   txtClientName.Locked = Not bUnlock
   txtVATReg.Locked = Not bUnlock
   'txtYearEndDate.Locked = Not bUnlock
   'cmdResidency.Enabled = bUnlock
   cmdCTSec.Enabled = bUnlock
End Sub

Private Sub cmdAgmntEdit_Click()
   If MsgBox("Do you want to edit the Management Fee?", vbQuestion + vbYesNo, "Edit Management Fee") = vbNo Then Exit Sub

   AgreementButtonMode EditMode 'Management Fee
   AGREEMENT_ADDNEW_MODE = False 'Means edit mode on
   If txtLastChargeDate.text = "" Then
        txtLastChargeDate.Locked = False
   End If
   flxManagementFee_displaytextboxes
   FocusControl cmdCommandArray(9)
   cmdAgmntEdit.Enabled = False
   cmdDeleteMgtFee.Enabled = False
   flxManagementFee.Enabled = False
End Sub

'Private Function findLastChargeDate(strPropertyID As String)
'    Dim adoConn As New adodb.Connection
'    Dim rsChargedate As New adodb.Recordset
'    adoConn.Open getConnectionString
'    rsChargedate.Open "Select max(R.ChargeDate) as chrgDate from tlbReceiptSplit S,tlbreceipt R,Units U where U.UnitNumber=R.UnitID AND U.PropertyID='" & strPropertyID & "'", adoConn, adOpenStatic, adLockReadOnly
'    If Not rsChargedate.EOF Then
'        findLastChargeDate = IIf(IsNull(rsChargedate("chrgDate").Value) = True, "", rsChargedate("chrgDate").Value)
'    End If
'    rsChargedate.Close
'    Set rsChargedate = Nothing
'    adoConn.Close
'    Set adoConn = Nothing
'
'End Function
Private Sub cmdAgmntSave_Click()
'tlbAgrement table is the is detail table tlbClientProagr table , also there is another table for detail is the tlbpayables which is of dirrerent type.
   'On Error GoTo ErrorHandler
    If txtChargeType.Tag = "" Then
        MsgBox "Please enter Charge type", vbInformation, "Charge type"
        cmdCommandArray(9).SetFocus
        Exit Sub
    End If
    If txtChargingMethod.Tag = "" Then
        MsgBox "Please enter Charging Method", vbInformation, "Charging Method"
        cmdCommandArray(10).SetFocus
        Exit Sub
    End If
    If txtFundMngtFee.text = "" Then
        MsgBox "Please enter fund", vbInformation, "Fund!!"
        cmdCommandArray(7).SetFocus
        Exit Sub
    End If
    If txtManagingAgentAC.text = "" Then
        MsgBox "Please enter Managing Agent", vbInformation, "Managing Agent empty!"
        cmdCommandArray(6).SetFocus
        Exit Sub
    End If
    If txtSTART_DATE.text = "" Then
        MsgBox "Please enter start Date", vbInformation, "Start date"
        txtSTART_DATE.SetFocus
        Exit Sub
    End If
    If txtNtDueDate.text = "" Then
        MsgBox "Please enter next due date", vbInformation, "Due date"
        txtNtDueDate.SetFocus
        Exit Sub
    End If
    If txtFrequecymngtFee.text = "" Then
        MsgBox "Please enter frequency", vbInformation, "Frequency"
        cmdCommandArray(12).SetFocus
        Exit Sub
    End If
     If txtChargeBasis.text = "" Then
        MsgBox "Please enter a Charge Basis", vbInformation, " Charge Basis"
        FocusControl txtChargeBasis
        Exit Sub
     End If
   Dim nChoice As Integer
   Dim DicFundID As New Dictionary
   Dim DicFundIDOnceMore As New Dictionary
   Dim iRow As Integer
   DicFundID.removeAll
   DicFundIDOnceMore.removeAll
   For iRow = 1 To flxManagementFee.Rows - 1
        If UCase(flxManagementFee.TextMatrix(iRow, 10)) = UCase("RE_ED") Then
            If Not DicFundID.Exists(flxManagementFee.TextMatrix(iRow, 7)) Then
                    DicFundID.Add flxManagementFee.TextMatrix(iRow, 7), flxManagementFee.TextMatrix(iRow, 8)
'            Else
'                    MsgBox "This fund already exits in the agreement for Received Basis. Please add a different fund."
'                    Exit Sub
            End If
        End If
   Next
   
'   If DicFundID.Exists((txtFundMngtFee.Tag)) And txtChargingMethod.Tag = UCase("RE_ED") Then
'            MsgBox "This fund already exits in the agreement for Received Basis. Please add a different fund."
'            Exit Sub
'   End If
   
   '*******************************************************************
   'code added by mahboob 01/08/2023
   If AGREEMENT_ADDNEW_MODE Then
            If DicFundID.Exists((txtFundMngtFee.Tag)) And txtChargingMethod.Tag = UCase("RE_ED") Then
                     MsgBox "This fund already exits in the agreement for Received Basis. Please add a different fund."
                     DicFundID.removeAll
                     Exit Sub
            End If
   End If
    '*******************************************************************
    
     '*******************************************************************
   If Not AGREEMENT_ADDNEW_MODE Then    'When you are on edit mode
   'code added by anol 02/08/2023
        DicFundIDOnceMore.Add txtFundMngtFee.Tag, txtFundMngtFee.text ' take in consideration what you are going to add
            For iRow = 1 To flxManagementFee.Rows - 1
                'If UCase(flxManagementFee.TextMatrix(iRow, 10)) = UCase("RE_ED") Then
                  If iRow <> rRow Then 'when you are editing need not to consider what is currently selected
                    If DicFundIDOnceMore.Exists(flxManagementFee.TextMatrix(iRow, 7)) Then
                            MsgBox "This fund already exits in the agreement for Received Basis. Please add a different fund."
                            DicFundIDOnceMore.removeAll
                             Exit Sub
                    Else
                           DicFundIDOnceMore.Add flxManagementFee.TextMatrix(iRow, 7), flxManagementFee.TextMatrix(iRow, 8)
        
                    End If
                End If
           Next
    End If
    
   
    '*******************************************************************
    
   Dim lastchargeDate As String
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   'lastchargeDate = findLastChargeDate(szPropertySelection1, adoconn)
   adoConn.Close
   
   'we are not dealing with last charge date from this setupscreen this will be handled when we are charging it.2023-02-17
'   If IsDate(txtLastChargeDate.text) = False And lastchargeDate = "" Then
'        MsgBox "Please enter last charge date", vbInformation, "Warning"
'        FocusControl txtLastChargeDate
'        txtLastChargeDate.Locked = False
'        Exit Sub
'   End If
   
   
'Here you shall put a validation so that user cannot enter duplicate charge type for  a single property


        nChoice = MsgBox("Do you wish to save this Management Fee Charge?", vbQuestion + vbYesNo, "Please confirm")
        If nChoice = vbNo Then
            AgreementButtonMode DefaultMode
            AgreementClearMode ClearOnlyTextBoxes
            Exit Sub
        ElseIf nChoice = vbYes Then
              AgreementButtonMode NewEntryMode
              flxManagementFee.Enabled = True
            '        AGREEMENT_ADDNEW_MODE = True
        End If
        If nChoice = vbCancel Then Exit Sub
        Dim conAgr As New ADODB.Connection
        Dim rstAgr As New ADODB.Recordset, rstCPA_ID As New ADODB.Recordset
        Dim szSQL As String
        Dim strCPA_ID As Integer
        'On Error GoTo ErrorHandler
        conAgr.Open getConnectionString
        conAgr.BeginTrans
      
      'rstAgr.AddNew
     
          szSQL = "SELECT CPA_ID FROM ClientProAgr WHERE ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
                      "ClientProAgr.PropertyID = '" & szPropertySelection1 & "';"
          rstCPA_ID.Open szSQL, conAgr, adOpenStatic, adLockReadOnly
    
          strCPA_ID = rstCPA_ID!CPA_ID
          rstCPA_ID.Close
          Set rstCPA_ID = Nothing
        If AGREEMENT_ADDNEW_MODE Then
           
        Else
             strtlbAgreementID = flxManagementFee.TextMatrix(flxManagementFee.row, 1)
        End If
        'tlbAgreement is for management fee details
        If AGREEMENT_ADDNEW_MODE Then
             szSQL = "SELECT * FROM tlbAgreement;"
        Else
              szSQL = "SELECT * FROM tlbAgreement WHERE AGREEMENT_ID = " & strtlbAgreementID & " AND CPA_ID = " & strCPA_ID & ";"
        End If
        rstAgr.Open szSQL, conAgr, adOpenDynamic, adLockOptimistic
        
        
        
        With rstAgr
        If AGREEMENT_ADDNEW_MODE Then
                .AddNew
               '!AGREEMENT_ID = ""' this will be auto increamented,when you edit AGREEMENT_ID is unchanged
                !CPA_ID = strCPA_ID
        End If
        !Charge_Type = txtChargeType.Tag
        'DEMAND_TYPE Column in table tlbAgreement has been dropped after new requirement in 2022-11-28
'        !DEMAND_TYPE = txtDemandTypemngtFee.Tag
        !Fund = txtFundMngtFee.Tag
        '      !Handling = cboHandling.Value
        !CHARGE_METHOD = txtChargingMethod.Tag
        !CHARGE_BASIS = txtChargeBasis.Tag
        !AnnualCharge = Val(txtTotalAmountPerYear.text)
        If IsDate(txtSTART_DATE.text) Then
            !START_DATE = Format(txtSTART_DATE.text, "dd/mmmm/yyyy")
        End If
         If Trim(txtEND_DATE.text) = "" Then
            !END_DATE = ""
         Else
            !END_DATE = Format(txtEND_DATE.text, "dd/mmmm/yyyy")
         End If
'        !Frequency = txtFrequecymngtFee.Tag
'****************************************************************************
         'Code changed by Mahboob 01/08/2023
'          !Frequency = txtFrequecymngtFee.Tag
         
         '****************************************************************************
         'Code changed by anol 22/08/2023
'          !Frequency = txtFrequecymngtFee.Tag
        !Frequency = IIf(txtFrequecymngtFee.text = "N/A", Null, txtFrequecymngtFee.Tag)
        If IsDate(txtNtDueDate.text) Then
            !NtDueDate = Format(txtNtDueDate.text, "dd/mmmm/yyyy")
        End If
        !ManagingAgentID = txtManagingAgentAC.Tag
        
        !amount = Val(txtAmount.text)
        !TotalAmount = Val(txtTotalAmountPerYear.text)
        !EachPeriod = Val(txtPeriod.text)
        If Trim(txtStopDatemngtFee.text) = "" Then
            !StopDate = ""
        Else
            !StopDate = txtStopDatemngtFee.text
        End If
        'we are not dealing with last charge date from this setupscreen this will be handled when we are charging it.2023-02-17
'        If lastchargeDate = "" Then
'             !lastchargeDate = Format(txtLastChargeDate.text, "dd/MMM/yyyy")
'        Else
'            !lastchargeDate = Format(lastchargeDate, "dd/MMM/yyyy")
'        End If
'        If lastchargeDate = "" Then
'             !lastchargeDate = txtLastChargeDate.text
'        Else
'            !lastchargeDate = lastchargeDate
'        End If
        If IsDate(txtNtDueDate) Then
            txtComparenextDueDate1 = DateAdd("d", 1, txtNtDueDate)
            dtFDD = NextDueDate1(CInt(txtFrequecymngtFee.Tag), txtComparenextDueDate1, szPropertySelection1)
            !FDD = dtFDD
        End If
'        !FDD = FindNextDueDate(dtNtDueDate, _
'                                                adoRstRC!BRfrequency, _
'                                                adoRstRC!BRDemandType, _
'                                                adoRstRC!propertyID, adoConn) ' Calculation of FDD must be done from startdate, Frquency and next due date, it should be equalt to next due date make some RND
        !CapAmount = Val(txtCapAmount.text)
        .Update
      .Close
   End With
   Dim rsCheck As New ADODB.Recordset
   rsCheck.Open "Select C.PropertyID as T,R.PropertyID as N from ClientProAgr C,tlbAgreement A,ChargeTypes R where A.CPA_ID=C.CPA_ID " & _
                "AND R.ID=A.CHARGE_TYPE AND C.PropertyID<>R.PropertyID", conAgr, adOpenStatic, adLockReadOnly
                If Not rsCheck.EOF Then
                    rsCheck.Close
                    conAgr.RollbackTrans
                    conAgr.Close
                    MsgBox "Could not save the record, charge type Id is not correct", vbInformation, "warning"
                    Exit Sub
                Else
                    conAgr.CommitTrans
                End If
   conAgr.Close
   Set rstAgr = Nothing
   Set conAgr = Nothing

   Call LoadflxAgreement(szPropertySelection1)        'refresh the grid

   'MousePointer = vbDefault

'   cboProperty.Locked = False
   MsgBox "Agreement successfully updated.", vbInformation + vbOKOnly, "Agreement"

   AgreementButtonMode DefaultMode
   AgreementClearMode ClearOnlyTextBoxes
   Exit Sub

ErrorHandler:

'   rstAgr.Close
'   Set rstAgr = Nothing
'   conAgr.Close
   Set conAgr = Nothing

   MsgBox Err.Number & Err.description & " " & "Do not leave any field blank.", vbCritical + vbOK, "PCM Error: 125"
   AgreementButtonMode DefaultMode
   MousePointer = vbDefault
End Sub

Private Sub cmdPaySave_Click()

    If txtPayableType.text = "" Then
        MsgBox "Please enter Payable type", vbInformation, "Payable type"
        cmdCommandArray(0).SetFocus
        Exit Sub
    End If
'    If txtPayDemandType.text = "" Then
'        MsgBox "Please enter Demand type", vbInformation, "Demand type"
'        cmdCommandArray(1).SetFocus
'        Exit Sub
'    End If
    If txtPayFund.text = "" Then
        MsgBox "Please enter fund", vbInformation, "Fund!!"
        cmdCommandArray(2).SetFocus
        Exit Sub
    End If
    If txtClientLandlord.text = "" Then
        MsgBox "Please enter Client/Landlord", vbInformation, "Client/Landlord empty!"
        cmdCommandArray(3).SetFocus
        Exit Sub
    End If
'    If txtPAY_START_DATE.text = "" Then
'        MsgBox "Please enter start Date", vbInformation, "Start date"
'        txtPAY_START_DATE.SetFocus
'        Exit Sub
'    End If
'    If txtPAY_NtDueDate.text = "" Then
'        MsgBox "Please enter next due date", vbInformation, "Due date"
'        txtPAY_NtDueDate.SetFocus
'        Exit Sub
'    End If
'    If txtPayFrequency.text = "" Then
'        MsgBox "Please enter frequency", vbInformation, "Frequency"
'        cmdCommandArray(4).SetFocus
'        Exit Sub
'    End If
      


'   If cboPAYABLE_METHOD.text = "FIXED" And txtPAY_AnnualCharge.text = "" Then
'      MsgBox "Please enter the amount.", vbCritical + vbOKOnly, "Saving Data"
'      txtPAY_AnnualCharge.SetFocus
'   End If

   Dim nChoice As Integer

   nChoice = MsgBox("Do you wish to save this rent payable entry?", vbQuestion + vbYesNo, "Please confirm")
   If nChoice = vbNo Then
      PayableButtonMode DefaultMode
      PayableClearMode ClearOnlyTextBoxes
      Exit Sub
   End If
   If nChoice = vbCancel Then
      Exit Sub
   End If

  ' MousePointer = vbHourglass

   Dim conPay As New ADODB.Connection
   Dim rstPay As New ADODB.Recordset, rstCPA_ID As New ADODB.Recordset
   Dim szSQL As String
   Dim strCPAID As String

   'On Error GoTo ErrorHandler


   conPay.Open getConnectionString
      szSQL = "SELECT CPA_ID FROM ClientProAgr " & _
              "WHERE ClientProAgr.ClientID = '" & txtClientID.text & "' And " & _
                  "ClientProAgr.PropertyID = '" & szPropertySelection1 & "';"
      rstCPA_ID.Open szSQL, conPay, adOpenStatic, adLockReadOnly

      strCPAID = rstCPA_ID!CPA_ID

      rstCPA_ID.Close
      Set rstCPA_ID = Nothing

      szSQL = "SELECT * FROM tlbPayable WHERE CPA_ID = " & strCPAID & " AND PAYABLE_ID=" & strtlbPayableID & ";"
      rstPay.Open szSQL, conPay, adOpenDynamic, adLockOptimistic
      With rstPay
          If PAYABLE_ADDNEW_MODE Then
                .AddNew
                '!PAYABLE_ID = "" ' This Id shall be auto increamented while adding new
                !CPA_ID = strCPAID 'strtlbPayableID
          End If
          !PAYABLE_TYPE = txtPayableType.Tag
'          !PAY_DEMAND_TYPE = txtPayDemandType.Tag
          !PAY_FUND = txtPayFund.Tag
          '!PAY_HANDLING = cboPAY_HANDLING.Value
          !PAYABLE_METHOD = "RECEIVED" 'cboPAYABLE_METHOD.text    'Rent Payable will only be calculated on a received basis as per spec issue 910
    '      !PAY_AnnualCharge = CCur(IIf(txtPAY_AnnualCharge.text = "", 0, txtPAY_AnnualCharge.text))
          '!PAY_START_DATE = Format(txtPAY_START_DATE.text, "dd/mmmm/yyyy")
'          If txtPAY_END_DATE.text <> "" Then
'                !PAY_END_DATE = Format(txtPAY_END_DATE.text, "dd/mmmm/yyyy")
'          End If
          '!PAY_FREQUENCY = txtPayFrequency.Tag
          !PayeeType = txtPayeeType.text
          !clientLandlordID = txtClientLandlord.text
          !PAYABLE_BASIS_ = txtPayableBasis.Tag
          '!AmountOrPercentage = X
          '!ONDD = IIf(chkONDD.Value = 1, True, False)
          If txtPayableBasis.Tag = "FA" Then
                !Percentage = 100
          Else 'total amount
                !Percentage = Val(txtPercentage.text)
          End If
          '!PAY_NtDueDate = Format(txtPAY_NtDueDate.text, "dd/mmmm/yyyy")
          If Trim(txtStopDate.text) <> "" Then
                !StopDate = Format(txtStopDate.text, "dd/mmmm/yyyy")
          End If
          !ClientID = txtClientID.text
          .Update
          .Close
       End With
    
       conPay.Close
       Set rstPay = Nothing
       Set conPay = Nothing

       Call LoadflxAgreement(szPropertySelection1)        'refresh the grid
       Dim strtxtPayableType As String
        strtxtPayableType = txtPayableType.text
        txtPayableType.Tag = ""
        txtPayableType.text = ""
'        txtPayDemandType.Tag = ""
'        txtPayDemandType.text = ""
        txtPayFund.Tag = ""
        txtPayFund.text = ""
        txtClientLandlord.text = ""
'        txtPAY_START_DATE.text = ""
'        txtPayFrequency.Tag = ""
'        txtPayFrequency.text = ""
'        txtPAY_NtDueDate.text = ""
        txtPayableBasis.text = ""
        txtPercentage.text = ""
        txtStopDate.text = ""
        txtPayeeType.text = ""
        txtPayeeType.Tag = ""
'        txtPAY_END_DATE.text = ""
        
        MsgBox strtxtPayableType & " has been successfully saved.", vbInformation + vbOKOnly, "Payable type"

        PayableButtonMode DefaultMode
        cmdClose3.Enabled = True
        FocusControl cmdClose3
        
        Exit Sub

ErrorHandler:

   rstPay.Close
   Set rstPay = Nothing
   conPay.Close
   Set conPay = Nothing

   MsgBox Err.Number & Err.description & " " & "Do not leave any fields empty.", vbCritical + vbOK, "PCM Error: 125"
   PayableButtonMode DefaultMode
  ' MousePointer = vbDefault
End Sub

Private Sub cmdClient_Click()
    lblClientID(2).Visible = True
    TextBox1.Visible = True
    cmdClient.Enabled = False
    cmdAddNewClient.Enabled = False
    Dim conClient As New ADODB.Connection
    conClient.Open getConnectionString
    txtSearchClientID.text = ""
    tabMain.Enabled = False
    picMain.Enabled = False
'    Picture1.Visible = True
'    Picture1.Refresh
    strCommandSource = "Client"
    Frame5.Top = picMain.Top + txtClientID.Top + txtClientID.Height + 5
    Frame5.Left = picMain.Left + txtClientID.Left + 5
    Frame5.Visible = True
   
    Frame5.ZOrder 0
    
'    FlxDemandsConfigure flxClientList 'configure flxclientlist
    
    ClientAccountBalance conClient  'Load all curent balance for the client
    LoadAllClientFlxGrd conClient   'LoadflxClientlist
'    Picture1.Visible = False
    conClient.Close
    Set conClient = Nothing
   
    
    txtClientID.Locked = True
   ' txtSearchClientID.SetFocus
    FocusControl txtSearchClientID
    cmdClient.Enabled = True
    cmdAddNewClient.Enabled = True
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDeleteClient_Click()
'===========================================================================================
'This button should be visible, because user should not get facility to delete any record.
'we should give user a facility to see or remove the recode from the current list.
'===========================================================================================
   If txtClientID.text = "" Then
      MsgBox "Please select a client to delete.", vbInformation, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   Dim conClient As New ADODB.Connection, conLandlord As New ADODB.Connection
   Dim rstClient As New ADODB.Recordset, rstLandlord As New ADODB.Recordset
   Dim szSQL As String

    conClient.Open getConnectionString
    szSQL = "SELECT * " & _
         "FROM Property " & _
         "WHERE ClientID = '" & txtClientID.text & "';"
    rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly
    
    If Not rstClient.EOF Then
        MsgBox "This client could not be deleted. This client has property in the database." & _
              (Chr$(13) + Chr$(10)) & "To delete this client delete all properties of this client.", vbCritical + vbOKOnly, "Delete Not Possible"
        rstClient.Close
        Set rstClient = Nothing
        conClient.Close
        Set conClient = Nothing
        Exit Sub
    End If
    
    rstClient.Close
    Set rstClient = Nothing
    szSQL = "SELECT * " & _
         "FROM FundMatrix " & _
         "WHERE ClientID = '" & txtClientID.text & "';"
    rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly
    If Not rstClient.EOF Then
        MsgBox "This client could not be deleted. There is fund assignment in FundMatrix." & _
              (Chr$(13) + Chr$(10)) & "To delete this client please delete fund assignment.", vbCritical + vbOKOnly, "Delete Not Possible"
        rstClient.Close
        Set rstClient = Nothing
        conClient.Close
        Set conClient = Nothing
        Exit Sub
    End If
    rstClient.Close
    Set rstClient = Nothing
    
    szSQL = "SELECT * " & _
         "FROM NLPOSTING " & _
         "WHERE ClientID = '" & txtClientID.text & "' AND DeleteFlag=false;"
    rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly
    If Not rstClient.EOF Then
        MsgBox "This client could not be deleted. There are transactions in database related to this client" & _
              (Chr$(13) + Chr$(10)) & "To delete this client please delete transactions.", vbCritical + vbOKOnly, "Delete Not Possible"
        rstClient.Close
        Set rstClient = Nothing
        conClient.Close
        Set conClient = Nothing
        Exit Sub
    End If
    rstClient.Close
    Set rstClient = Nothing
    
   ' MousePointer = vbHourglass
    If MsgBox("Are you sure to delete current client?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub
    
    conClient.Execute "DELETE  FROM CLIENT WHERE CLIENTID = '" & txtClientID.text & "';"
    conClient.Execute "DELETE  FROM Supplier WHERE SupplierID = '" & txtClientID.text & "' AND Type='CLIENT';"
    conClient.Execute "DELETE  FROM NominalLedger WHERE CLIENTID = '" & txtClientID.text & "';"
    
    Dim rsFinancialYear As New ADODB.Recordset
    rsFinancialYear.Open "Select * from financialYear where CLIENTID = '" & txtClientID.text & "'", conClient, adOpenStatic, adLockReadOnly
    While Not rsFinancialYear.EOF
            conClient.Execute "DELETE  FROM Periods where FYrID='" & rsFinancialYear("FYrID").Value & "'"
            rsFinancialYear.MoveNext
    Wend
    conClient.Execute "DELETE  FROM financialYear WHERE CLIENTID = '" & txtClientID.text & "';"
    
    rsFinancialYear.Close
    Set rsFinancialYear = Nothing
    
    
    Set rstClient = Nothing
    conClient.Close
    Set conClient = Nothing
    
    MsgBox "Client has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"
    cmdDeleteClient.Enabled = False
    
    txtClientID.text = ""
    txtClientName.text = ""
'    txtResidency.text = ""
    txtAcBalance(0).text = ""
    txtVATReg.text = ""
    txtYearEndDate.text = ""
    FocusControl cmdAddNewClient
    MousePointer = vbDefault
End Sub

Private Sub cmdEditClient_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      FocusControl cmdClient
      Exit Sub
   End If
   bNewEdit = False
   MainCommandButtonEnable True

   ADD_NEW_CLIENT = False
   LockingAllTextClientAddress False
   UnlockMainClientText True
   '**control mode by anol
'    cboProperty.Enabled = True
'    txtREVIEW_DATE.Locked = False
    cmdAgrTopSave.Enabled = True
    txtNoOfDaysToSendMFB4Due.Locked = False
    cmdAgrTopSave.Enabled = False
    tabAgreement.Enabled = True
    cmdAgmntAddNew.Enabled = True
    cmdAgrTopEdit.Enabled = True
    cmdPayAddNew.Enabled = True
    cmdSetDefaultAC.Enabled = True
    cmdDeleteBank.Enabled = True
    cmdBACS.Enabled = True
    cmdAddNewBank.Enabled = True
    cmdSaveBank.Enabled = True
    chkOverDraft.Enabled = False
    chkConsolidated.Enabled = True
'    cboDmdPropertyList.Enabled = True
    'cmdGSEdit.Enabled = True
    cmdImgLeftMove.Enabled = True
    cmdUploadImageAdd.Enabled = True
    cmdImgDelete.Enabled = True
    cmdDeleteClient.Enabled = False
    cmdAgmntEdit.Enabled = False
    cmdSaveBank.Enabled = False
    cmdSetDefaultAC.Enabled = False
    tabAgreement.Enabled = True
    txtYearEndDate.Locked = True
    FocusControl cmdSaveClient
    cmdVAT.Enabled = True
    chkOptedtoTax.Enabled = True
    cmdBrowseTemplate.Enabled = True
End Sub

Private Sub MainCommandButtonEnable(bEnabled As Boolean)
   cmdAddNewClient.Enabled = Not bEnabled
   cmdEditClient.Enabled = Not bEnabled
   cmdSaveClient.Enabled = bEnabled
   cmdDeleteClient.Enabled = Not bEnabled
   cmdCancelChange.Enabled = bEnabled

   cmdClient.Enabled = Not bEnabled
   'lstResidency.Enabled = bEnabled
End Sub
'
'Private Sub cmdFeeType_Click()
'   Dim sSQLQuery_ As String
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "CFT"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoFeeTypes.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT CODE, VALUE " & _
'              "FROM SECONDARYCODE " & _
'              "WHERE SECONDARYCODE.PRIMARYCODE = 'CFT' " & _
'              "ORDER BY VALUE;"
'
'   adoFeeTypes.RecordSource = sSQLQuery_
'   adoFeeTypes.CommandType = adCmdText
'   adoFeeTypes.Refresh
'End Sub
'
'Private Sub cmdFeeTypesCancel_Click()
'   Dim i As Integer
'
'   If MsgBox("Do you want to discard all modified/new fees?", vbQuestion + vbYesNo, "New Fees") = vbYes Then
'      i = 1
'      While i < flxFeeType.Rows
'         If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "1" Or flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "2" Then
'            flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "0"
'         End If
'         If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "3" Then
'            flxFeeType.RemoveItem i
'            i = i - 1
'         End If
'         i = i + 1
'      Wend
''      ClientGlobalSetting
''      ConfigureFlxFeeType
''      ComponentInFrameEnableMode Me, imgFeeTypes, DefaultMode
'   End If
'End Sub
'
'Private Sub cmdFeeTypesEdit_Click()
'   If flxFeeType.TextMatrix(flxFeeType.Row, flxFeeType.Cols - 1) = "0" Then
'      bNewEdit = False
''      ComponentInFrameEnableMode Me, imgFeeTypes, EditMode
'      flxFeeType.TextMatrix(flxFeeType.Row, flxFeeType.Cols - 1) = "1"
'   End If
'End Sub

Private Sub cmdFeeTypesNew_Click()
  ' MousePointer = vbHourglass

   bNewEdit = True
'   ComponentInFrameEnableMode Me, imgFeeTypes, NewEntryMode

   MousePointer = vbDefault
End Sub
'
'Private Sub cmdFeeTypesSave_Click()
'   If cmdUpdate.Enabled Then
'      MsgBox "Please update the data into grid by pressing the >> button.", vbInformation + vbOKOnly, "Save"
'      cmdUpdate.SetFocus
'      Exit Sub
'   End If
'   If MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save") = vbNo Then Exit Sub
'
'   Dim szSQL As String, i As Integer
'   Dim adoConn As New ADODB.Connection
'   Dim oResultSet As ADODB.Recordset
'
'   adoConn.Open getConnectionString
'   Set oResultSet = New ADODB.Recordset
'
'   i = 1
'   While i < flxFeeType.Rows
'      If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "2" Then
'         szSQL = "SELECT * " & _
'              "FROM ClientGDFees " & _
'              "WHERE ThisID = " & flxFeeType.TextMatrix(i, flxFeeType.Cols - 3) & ";"  'flxFeeType.TextMatrix(i, flxFeeType.Cols - 3) -> ThisID of the record
'
'         oResultSet.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
'
'         oResultSet.Fields("FeeType").Value = flxFeeType.TextMatrix(i, 0)
'         oResultSet.Fields("Handling").Value = flxFeeType.TextMatrix(i, 1)
'         oResultSet.Fields("Frequency").Value = flxFeeType.TextMatrix(i, 2)
'         oResultSet.Fields("NtDueDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 3)), "dd mmmm yyyy")
'         oResultSet.Fields("StDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 4)), "dd mmmm yyyy")
'         oResultSet.Fields("ChargeType").Value = CByte(flxFeeType.TextMatrix(i, 5))
'         oResultSet.Update
'         oResultSet.Close
'      End If
'      If flxFeeType.TextMatrix(i, flxFeeType.Cols - 1) = "3" Then
'         szSQL = "SELECT * " & _
'              "FROM ClientGDFees;"
'
'         oResultSet.Open szSQL, adoConn, adOpenStatic, adLockOptimistic
'
'         oResultSet.AddNew
'         oResultSet.Fields("FeeType").Value = flxFeeType.TextMatrix(i, 0)
'         oResultSet.Fields("Handling").Value = flxFeeType.TextMatrix(i, 1)
'         oResultSet.Fields("Frequency").Value = flxFeeType.TextMatrix(i, 2)
'         oResultSet.Fields("NtDueDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 3)), "dd mmmm yyyy")
'         oResultSet.Fields("StDate").Value = Format(CDate(flxFeeType.TextMatrix(i, 4)), "dd mmmm yyyy")
'         oResultSet.Fields("ChargeType").Value = CByte(flxFeeType.TextMatrix(i, 5))
'         oResultSet.Fields("ParentID").Value = CByte(flxFeeType.TextMatrix(i, flxFeeType.Cols - 2))
'         oResultSet.Update
'         oResultSet.Close
'      End If
'      i = i + 1
'   Wend
'   MsgBox "Data has been saved successfully", vbInformation + vbOKOnly, "Saved"
''   ClientGlobalSetting
''   ConfigureFlxFeeType
''   ComponentInFrameEnableMode Me, imgFeeTypes, DefaultMode
'End Sub
'
'Private Sub cmdFreq_Click()
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "FREQ"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'   Dim sSQLQuery_ As String
'
'   adoFreq.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT CODE, VALUE " & _
'              "FROM SECONDARYCODE " & _
'              "WHERE SECONDARYCODE.PRIMARYCODE = 'FREQ' " & _
'              "ORDER BY VALUE;"
'
'   adoFreq.RecordSource = sSQLQuery_
'   adoFreq.CommandType = adCmdText
'   adoFreq.Refresh
'End Sub

Private Sub cmdGridUnitLookup_Click()
    tabMain.Enabled = True
    picMain.Enabled = True
    Frame5.Visible = False
    FocusControl cmdClient
End Sub

'Private Sub cmdGSCancel_Click()
'   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
'
'   Dim i As Integer
'
'   On Error Resume Next
'   For i = 0 To 67
'      Label1(i).ForeColor = vbBlack
'   Next i
'
''   EnableGlobalControl False
'End Sub

'Private Sub cmdGSEdit_Click()
'  ' MousePointer = vbHourglass
'
'   EnableGlobalControl True
'
''   MousePointer = vbDefault
'End Sub

'Private Sub EnableGlobalControl(bEnable As Boolean)
'   Dim i As Integer
''   cboDmdPropertyList.Enabled = bEnable
''   For i = 0 To 5
''      If i < 6 Then fraPaymentDate(i).Enabled = bEnable
''   Next i
'
'   'cmdGSEdit.Enabled = Not bEnable
'   'cmdGSSave.Enabled = bEnable
'   'cmdGSCancel.Enabled = bEnable
'   cmdAutoSetup(0).Enabled = bEnable
'End Sub

Private Function ControlFilled() As Boolean
'   ControlFilled = True
'
''   If cboDmdPropertyList.text = "" Then
''      Label19(3).ForeColor = vbRed
''      ControlFilled = False
''   End If
'
'   Dim i As Integer
'
'   For i = 0 To 11         'MONTHLY
'      If cboDay(i).text = "" Or cboMonth(i).text = "" Then
'         Label1(35 + i).ForeColor = vbRed
'         ControlFilled = False
'      End If
'      If i < 4 Then        'QUARTERLY
'         If cboQDay(i).text = "" Or cboQMth(i).text = "" Then
'            Label1(47 + i).ForeColor = vbRed
'            ControlFilled = False
'         End If
'      End If
'      If i < 2 Then        'HALF YEARLY
'         If cboHDay(i).text = "" Or cboHMth(i).text = "" Then
'            Label1(51 + i).ForeColor = vbRed
'            ControlFilled = False
'         End If
'      End If
'   Next i
'   If cboYDay(0).text = "" Or cboYMth(0).text = "" Then        'YEARLY
'      Label1(67).ForeColor = vbRed
'      ControlFilled = False
'   End If
End Function

Private Sub cmdGSSave_Click()
   
End Sub

Private Sub cmdImgDelete_Click()
   If imgPremises.Picture = 0 Then Exit Sub
   If MsgBox("Are you sure to delete the image?", vbQuestion + vbYesNo, "Delete Image") = vbNo Then Exit Sub
   DeleteImage imgPremises, IMAGE_FILE_NAME_, szaPremisisIDType(0), szaPremisisIDType(1)
   MsgBox "File has been deleted successfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdImgLeftMove_Click()
   IMAGE_FILE_NAME_ = MoveNextImage(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), IMAGE_FILE_NAME_, lblImageName)
End Sub

Private Sub NewBankText(bLock As Boolean, bNew As Boolean)
   txtBANK_NAME.Locked = bLock
   txtBANK_ADDRESS1.Locked = bLock
   txtBANK_ADDRESS2.Locked = bLock
   txtBANK_ADDRESS3.Locked = bLock
   txtBANK_POST_CODE.Locked = bLock

   If Not bNew Then Exit Sub
   cboBank_ID.text = ""
   txtBANK_NAME.text = ""
   txtBANK_ADDRESS1.text = ""
   txtBANK_ADDRESS2.text = ""
   txtBANK_ADDRESS3.text = ""
   txtBANK_POST_CODE.text = ""
   txtNCCODE.text = ""
   txtNominal.text = ""
   txtPaymentMethod.text = ""
   txtBank_AC_Name.text = ""
   txtBANK_SC.text = ""
   txtBANK_AC_NUM.text = ""
   txtconsolidatedAccountName.text = ""
'   txtBacsRef.text = ""
End Sub

'Private Sub cmdClientAddAtch(1)_Click()
'   If cmbFiles.text = "" Then Exit Sub
'  ' MousePointer = vbHourglass
'
'   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
'      MsgBox "File has been moved from original location.", vbExclamation
'
'   MousePointer = vbDefault
'End Sub

Private Sub cmdResidency_Click()
'   lstResidency.Top = txtResidency.Top
'   lstResidency.Left = txtResidency.Left
'   lstResidency.Visible = True
'   lstResidency.ZOrder 0
'   lstResidency.SetFocus
'   lstResidency.ListIndex = 0
End Sub

Private Sub cmdPayAddNew_Click()
   If MsgBox("Do you wish to add a new rent payable record?", vbQuestion + vbYesNo, "Rent Payable") = vbNo Then Exit Sub
   PayableButtonMode NewEntryMode
   PayableClearMode ClearOnlyTextBoxes
   PAYABLE_ADDNEW_MODE = True
   cmdCommandArray(0).SetFocus
End Sub

Private Sub cmdPayCancel_Click()
   If MsgBox("Do you want to cancel?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   PayableButtonMode DefaultMode
   PayableClearMode ClearOnlyTextBoxes
End Sub

Private Sub cmdPayEdit_Click()
   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit") = vbNo Then Exit Sub

   PayableButtonMode EditMode
   PAYABLE_ADDNEW_MODE = False
   flxPayable_displaytextboxes
   cmdPayEdit.Enabled = False
   cmdCommandArray(0).SetFocus
End Sub

Private Sub AlterDefaultAccount(lMY_ID As Long, conBank As ADODB.Connection)
   conBank.Execute "UPDATE tlbClientBanks " & _
                   "SET DEFAULT_AC = FALSE " & _
                   "WHERE MY_ID = " & lMY_ID & ";"
End Sub

Private Function MaxClientBankID(conBank As ADODB.Connection) As Integer
   Dim rstBank As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT Max(tlbClientBanks.MY_ID)+1 AS MaxOfMY_ID " & _
           "FROM tlbClientBanks;"
   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic

   MaxClientBankID = CInt(IIf(rstBank.EOF, 1, IIf(IsNull(rstBank!MaxOfMY_ID), 1, rstBank!MaxOfMY_ID)))  'CInt(rstBank!MaxOfMY_ID)

   rstBank.Close
   Set rstBank = Nothing
End Function

Private Sub cmdSaveBank_Click()
    Dim rsBankFund As New ADODB.Recordset
    Dim conBank As New ADODB.Connection
    Dim rstBank As New ADODB.Recordset
    Dim iRow As Long
    Dim szSQL As String, szWhere As String, lSpare As Long
   
   If cboBank_ID.text = "" Then
        ShowMsgInTaskBar "Please select the bank ID.", "Y", "N"
        FocusControl cboBank_ID
        Exit Sub
   End If
   If txtPaymentMethod.text = "" Then
        ShowMsgInTaskBar "Please select the correct Payment Method.", "Y", "N"
        FocusControl cmdPaymentTypeNew(0)
        Exit Sub
   End If
'   If txtNCCODE.ListCount <= 0 Then
'      ShowMsgInTaskBar "Please setup the client's chart of account", "Y", "N"
'      cboNC.SetFocus
'      Exit Sub
'   End If
   If txtBANK_SC.text = "" Then
        ShowMsgInTaskBar "Please enter  sort code", "Y", "N"
        FocusControl txtBANK_SC
        Exit Sub
   End If
   If txtNCCODE.text = "" Then
      ShowMsgInTaskBar "Please select the bank code", "Y", "N"
      FocusControl cmdNC
      Exit Sub
   End If

   If txtBank_AC_Name.text = "" And chkConsolidated.Value = 0 Then
      ShowMsgInTaskBar "Please enter the Bank Account Name", "Y", "N"
      FocusControl txtBank_AC_Name
      Exit Sub
   End If

   If txtBANK_AC_NUM.text = "" Then
      ShowMsgInTaskBar "Please enter the Account Number", "Y", "N"
      txtBANK_AC_NUM.SetFocus
      Exit Sub
   End If

   
   flxOtherBankDetails.Enabled = True
   Frame14.Enabled = True
   On Error GoTo ErrorHandler

   conBank.Open getConnectionString

   If bBankNewEdit Then
      szWhere = ""
   Else
      szWhere = " Where MY_ID = " & flxOtherBankDetails.TextMatrix(iSlectedRow, 13) & ";"
   End If

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks" & szWhere
   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
   With rstBank
      If bBankNewEdit Then .AddNew

      !CLIENT_ID = txtClientID.text
      !BANK_ID = cboBank_ID.text
      'If chkConsolidated.Value = 0 Then
            !Bank_AC_Name = txtBank_AC_Name.text
      'Else
       '     !Bank_AC_Name = cboconsolidatedAccountName.text
      'End If
      !BANK_AC_NUM = txtBANK_AC_NUM.text
      !BANK_SC = txtBANK_SC.text
      !nominalCode = txtNCCODE.text              'Non SAGE version
      !AllowOverDraft = IIf(chkOverDraft.Value = 0, False, True)
      'bellow line has been added by anol 04 May 2015 Implementation of Consolidated function
      !Consolidated = IIf(chkConsolidated.Value = 0, False, True)
      'End of modification
      'Implementation of fund in relation with bank account for automatic allocation
'      If txtFundCode.text = "" Or txtFundCode.text = "N/A" Then
'            !FundID = Null
'      Else
'            !FundID = txtFundCode.Tag
'      End If
      If chkConsolidated.Value = 1 Then
            '!ConsBankACNumber = returnConBanID(cboconsolidatedAccountName.text)
            !ConsBankACNumber = txtconsolidatedAccountName.Tag
            !ConsolidatedBankID = intConsolidatedBankID 'IIf(!ConsBankACNumber = "", Null, !ConsBankACNumber)
            !ConsBankACNumber = txtBANK_AC_NUM.text
            !ConsSortCode = txtBANK_SC.text
      Else
            !ConsolidatedBankID = Null
            !ConsBankACNumber = Null
            !ConsSortCode = Null
      End If
'      End If
      If txtOverDraft.text = "" Then
         !OverdraftLimit = Null
      Else
         !OverdraftLimit = CCur(txtOverDraft.text)
      End If
      If iTotalBankAC = 0 Then
         !DEFAULT_AC = True
      Else
         If Not bDefaultAccount Then
            !DEFAULT_AC = False
         Else
            If bBankNewEdit Then   'set new bank a/c default
            'need to alter and set this ac as default
               If lDefaultBankID <> 0 Then Call AlterDefaultAccount(lDefaultBankID, conBank)
               !DEFAULT_AC = True
               lDefaultBankID = flxOtherBankDetails.TextMatrix(iSlectedRow, 13)
            Else     'edit mode
            'in the edit mode first check the exiting situation.
               If Not !DEFAULT_AC Then             'in the else condition nothing to be changed
                  'need to alter and set this ac as default
                  If lDefaultBankID <> 0 Then Call AlterDefaultAccount(lDefaultBankID, conBank)
                  lDefaultBankID = flxOtherBankDetails.TextMatrix(iSlectedRow, 13)
                  !DEFAULT_AC = True
               End If
            End If
         End If
      End If
      !PaymentMethod = txtPaymentMethod.text
'      !BacsRef = txtBacsRef.text

      .Update
      .Close
   End With

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks " & _
           "WHERE ISNULL(spare1);"
   rstBank.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
   While Not rstBank.EOF
      rstBank!spare1 = rstBank!My_ID
      rstBank.Update
      rstBank.MoveNext
   Wend
   'Here I delete all entries from BankFund and add  entries BankFund which are rticked from the list
   conBank.Execute "Delete from BankFund where ClientID='" & txtClientID.text & "' and BankCode='" & txtNCCODE.text & "'"
   szSQL = "SELECT * " & _
           "FROM BankFund "
   rsBankFund.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
   For iRow = 1 To flxBankAccountFund.Rows - 1
        If flxBankAccountFund.TextMatrix(iRow, 0) = "X" Then
            rsBankFund.AddNew
            rsBankFund!ClientID = txtClientID.text
            rsBankFund!BankCode = txtNCCODE.text
            rsBankFund!fundID = flxBankAccountFund.TextMatrix(iRow, 1)
            rsBankFund.Update
       End If
   Next iRow
   rsBankFund.Close
   
   szSQL = "SELECT * " & _
           "FROM Client  where ClientID='" & txtClientID.text & "' "
   rsBankFund.Open szSQL, conBank, adOpenDynamic, adLockOptimistic
   'rsBankFund!ShowBankAccountFunds = IIf(chkShowFundBankAccount.Value = 1, True, False)
   rsBankFund.Update
   rsBankFund.Close
   chkShowFundBankAccount.Value = 0
   
   cmdSetDefaultAC.Enabled = True

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   ConfigFlxOtherBankDetails 'by anol 20170314
   LoadFlxOtherBankDetails

   EnableDisableAcText True
   CommandButtonEnabled True
   If bBankNewEdit Then
        'cmdEditBank.Enabled = False
   End If
   cmdDeleteBank.Enabled = True
'   cmdEditBank.Enabled = True
   cmdSetDefaultAC.Enabled = False
   cmdBACS.Enabled = False
   cmdDeleteBank.Enabled = False
   cmdClient.Enabled = True
   Exit Sub
NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub cmdSaveClient_Click()
   If txtClientID.text = "" Then
      MsgBox "Please type client id.", vbCritical + vbOKOnly, "Client"
      txtClientID.SetFocus
      Exit Sub
   End If
   If txtClientName.text = "" Then
      MsgBox "Please type client's name.", vbCritical + vbOKOnly, "Client"
      txtClientName.SetFocus
      Exit Sub
   End If

'   If txtResidency.text = "" Then
'      MsgBox "Please select client's residency.", vbCritical + vbOKOnly, "Client"
'      txtResidency.SetFocus
'      Exit Sub
'   End If
   If txtVATReg.text = "" Then
      If MsgBox("Is this client registered for  VAT?" & (Chr(13) + Chr(10)), vbQuestion + vbYesNo + vbDefaultButton2, "TAX/VAT Number") = vbYes Then
         txtVATReg.SetFocus
         Exit Sub
      End If
   End If
   If tabMain.Tab = 3 Then
        If cmdSaveBank.Enabled = True Then
             MsgBox "Please save Bank details first to continue.", vbCritical + vbOKOnly, "Warning"
             FocusControl cmdSaveBank
             Exit Sub
        End If
   
   End If
'   If txtYearEndDate.text = "" Then
'      MsgBox "Please type year end date.", vbCritical + vbOKOnly, "Client"
'      txtYearEndDate.SetFocus
'      Exit Sub
'   End If

   If txtAcBalance(0).text = "" Then txtAcBalance(0).text = "0.00"

   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim Rst2 As New ADODB.Recordset

   szSQL = "SELECT ClientID, ClientName, LandLordSageCustAC, " & _
                  "LandLordSageSuppAC, Residency, AcBalance, VATReg, " & _
                  "YearEndDate, CT,RentSummaryTemplate,LastModifiedBy,LastModifiedDate,CreatedBy,CreatedDate " & _
           "FROM Client " & _
           "WHERE ClientID = '" & txtClientID.text & "';"

   adoConn.Open getConnectionString
   If SaveEdit(Me, picMain, adoConn, szSQL, bNewEdit) Then
   
   '  Here we are syncronyzing supllier table with client fields if the record does not exists in the supplier table. except two field vatcode and opted to tax
   
      szSQL = "SELECT ClientID,ClientName,ClientAddressLine1,ClientAddressLine2,ClientAddressLine3,ClientAddressLine4,ClientPostCode,ClientOfficeAddressLine1,ClientOfficeAddressLine2, " & _
      "ClientOfficeAddressLine3,ClientOfficeAddressLine4,ClientOfficePostCode,ClientOfficeEmail,ClientPersonalEmail,'CLIENT' From Client C  LEFT JOIN Supplier S ON  C.ClientID=S.SupplierID  " & _
      "where isnull(S.SupplierID)"
      Rst2.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
      If Not Rst2.EOF Then
             adoConn.Execute "Insert into Supplier(SupplierID,SupplierName,SupplierAddressLine1,SupplierAddressLine2,SupplierAddressLine3,SupplierAddressLine4,SupplierPostCode," & _
             "SupplierOfficeAddressLine1,SupplierOfficeAddressLine2,SupplierOfficeAddressLine3,SupplierOfficeAddressLine4,SupplierOfficePostCode,SupplierOfficeEmail," & _
             "SupplierPersonalEmail,TYPE)  " & szSQL
             
      End If
      Rst2.Close
      Set Rst2 = Nothing
      
      adoConn.Close
      Set adoConn = Nothing
      
      If bNewEdit Then                                               'Create Control code entry
          MsgBox "Client : '" & txtClientName.text & "' added successfully. Please setup the Nominal Control Accounts for this Client.", vbInformation + vbOKOnly, "Information Saved."
      Else
          FocusControl cmdClose
          MsgBox "This Client's record has been successfully created.", vbOKOnly, "Client record has been created"
      End If
       cmdDeleteClient.Enabled = False

       txtClientID.Locked = True
       tabMain.Tab = 0
       UnlockMainClientText False
       MainCommandButtonEnable False
    
       NewBankText True, True
       flxOtherBankDetails.Clear
       chkOptedtoTax.Enabled = False
       cmdBrowseTemplate.Enabled = False
   Else
      MsgBox "This record has not been saved.", vbOKOnly, "Record not saved"
      txtClientID.text = ""
      FocusControl txtClientID
   End If
'   cmdDeleteClient.Enabled = False
'
'   txtClientID.Locked = True
'   tabMain.Tab = 0
'   UnlockMainClientText False
'   MainCommandButtonEnable False
'
'   NewBankText True, True
'   flxOtherBankDetails.Clear
'   chkOptedtoTax.Enabled = False
'   cmdBrowseTemplate.Enabled = False
End Sub
Private Function SaveEdit(frmCurrent As Form, ByVal oContainer As Control, ByVal oConnector As ADODB.Connection, ByVal sSQLQuery As String, ByVal IsNewRecord As Boolean) As Boolean
    Dim iFieldsCount, iControlCount, i, j As Integer
    Dim sNextField As String
    Dim oControl As Control

    Dim oResultSet As New ADODB.Recordset
    On Error GoTo Exception
    oResultSet.Open sSQLQuery, oConnector, adOpenStatic, adLockOptimistic
    
    If IsNewRecord Then
       If oResultSet.EOF Or oResultSet.BOF Then
           oResultSet.AddNew
           oResultSet!CreatedBy = User
           oResultSet!CreatedDate = Now
       Else
           MsgBox "WARNING !! This reference already exists. Please enter a unique reference.", vbInformation
'           txtClientID.text = ""
'           txtClientName.text = ""
'           cmdCancelChange_Click
           FocusControl txtClientID
           SaveEdit = False
           Exit Function
       End If
    Else
       If oResultSet.EOF Or oResultSet.BOF Then
           MsgBox "WARNING !! This record does not exist.", vbInformation
           SaveEdit = False
           Exit Function
       End If
    End If
    oResultSet!ClientID = txtClientID.text
    oResultSet!ClientName = txtClientName.text
    oResultSet!VATReg = txtVATReg.text
    oResultSet!CT = txtCT.text
    oResultSet!LastModifiedBy = User
    oResultSet!LastModifiedDate = Now
    'oResultSet!RentSummaryTemplate = txtRenSummaryStatement.text
   
 
   If Len(oResultSet.Fields(0).Value) > 10 Then
        MsgBox "Client ID cannot be more than 10 character", vbInformation, "Failed to save"
        Exit Function
   End If
   oResultSet.Update
   Call cmdClientDetailsSave_Click(oConnector) 'Save details address in this function
   SaveEdit = True
   oResultSet.Close
   Set oResultSet = Nothing
   Exit Function

Exception:
   SaveEdit = False
   MsgBox Err.description
   Set oResultSet = Nothing
End Function


Private Sub cmdSetDefaultAC_Click()
    'if you click this button then you are going to set a default bank account for a selected row
    'Now there is a scenario where if you change the selection after setting a default bank account.
    'i.e select a line from grid to make default bank account, click this buttton and then again change the selection from the grid
    'Then this problem will happen. Now what I need to do is to prevent the user from selection of the grid after you click this button- anol 2020-05-23
    'So need to set enabled false to that grid and release when save and cancel and when you change the client
    Dim X As Integer
    If flxOtherBankDetails.row = 0 Then MsgBox "Please select a bank to edit", vbInformation, "Warning"
    If txtNCCODE.text <> flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14) Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4) <> txtBank_AC_Name.text Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6) <> txtBANK_SC.text Or _
        UCase(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)) <> UCase(txtPaymentMethod.text) Or _
        flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5) <> txtBANK_AC_NUM.text Then
        X = MsgBox("Do you want to save the current changes you made?", vbYesNo, "Save?")
        If X = vbYes Then
            FocusControl cmdSaveBank
            Exit Sub
        Else
            Call restoreSelectedValueflxOtherBankDetails
        End If
    End If
    
    
   Dim i As Integer
   If iTotalBankAC = 0 Then Exit Sub      'there are no bank details has been inputed yet
   If bDefaultAccount Then
        MsgBox "This bank account is currently set as default. Please select a different bank account", vbInformation, "Warning"
        Exit Sub
   End If

   'If Not cmdSaveBank.Enabled Then Exit Sub
   Dim adoConn As New ADODB.Connection
   If MsgBox("Do you wish to set this bank account as the default bank account for this client?", vbQuestion + vbYesNo, "Default Bank Account settings") = vbNo Then
      Exit Sub
   Else
      cmdSetDefaultAC.Enabled = False
      cmdSaveBank.Enabled = False
      bDefaultAccount = True
      adoConn.Open getConnectionString
      'MsgBox iSlectedRow
      For i = 1 To flxOtherBankDetails.Rows - 1
            If i = iSlectedRow Then
                flxOtherBankDetails.TextMatrix(i, 7) = "YES"
                   adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET DEFAULT_AC = TRUE " & _
                   "WHERE MY_ID = " & flxOtherBankDetails.TextMatrix(i, 13) & ";"
            Else
                flxOtherBankDetails.TextMatrix(i, 7) = "NO"
                adoConn.Execute "UPDATE tlbClientBanks " & _
                   "SET DEFAULT_AC = FALSE " & _
                   "WHERE MY_ID = " & flxOtherBankDetails.TextMatrix(i, 13) & ";"
            End If
            'IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", True, False)
      Next
      adoConn.Close
      Set adoConn = Nothing
      MsgBox "This account has been set as the default bank account.", vbInformation, "Default Bank Account"
      flxOtherBankDetails.Enabled = False
      FocusControl cmdSaveBank
      
        'Now cancel button start
        flxOtherBankDetails.Enabled = True
        bOverdraftWarning = False
        CommandButtonEnabled True
        EnableDisableAcText True
        NewBankText True, True
        flxOtherBankDetails_RowColChange
        cmdSetDefaultAC.Enabled = True
        'added by anol 11 Mar 2015
        cboBank_ID.Locked = True
        cmdSaveBank.Enabled = False
        cmdSetDefaultAC.Enabled = False
    '    cmdDeleteBank.Enabled = False
        cmdBACS.Enabled = False
        cmdClient.Enabled = True
        Frame14.Enabled = True
        cmdClient.Enabled = True
   End If
End Sub

'Private Sub cmdUnitMemoCancel_Click()
'   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
'   MemoButtonEnable False
'End Sub

'Private Sub cmdUnitMemoEdit_Click()
'   MemoButtonEnable True
'End Sub

'Private Sub cmdUnitMemoSave_Click()
'   If (NEW_TYPE = "Landlord") Then
'      If SaveMemo("Landlord", "LandlordMemo", txtClientID.text, "LandlordID", txtNote) Then
'         ShowMsgInTaskBar "Memo has been saved successfully."
'      End If
'   Else
'      If SaveMemo("Client", "ClientMemo", txtClientID.text, "ClientID", txtNote) Then
'         ShowMsgInTaskBar "Memo has been saved successfully."
'      End If
'   End If
'   MemoButtonEnable False
'End Sub

'Private Sub MemoButtonEnable(bEnable As Boolean)
'   txtNote.Locked = Not bEnable
'   cmdUnitMemoEdit.Enabled = Not bEnable
'   cmdUnitMemoSave.Enabled = bEnable
'   cmdUnitMemoCancel.Enabled = bEnable
'End Sub

Private Function CreateClientId(szName As String) As String
   Dim szSQL As String, i As Integer, szChar As String, j As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRSTClient As New ADODB.Recordset, adoRSTLandlord As New ADODB.Recordset

   For i = 1 To Len(szName) - 1
      szChar = UCase(Mid(szName, i, 1))
      If (szChar >= "A" And szChar <= "Z") Then
         CreateClientId = CreateClientId & szChar
         j = j + 1
      End If
      If j = 8 Then Exit For
   Next i

   If j < 8 Then CreateClientId = Left(CreateClientId & "01234567", 8)

   adoConn.Open getConnectionString

   szSQL = "SELECT ClientID " & _
           "FROM Client " & _
           "WHERE Client.ClientID = '" & CreateClientId & "';"
   adoRSTClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT LandlordID " & _
           "FROM Landlord " & _
           "WHERE Landlord.LandlordID = '" & CreateClientId & "';"
   adoRSTLandlord.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   j = 1
   Do
      If adoRSTClient.EOF And adoRSTLandlord.EOF Then Exit Do

      adoRSTClient.Close
      adoRSTLandlord.Close

      CreateClientId = Left(CreateClientId & "01234567", 6) & Format(j, "00")
      szSQL = "SELECT ClientID " & _
              "FROM Client " & _
              "WHERE Client.ClientID = '" & CreateClientId & "';"
      adoRSTClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      szSQL = "SELECT LandlordID " & _
              "FROM Landlord " & _
              "WHERE Landlord.LandlordID = '" & CreateClientId & "';"
      adoRSTLandlord.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      j = j + 1
   Loop

   adoRSTClient.Close
   adoRSTLandlord.Close
   adoConn.Close
   Set adoRSTClient = Nothing
   Set adoRSTLandlord = Nothing
   Set adoConn = Nothing
End Function

Private Function GetParentID(szParent As String) As String
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   szSQL = "SELECT Record_ID " & _
           "FROM ClientGlobalData " & _
           "WHERE ClientGlobalData.ClientID = '" & flxClientList.TextMatrix(flxClientList.row, 1) & "';"

   Dim adoRST As ADODB.Recordset
   Set adoRST = New ADODB.Recordset

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   GetParentID = adoRST.Fields("Record_ID").Value

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Function

Private Sub cmdUploadImageAdd_Click()
   On Error GoTo Err
   If MsgBox("Do you want to add new image?", vbQuestion + vbYesNo, "Image Attachment") = vbNo Then Exit Sub
   IMAGE_FILE_NAME_ = AddNewImage(imgPremises, szaPremisisIDType(1), szaPremisisIDType(0), lblImageName)
   MsgBox "Image has been uploaded successfull."
   Exit Sub
Err:
    MsgBox Err.description
End Sub

Private Sub flxACHistory_Click()
   Dim iCurRowHeight As Integer
   Dim iRow          As Integer
   Dim adoConn       As New ADODB.Connection
   Dim adoRST        As New ADODB.Recordset
   Dim szSQL         As String

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "" Then GoTo ChildGrid

'****************************************************** EXPANDING THE GRID *********************************
   iRow = flxACHistory.row
   iCurRowHeight = flxACHistory.RowHeight(iRow)

   If flxACHistory.col = 0 Then
      If flxACHistory.TextMatrix(iRow, 0) = "-" Then Exit Sub
      If flxACHistory.TextMatrix(iRow, 0) = "+" And flxACHistory.RowHeight(iRow + 1) = 0 Then
      If flxACHistory.TextMatrix(iRow, 0) = "" Then Exit Sub
         flxACHistory.TextMatrix(iRow, 0) = ">"
         For iRow = iRow + 1 To flxACHistory.Rows - 1
            If flxACHistory.TextMatrix(iRow, 0) = "+" Or flxACHistory.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxACHistory.TextMatrix(iRow, 0) = "-" Then flxACHistory.RowHeight(iRow) = iCurRowHeight
         Next iRow
      ElseIf flxACHistory.TextMatrix(iRow, 0) = ">" And flxACHistory.RowHeight(iRow + 1) = iCurRowHeight Then
         flxACHistory.TextMatrix(iRow, 0) = "+"
         For iRow = iRow + 1 To flxACHistory.Rows - 1
            If flxACHistory.TextMatrix(iRow, 0) = "+" Or flxACHistory.TextMatrix(iRow, 0) = ">" Then Exit For
            If flxACHistory.TextMatrix(iRow, 0) = "-" Then flxACHistory.RowHeight(iRow) = 0
         Next iRow
      End If
   End If
'***********************************************************************************************************
   HighLightRowFlxGrid flxACHistory, flxACHistory.row

   If flxACHistory.TextMatrix(flxACHistory.row, 0) = "-" Then Exit Sub

ChildGrid:
'  Displaying the splits ************************************************************************************
   
   ConfigFlxACHistorySplit
   adoConn.Open getConnectionString

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PC" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbPayment AS P, tblPurInv AS I, tblPurInvSRec AS S " & _
              "WHERE P.PI = I.MY_ID AND " & _
                  "I.MY_ID = S.ParentID AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 11) & " " & _
              "ORDER BY S.MY_ID;"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NOMINAL_CODE").Value
            .TextMatrix(iRow, 5) = IIf(IsNull(adoRST.Fields.Item("JOB_ID").Value), "", adoRST.Fields.Item("JOB_ID").Value)
            .TextMatrix(iRow, 6) = adoRST.Fields.Item("UNIT_ID").Value
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("DEPT_ID").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("DESCRIPTION").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRST.Close
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PP" And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "PPR" Then
      szSQL = "SELECT S.*, P.ExtRef, P.UnitID, P.FundID " & _
              "FROM tlbPayment AS P, PayTransactions AS S " & _
              "WHERE P.TransactionID = S.FromTran AND " & _
                  "P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 11) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRST.Fields.Item("UnitID").Value), "", adoRST.Fields.Item("UnitID").Value)
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("FundID").Value
            .TextMatrix(iRow, 8) = flxACHistory.TextMatrix(flxACHistory.row, 5)
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("PaymentAmount").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("PaymentAmount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then
      szSQL = "SELECT P.* " & _
              "FROM tlbPayment AS P " & _
              "WHERE P.TransactionID = " & flxACHistory.TextMatrix(flxACHistory.row, 11) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = iRow
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = flxACHistory.TextMatrix(flxACHistory.row, 3)
            .TextMatrix(iRow, 3) = adoRST.Fields.Item("ExtRef").Value
            .TextMatrix(iRow, 4) = adoRST.Fields.Item("NominalCode").Value
            .TextMatrix(iRow, 5) = ""
            .TextMatrix(iRow, 6) = IIf(IsNull(adoRST.Fields.Item("UnitID").Value), "", adoRST.Fields.Item("UnitID").Value)
            .TextMatrix(iRow, 7) = adoRST.Fields.Item("FundID").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Details").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "PPR" Then _
               .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "PA" Then _
               .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI" Or _
      Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SC" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbReceipt AS R, DemandRecords AS D, DemandSplitRecords AS S " & _
              "WHERE R.DemandRef = D.DemandID AND " & _
                  "D.DemandID = S.DemandID AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SI", 1, 2) & " AND " & _
                  "R.SlNumber = " & Mid(flxACHistory.TextMatrix(flxACHistory.row, 1), 3, _
                                    Len(flxACHistory.TextMatrix(flxACHistory.row, 1)) - 2) & " " & _
              "ORDER BY S.SplitID;"

      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = adoRST.Fields.Item("SplitID").Value
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = adoRST.Fields.Item("DueDate").Value
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 10) = Format(adoRST.Fields.Item("TotalAmount").Value, "0.00")
            .TextMatrix(iRow, 11) = ""
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
         adoRST.Close
      End With
   End If

   If (Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR" Or _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SA") And _
       Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) <> "SRR" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S " & _
              "WHERE R.TransactionID = S.RptHeader AND " & _
                  "R.Type = " & IIf(Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 2) = "SR", 3, 4) & " AND " & _
                  "R.SlNumber = " & Mid(flxACHistory.TextMatrix(flxACHistory.row, 1), 3, _
                                    Len(flxACHistory.TextMatrix(flxACHistory.row, 1)) - 2) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = ""
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = IIf(IsNull(adoRST.Fields.Item("DueDate").Value), "", adoRST.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If

   If Left(flxACHistory.TextMatrix(flxACHistory.row, 1), 3) = "SRR" Then
      szSQL = "SELECT S.* " & _
              "FROM tlbReceipt AS R, tlbReceiptSplit AS S " & _
              "WHERE R.TransactionID = S.RptHeader AND " & _
                  "R.Type = 23 AND " & _
                  "R.SlNumber = " & Mid(flxACHistory.TextMatrix(flxACHistory.row, 1), 4, _
                                    Len(flxACHistory.TextMatrix(flxACHistory.row, 1)) - 3) & ";"
'Debug.Print szSQL
      adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      iRow = 1
      With flxACHistorySplit
         While Not adoRST.EOF
            .TextMatrix(iRow, 0) = ""
            .TextMatrix(iRow, 1) = flxACHistory.TextMatrix(flxACHistory.row, 2)
            .TextMatrix(iRow, 2) = IIf(IsNull(adoRST.Fields.Item("DueDate").Value), "", adoRST.Fields.Item("DueDate").Value)
            .TextMatrix(iRow, 8) = adoRST.Fields.Item("Description").Value
            .TextMatrix(iRow, 9) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            .TextMatrix(iRow, 10) = ""
            .TextMatrix(iRow, 11) = Format(adoRST.Fields.Item("Amount").Value, "0.00")
            adoRST.MoveNext
            If Not adoRST.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
      End With
   End If
   
   adoConn.Close
   Set adoConn = Nothing
End Sub

'Private Sub flxAgreement_RowColChange()
'   populateControl Me, flxAgreement
'   cmdAgmntEdit.Enabled = True
'End Sub

Private Sub GetFinancialYearEnd(Conn As ADODB.Connection)
    Dim strSQL As String
    Dim rsFinancialYear As New ADODB.Recordset
    Dim rsFinancialYear2 As New ADODB.Recordset
    strSQL = "Select FY_Enddate from FinancialYear where ClientID='" & txtClientID.text & "' AND  setascurrent=true AND status "
    rsFinancialYear.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    If Not rsFinancialYear.EOF Then
        txtYearEndDate.text = rsFinancialYear("FY_Enddate").Value
    Else
            strSQL = "Select FY_Enddate from FinancialYear where ClientID='" & txtClientID.text & "' AND status order by FY_Enddate DESC"
            rsFinancialYear2.Open strSQL, Conn, adOpenStatic, adLockReadOnly
            If Not rsFinancialYear2.EOF Then
                txtYearEndDate.text = rsFinancialYear2("FY_Enddate").Value
            Else
                txtYearEndDate.text = ""
            End If
            rsFinancialYear2.Close
            Set rsFinancialYear2 = Nothing
    End If
    rsFinancialYear.Close
    Set rsFinancialYear = Nothing

End Sub
Private Sub flxClientList_Click()
   Dim adoConn As New ADODB.Connection
   Dim rsClient As New ADODB.Recordset
   Dim rsClient1 As New ADODB.Recordset

   Dim rsClientGlobalData As New ADODB.Connection
   tabMain.Enabled = True
   picMain.Enabled = True
   If strCommandSource = "VAT" Then
        Frame5.Visible = False
        txtAcBalance(1).text = flxClientList.TextMatrix(flxClientList.row, 0) & " / " & flxClientList.TextMatrix(flxClientList.row, 1)
        If Trim(txtAcBalance(1).text) = "/" Then
            txtAcBalance(1).text = ""
        End If
        txtAcBalance(1).Tag = flxClientList.TextMatrix(flxClientList.row, 2)
        chkOptedtoTax.Value = 1
        Exit Sub
   End If
   If strCommandSource = "Client" Then
            cmdEditClient.Enabled = True
            flxOtherBankDetails.Enabled = True
            If flxClientList.TextMatrix(flxClientList.row, 1) = "" Then Exit Sub
            cmdDeleteClient.Enabled = True
            Dim sSQLQuery_    As String
            Dim sFilter       As String
            Dim sType         As String
            Dim X As Integer
            Call ClearPaymentDates
            'Exit Sub
            sType = flxClientList.TextMatrix(flxClientList.row, 3)
            NEW_TYPE = "Client"
            Me.Caption = "Client"
            Label1(0).Caption = "Client ID"

            txtClientID.text = flxClientList.TextMatrix(flxClientList.row, 1)
            txtAcBalance(0).text = flxClientList.TextMatrix(flxClientList.row, 3)
            Dim SQLStr  As String
            adoConn.Open getConnectionString
            GetFinancialYearEnd adoConn
            rsClient1.Open "Select * from Supplier S LEFT JOIN tlbVatCode V ON  S.vatCode=cstr(V.VAT_ID) where supplierID='" & txtClientID.text & "' and Type='Client'", adoConn, adOpenStatic, adLockReadOnly
            If Not rsClient1.EOF Then
                txtAcBalance(1).Tag = IIf(IsNull(rsClient1("VAT_ID").Value), "", rsClient1("VAT_ID").Value)
               ' Label1(27).Caption = IIf(IsNull(rsClient1("VAT_CODE").Value), "", rsClient1("VAT_CODE").Value) ' rsClient1("VAT_CODE").Value
                txtAcBalance(1).text = IIf(IsNull(rsClient1("VAT_CODE").Value), "", rsClient1("VAT_CODE").Value) & " / " & IIf(IsNull(rsClient1("VAT_Rate").Value), "", rsClient1("VAT_Rate").Value) 'rsClient1("VAT_Rate").Value
                If Trim(txtAcBalance(1).text) = "/" Then
                    txtAcBalance(1).text = ""
                End If
                chkOptedtoTax.Value = IIf(rsClient1("OptedtoTax").Value = False, 0, 1)
            End If
            rsClient1.Close
            'SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'RAT'"
            rsClient.Open "Select C.*,(SELECT VALUE FROM SECONDARYCODE WHERE CODE =C.PaymentType and PRIMARYCODE = 'RAT')as PaymentTypeName from Client C where clientID='" & txtClientID.text & "'", adoConn, adOpenKeyset, adLockOptimistic
            If Not rsClient.EOF Then
                txtPaymentType.Tag = IIf(IsNull(rsClient!PaymentType) = True, "", rsClient!PaymentType)
                txtPaymentType.text = IIf(IsNull(rsClient!PaymentTypeName) = True, "", rsClient!PaymentTypeName)
                txtPaymentTerms.text = IIf(IsNull(rsClient!PaymentTerms) = True, "", rsClient!PaymentTerms) ' rsClient!PaymentTerms
                txtRemittanceTemplate.text = IIf(IsNull(rsClient!RemittanceTemplate) = True, "", rsClient!RemittanceTemplate) 'rsClient!RemittanceTemplate
                txtRenSummaryStatement.text = IIf(IsNull(rsClient!RentSummaryTemplate) = True, "", rsClient!RentSummaryTemplate) 'rsClient!RentSummaryTemplate
                'txtRenSummaryStatement.text = IIf(IsNull(rsClient!RentSummaryTemplate), "", rsClient!RentSummaryTemplate) 'rsClient!Comments2
                
                txtClientHomeTel(11).text = IIf(IsNull(rsClient!AccountName) = True, "", rsClient!AccountName) 'rsClient!AccountName
                txtClientHomeTel(12).text = IIf(IsNull(rsClient!SortCode) = True, "", rsClient!SortCode) 'rsClient!SortCode
                txtClientHomeTel(13).text = IIf(IsNull(rsClient!AccountNumber) = True, "", rsClient!AccountNumber) 'rsClient!AccountNumber
                txtClientHomeTel(14).text = IIf(IsNull(rsClient!BankPaymentRef) = True, "", rsClient!BankPaymentRef)  'rsClient!BankPaymentRef
                chkUsePayableTemplate.Value = IIf(IsNull(rsClient!UsePayableTemplate) = True, "0", rsClient!UsePayableTemplate) ' rsClient!UsePayableTemplate
            Else
                txtPaymentType.Tag = ""
                txtPaymentTerms.text = ""
                txtRemittanceTemplate.text = ""
                txtRenSummaryStatement.text = ""
                'txtRenSummaryStatement.text = IIf(IsNull(rsClient!RentSummaryTemplate), "", rsClient!RentSummaryTemplate) 'rsClient!Comments2
                
                txtClientHomeTel(11).text = ""
                txtClientHomeTel(12).text = ""
                txtClientHomeTel(13).text = ""
                txtClientHomeTel(14).text = ""
                chkUsePayableTemplate.Value = 0
            End If
            rsClient.Close
    
    
            SQLStr = "SELECT G.*, F.FinancialYear AS CBY, F.FYrID " & _
                    "FROM (GlobalData AS G INNER JOIN Property AS P ON G.PropertyID = P.PropertyID) " & _
                          "LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
                    "WHERE P.ClientID = '" & txtClientID.text & "';"
            
            rsClient.Open SQLStr, adoConn, adOpenKeyset, adLockReadOnly
            
            If rsClient.EOF Then
                txtYearEndDate.text = ""
            Else
                While Not rsClient.EOF
                    rsClient.MoveNext
                Wend
            End If
XX:
           rsClient.Close
           adoConn.Close
        
           
           fmeLoading.ZOrder 0
           fmeLoading.Visible = True
           fmeLoading.Refresh
           adoConn.Open getConnectionString
           If (NEW_TYPE <> "Landlord") Then
              sSQLQuery_ = "SELECT ClientID, ClientName, ClientAddressLine1, ClientAddressLine2, " & _
                             "ClientAddressLine3, ClientPostCode, ClientOfficeEmail, ClientPersonalEmail, " & _
                             "ClientHomeTel, ClientMobile, '', ClientOfficeAddressLine1, " & _
                             "ClientOfficeAddressLine2, ClientOfficeAddressLine3, ClientOfficePostCode, " & _
                             "ClientOfficeTel, ClientMemo, LandLordSageCustAC, LandLordSageSuppAC, " & _
                             "BANK_ID, CommissionType, CommissionAmt, BGRPayable, VATReg, AcBalance, " & _
                             "Residency, YearEndDate, PaymentMethod, BacsRef, HomeOfficeAdd, CompReg, RegAdd1, " & _
                             "RegAdd2, RegAdd3, RegPostCode, CT, ClientAddressLine4, ClientOfficeAddressLine4, RegAdd4,groupCode,Comments1,Comments2 " & _
                           "FROM CLIENT " & _
                           "WHERE CLIENT.ClientID = '" & txtClientID.text & "';"
           End If
           rsClient.Open sSQLQuery_, adoConn, adOpenKeyset, adLockOptimistic
        
           
           If Not Fill_Form() Then
                 MsgBox "Error in Database.", vbExclamation
              
           End If
        
           LoadClientProperty
'           PrepareList4Property cboDmdPropertyList
          ' PrepareList4Property cboProperty ' added by anol 2020-07-02
        
           lblLoading.Caption = "Please wait, tree is building..."
           fmeLoading.Refresh
           Debug.Print time
           DrawLandLordTree tvwLandLord, imgList, txtClientID.text, True, NEW_TYPE
           Debug.Print time
           lblLoading.Caption = "Please wait, global data is loading..."
        
           fmeLoading.Refresh
        

        
           fmeLoading.Visible = False
           tabMain.Tab = 0
           Call LoadFlxACHistory(adoConn, "")
           If flxClientList.TextMatrix(flxClientList.row, 3) = "Landlord" Then
              LoadFlxACHistory_Sales
           End If
           cmdClient.SetFocus
           'I am clearing bank details when client changes Added by anol 02 Nov 2015
           cboBank_ID.ListIndex = -1
           cboBank_ID.text = ""
           txtBANK_NAME.text = ""
           txtBANK_ADDRESS1.text = ""
           txtBANK_ADDRESS2.text = ""
           txtBANK_POST_CODE.text = ""
           txtNCCODE.text = ""
           txtNominal.text = ""
           txtPaymentMethod.text = ""
           txtBank_AC_Name.text = ""
           txtBANK_SC.text = ""
           txtBANK_AC_NUM.text = ""
           txtconsolidatedAccountName.text = ""
           chkOverDraft.Value = 0
           txtOverDraft.text = ""
           chkConsolidated.Value = 0
          
           'End of addition
           Call LoadGridMemo
           Call ViewMemo
           Call buttonmmodeOnGridClick
           cmdAgmntSave.Enabled = False
     ElseIf strCommandSource = "NominalCode" Then
           txtNCCODE.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtNominal.text = flxClientList.TextMatrix(flxClientList.row, 2)
           Call cboNC_LostFocus
     ElseIf strCommandSource = "PayeeTYPES" Then
           txtPayeeType.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtPayeeType.Tag = flxClientList.TextMatrix(flxClientList.row, 2)
           txtClientLandlord.text = ""
           txtClientLandlord.Tag = ""
     ElseIf strCommandSource = "SupplierTypesFilter" Then
           txtFilterClient.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtFilterClient.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
           txtSupplierFilter.text = ""
           adoConn.Open getConnectionString
           Call LoadFlxACHistory(adoConn, "")
           adoConn.Close
             If txtSupplierFilter.text = "Client" Then
                cmdSupplierFilter.Enabled = False
           Else
                cmdSupplierFilter.Enabled = True
           End If
           FocusControl cmdClientFilter
     ElseIf strCommandSource = "ClientFilter" Then
           txtSupplierFilter.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtSupplierFilter.Tag = flxClientList.TextMatrix(flxClientList.row, 2)
           FocusControl cmdSupplierFilter
           adoConn.Open getConnectionString
           Call LoadFlxACHistory(adoConn, "")
         
     ElseIf strCommandSource = "SupplierFilter" Or strCommandSource = "AgentFilter" Or strCommandSource = "LandlordFilter" Or strCommandSource = "LesseeFilter" Then
           txtSupplierFilter.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtSupplierFilter.Tag = flxClientList.TextMatrix(flxClientList.row, 2)
           FocusControl cmdSupplierFilter
           adoConn.Open getConnectionString
           Call LoadFlxACHistory(adoConn, "6")
           adoConn.Close
    ElseIf strCommandSource = "ConsolidatedBank" Then
           txtconsolidatedAccountName.text = flxClientList.TextMatrix(flxClientList.row, 1)
           txtconsolidatedAccountName.Tag = flxClientList.TextMatrix(flxClientList.row, 2)
           intConsolidatedBankID = flxClientList.TextMatrix(flxClientList.row, 3)
     ElseIf strCommandSource = "PAYABLETYPES" Then
           txtPayableType.Tag = flxClientList.TextMatrix(flxClientList.row, 1) 'ID in 1 Demand desc  in 2
           txtPayableType.text = flxClientList.TextMatrix(flxClientList.row, 2)
           Call FillPayDueDate
           FocusControl cmdCommandArray(2)
    ElseIf strCommandSource = "chargeTypes" Then
           txtChargeType.Tag = flxClientList.TextMatrix(flxClientList.row, 1) 'ID in 1 Demand desc  in 2
           txtChargeType.text = flxClientList.TextMatrix(flxClientList.row, 2)
           FocusControl cmdCommandArray(7)
           'Below paragraph is obsolate now
'     ElseIf strCommandSource = "DEMANDTYPESMangFee" Then
'           txtDemandTypemngtFee.Tag = flxClientList.TextMatrix(flxClientList.row, 1) 'ID in 1 Demand desc  in 2
'           txtDemandTypemngtFee.text = flxClientList.TextMatrix(flxClientList.row, 2)
'           FocusControl cmdCommandArray(7)
           
     ElseIf strCommandSource = "DEMANDTYPES" Then
'           txtPayDemandType.Tag = flxClientList.TextMatrix(flxClientList.row, 1) 'ID in 1 Demand desc  in 2
'           txtPayDemandType.text = flxClientList.TextMatrix(flxClientList.row, 2)
'            FocusControl cmdCommandArray(2)
     ElseIf strCommandSource = "FUND" Then
           txtPayFund.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID,2 code,3 fund Name
           txtPayFund.text = flxClientList.TextMatrix(flxClientList.row, 2)
           FocusControl cmdCommandArray(1)
    ElseIf strCommandSource = "FUNDMangFee" Then
           txtFundMngtFee.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID,2 code,3 fund Name
           txtFundMngtFee.text = flxClientList.TextMatrix(flxClientList.row, 2)
           strSelectedFundName = flxClientList.TextMatrix(flxClientList.row, 3)
           FocusControl cmdCommandArray(6)
    
     ElseIf strCommandSource = "ClientLandlord" Then
           txtClientLandlord.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
           txtClientLandlord.text = flxClientList.TextMatrix(flxClientList.row, 1)
           FocusControl cmdCommandArray(5)
'     ElseIf strCommandSource = "Frequencies" Then
'           txtPayFrequency.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 Freq Name
'           txtPayFrequency.text = flxClientList.TextMatrix(flxClientList.row, 2)
'           FocusControl cmdCommandArray(5)
'           Call FillPayDueDate
        ElseIf strCommandSource = "PayableBasis" Then
            txtPayableBasis.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
            txtPayableBasis.text = flxClientList.TextMatrix(flxClientList.row, 2)
            If txtPayableBasis.Tag = "FA" Then
                txtPercentage.Alignment = vbLeftJustify
                txtPercentage.Locked = True
                txtPercentage.text = "N/A"
                FocusControl txtStopDate
            Else
                txtPercentage.Alignment = vbRightJustify
                txtPercentage.Locked = False
                txtPercentage.text = ""
                FocusControl txtPercentage
            End If
        ElseIf strCommandSource = "ChargingBasis" Then
               txtChargeBasis.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
               txtChargeBasis.text = flxClientList.TextMatrix(flxClientList.row, 2)
               If txtChargeBasis.Tag = "PC" Then 'This can be "Percentage" or "Annual"
    '                txtPercentage.Locked = True
    '                txtPercentage.text = ""
    '                FocusControl txtStopDate
               Else
    '                txtPercentage.Locked = False
    '                txtPercentage.text = ""
    '                FocusControl txtPercentage
               End If
               If txtChargeBasis.text = "Annual" Then
                    Label4(26).Caption = "Amount"
                ElseIf txtChargeBasis.text = "Percentage" Then
                    Label4(26).Caption = "Percentage"
                End If
               FocusControl txtSTART_DATE
        ElseIf strCommandSource = "FrequenciesMngtFee" Then
                txtFrequecymngtFee.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtFrequecymngtFee.text = flxClientList.TextMatrix(flxClientList.row, 2)
                FocusControl txtAmount
                'Here you need to write code for populating the next due date
                If txtSTART_DATE.text <> "N/A" Then
                        If txtSTART_DATE.text <> "" And Trim(txtNtDueDate.text) = "" Then
                            Call NextDueDate(CInt(txtFrequecymngtFee.Tag), txtSTART_DATE, txtNtDueDate, szPropertySelection1)
                            'Now set FDD for this charge type
                            dtFDD = NextDueDate1(CInt(txtFrequecymngtFee.Tag), txtNtDueDate, szPropertySelection1)
                        Else
                            txtComparenextDueDate1 = txtNtDueDate.text
                            Call NextDueDate(CInt(txtFrequecymngtFee.Tag), txtSTART_DATE, txtComparenextDueDate1, szPropertySelection1)
                            If txtComparenextDueDate1 <> txtNtDueDate.text Then
                                    If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
                                              txtNtDueDate = txtComparenextDueDate1
                                    Else
                                             FocusControl txtNtDueDate
                                            txtNtDueDate.SelStart = 0
                                            txtNtDueDate.SelLength = Len(txtNtDueDate)
                                    End If
                            End If
                        End If
                End If
                
                
'                If txtPayableFrom.text <> "" And Trim(txtSCNextDueDt.text) = "" Then
'                        'Do not ask any ques just put the calculated value
'                         NextDueDate txtFreqSC, txtPayableFrom, txtSCNextDueDt, txtSCDemandType.Tag
'                ElseIf txtPayableFrom.text <> "" And Trim(txtSCNextDueDt.text) <> "" Then
'                         txtComparenextDueDate1 = txtSCNextDueDt.text
'                          NextDueDate txtFreqSC, txtPayableFrom, txtComparenextDueDate1, txtSCDemandType.Tag
'                          If txtComparenextDueDate1 <> txtSCNextDueDt.text Then
'                                If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
'                                      txtSCNextDueDt = txtComparenextDueDate1
'                                End If
'                End If
                
        ElseIf strCommandSource = "ManagingAgent" Then
                txtManagingAgentAC.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtManagingAgentAC.text = flxClientList.TextMatrix(flxClientList.row, 2)
                FocusControl cmdCommandArray(10)
        ElseIf strCommandSource = "ChargingMethod" Then
                txtChargingMethod.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtChargingMethod.text = flxClientList.TextMatrix(flxClientList.row, 2) '1 ID, 2 client Name
                 txtChargeBasis.text = ""
                FocusControl cmdCommandArray(11)
                If txtChargingMethod.text = "FIXED" Then
                    txtSTART_DATE.Enabled = True
                    txtSTART_DATE.text = ""
                    txtFrequecymngtFee.Enabled = True
                    txtFrequecymngtFee.text = ""
                    txtNtDueDate.Enabled = True
                    txtNtDueDate.text = ""
                    txtTotalAmountPerYear.text = ""
                    txtChargeBasis.text = ""
                    txtPeriod.Enabled = True
                    txtPeriod.text = ""
                    txtAmount.text = ""
                    txtNtDueDate.Enabled = True
                    txtFrequecymngtFee.Enabled = True
                    txtPeriod.Enabled = True
                    txtTotalAmountPerYear.Enabled = True
                     cmdCommandArray(12).Enabled = True
                ElseIf txtChargingMethod.text = UCase("Receivable") Or txtChargingMethod.text = UCase("RECEIVED") Then
                    txtSTART_DATE.Enabled = False
                    txtSTART_DATE.text = "N/A"
                    txtNtDueDate.text = "N/A"
                    txtNtDueDate.Enabled = False
                    txtFrequecymngtFee.text = "N/A"
                    txtFrequecymngtFee.Enabled = False
                    txtTotalAmountPerYear.text = "N/A"
                    txtPeriod.Enabled = False
                    txtPeriod.text = "N/A"
                    txtAmount.text = ""
                    txtTotalAmountPerYear.Enabled = False
                    cmdCommandArray(12).Enabled = False
'                ElseIf txtChargingMethod.text = UCase("Receivable") Then
'                    txtSTART_DATE.Enabled = False
'                    txtSTART_DATE.text = "N/A"
'                    txtFrequecymngtFee.text = "N/A"
'                    txtFrequecymngtFee.Enabled = False
'                    txtNtDueDate.text = "N/A"
'                    txtNtDueDate.Enabled = False
'                    txtTotalAmountPerYear.Enabled = False
'                    txtTotalAmountPerYear.text = "N/A"
'                    txtAmount.text = ""
'                    txtPeriod.text = "N/A"
'                    txtPeriod.Enabled = False
'                    cmdCommandArray(12).Enabled = False
                End If
        ElseIf strCommandSource = "PaymentType" Then
                txtPaymentType.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtPaymentType.text = flxClientList.TextMatrix(flxClientList.row, 2) '1 ID, 2 client Name
                FocusControl txtPaymentTerms
        ElseIf strCommandSource = "PaymentType1" Then
                txtPaymentMethod.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtPaymentMethod.text = flxClientList.TextMatrix(flxClientList.row, 2) '1 ID, 2 client Name
                FocusControl txtPaymentTerms
                
        ElseIf strCommandSource = "FrequenciesMngtFee" Then
                txtFrequecymngtFee.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID, 2 client Name
                txtFrequecymngtFee.text = flxClientList.TextMatrix(flxClientList.row, 2) '1 ID, 2 client Name
                FocusControl txtTotalAmountPerYear
       ElseIf strCommandSource = "BankFund" Then
'           txtFundCode.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID,2 code,3 fund Name
'           txtFundCode.text = flxClientList.TextMatrix(flxClientList.row, 2) '1 ID,2 code,3 fund Name
'           txtFundName.text = flxClientList.TextMatrix(flxClientList.row, 3)
'           FocusControl cmdAddNewBank
        End If
     Frame5.Visible = False
End Sub
Public Function GetFNCGlobalDataPropertyWise(szPropertyID As String, Conn As ADODB.Connection, szChargeTypeID As String) As Boolean
   'gets the global data from the global data table and puts the payment dates and
   'VAT rate, base rate, number of days to send demands before due, and price per sq foot for
   'service charge and puts then to global variables for when needed by program later.

   'This procedure will be called when program is opened and when the global data is
   'changed.

   Dim i As Integer, iDateSet As Integer
   Dim rst As ADODB.Recordset
   Dim SQLStr As String
   Set rst = New ADODB.Recordset
'   SQLStr = "SELECT PaymentDates FROM ChargeTypes WHERE ID = " & CInt(szChargeTypeID) & ";"
'   Rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic
'
'   iDateSet = Rst.Fields(0).Value
'
'   Rst.Close
'
'   If iDateSet = 255 Then
'      szGDYearly = "AUTOMATIC"
'      Set Rst = Nothing
'      GetFNCGlobalDataPropertyWise = True
'      Exit Function
'   End If
   
   
'   Set Rst = New ADODB.Recordset
   SQLStr = "SELECT * FROM GlobalData WHERE PropertyID = '" & szPropertyID & "';"
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
        MsgBox "Please enter global data for this property", vbInformation, "Warning"
        rst.Close
        Exit Function
   End If
   rst.Close
   SQLStr = "SELECT PropertyID, MDueDate1,MDueDate2,MDueDate3,MDueDate4,MDueDate5,MDueDate6,MDueDate7," & _
            "MDueDate8,MDueDate9,MDueDate10,MDueDate11,MDueDate12,QDueDate1,QDueDate2,QDueDate3,QDueDate4,HYDueDate1,HYDueDate2,YDueDate " & _
            "FROM GlobalData WHERE GlobalData.PropertyID = '" & szPropertyID & "';"
   
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
       ShowMsgInTaskBar "Please complete the set up of the Global Settings for Fees and Charges.", , "N"
       rst.Close
       Set rst = Nothing
       GetFNCGlobalDataPropertyWise = False
       Exit Function
   End If
   If IsNull(rst!YDueDate) Then
        MsgBox "Please complete the set up of the Global Settings for Fees and Charges"
        Exit Function
   End If
   szGDYearly = rst!YDueDate
   szGDHalfYearly1 = rst!HYDueDate1
   szGDHalfYearly2 = rst!HYDueDate2
   szGDQuarterly1 = rst!QDueDate1
   szGDQuarterly2 = rst!QDueDate2
   szGDQuarterly3 = rst!QDueDate3
   szGDQuarterly4 = rst!QDueDate4

   szaMonthlyGD(0) = rst!MDueDate1
   szaMonthlyGD(1) = rst!MDueDate2
   szaMonthlyGD(2) = rst!MDueDate3
   szaMonthlyGD(3) = rst!MDueDate4
   szaMonthlyGD(4) = rst!MDueDate5
   szaMonthlyGD(5) = rst!MDueDate6
   szaMonthlyGD(6) = rst!MDueDate7
   szaMonthlyGD(7) = rst!MDueDate8
   szaMonthlyGD(8) = rst!MDueDate9
   szaMonthlyGD(9) = rst!MDueDate10
   szaMonthlyGD(10) = rst!MDueDate11
   szaMonthlyGD(11) = rst!MDueDate12

   rst.Close
   Set rst = Nothing
   GetFNCGlobalDataPropertyWise = True
End Function

Public Function GetGlobalDataPropertyWise(szPropertyID As String, Conn As ADODB.Connection, szDemandType As String) As Boolean
   'gets the global data from the global data table and puts the payment dates and
   'VAT rate, base rate, number of days to send demands before due, and price per sq foot for
   'service charge and puts then to global variables for when needed by program later.

   'This procedure will be called when program is opened and when the global data is
   'changed.

   Dim i As Integer, iDateSet As Integer
   Dim rst As ADODB.Recordset
   Dim SQLStr As String
   
   Set rst = New ADODB.Recordset
   
   SQLStr = "SELECT PaymentDates FROM DemandTypes WHERE ID = " & CInt(szDemandType) & ";"
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   iDateSet = rst.Fields(0).Value

   rst.Close

   If iDateSet = 255 Then
      szGDYearly = "AUTOMATIC"
      Set rst = Nothing
      GetGlobalDataPropertyWise = True
      Exit Function
   End If

   If iDateSet = 0 Then
      SQLStr = "SELECT PropertyID, TotalArea, TotalSC, SCperSqFoot, SCYearEnd, " & _
                  "BaseInterestRate, VATRate, VATpercentage, QuarterlyDueDate1, " & _
                  "QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate, " & _
                  "NoOfDaysToSendDemandsB4Due, GlobalBankCode, MonthlyDueDate1, " & _
                  "MonthlyDueDate2, MonthlyDueDate3, MonthlyDueDate4, MonthlyDueDate5, " & _
                  "MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, MonthlyDueDate9, " & _
                  "MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12 " & _
               "FROM GlobalData WHERE GlobalData.PropertyID = '" & szPropertyID & "';"
   Else
      SQLStr = "SELECT DateSetID, NameOfSet, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, MonthlyDueDate8, " & _
                  "MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, MonthlyDueDate12, " & _
                  "QuarterlyDueDate1, QuarterlyDueDate2, QuarterlyDueDate3, QuarterlyDueDate4, " & _
                  "HalfYearlyDueDate1, HalfYearlyDueDate2, YearlyDueDate " & _
               "FROM PaymentDates WHERE DateSetID = " & iDateSet & ";"
   End If
   rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
       ShowMsgInTaskBar "You Need to Enter the Global Data.", , "N"
       rst.Close
       Set rst = Nothing
       GetGlobalDataPropertyWise = False
       Exit Function
   End If
   szGDYearly = rst!YearlyDueDate
   szGDHalfYearly1 = rst!HalfYearlyDueDate1
   szGDHalfYearly2 = rst!HalfYearlyDueDate2
   szGDQuarterly1 = rst!QuarterlyDueDate1
   szGDQuarterly2 = rst!QuarterlyDueDate2
   szGDQuarterly3 = rst!QuarterlyDueDate3
   szGDQuarterly4 = rst!QuarterlyDueDate4

   szaMonthlyGD(0) = rst!MonthlyDueDate1
   szaMonthlyGD(1) = rst!MonthlyDueDate2
   szaMonthlyGD(2) = rst!MonthlyDueDate3
   szaMonthlyGD(3) = rst!MonthlyDueDate4
   szaMonthlyGD(4) = rst!MonthlyDueDate5
   szaMonthlyGD(5) = rst!MonthlyDueDate6
   szaMonthlyGD(6) = rst!MonthlyDueDate7
   szaMonthlyGD(7) = rst!MonthlyDueDate8
   szaMonthlyGD(8) = rst!MonthlyDueDate9
   szaMonthlyGD(9) = rst!MonthlyDueDate10
   szaMonthlyGD(10) = rst!MonthlyDueDate11
   szaMonthlyGD(11) = rst!MonthlyDueDate12

   rst.Close
   Set rst = Nothing
   GetGlobalDataPropertyWise = True
End Function


Public Sub NextDueDate(ByVal FrequencyID As Integer, txtTextStart As Control, ByVal txtTextNext As Control, ByVal PROPERTY_ID As String)
   If PROPERTY_ID = "" Then
       MsgBox "You must select a property!", vbOKOnly + vbCritical, "No property Selected"
       Exit Sub
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not GetFNCGlobalDataPropertyWise(PROPERTY_ID, adoConn, txtChargeType.Tag) Then Exit Sub

   adoConn.Close
   Set adoConn = Nothing

  If szGDYearly = "AUTOMATIC" Then
      Select Case FrequencyID
         Case 1:                              'Weekly in advance
            txtTextNext.text = txtTextStart.text
         Case 2:                              'Weekly in arrears
            txtTextNext.text = DateAdd("d", 7, txtTextStart.text)
         Case 3:                              'Fortnightly in advance
            txtTextNext.text = txtTextStart.text
         Case 4:                              'Fortnightly in arrears
            txtTextNext.text = DateAdd("d", 14, txtTextStart.text)
         Case 5:                              'Monthly in advance
            txtTextNext.text = txtTextStart.text
         Case 6:                              'Monthly in arrears
            txtTextNext.text = DateAdd("m", 1, txtTextStart.text)
         Case 7:                              'Quarterly in advance
            txtTextNext.text = txtTextStart.text
         Case 8:                              'Quarterly in arrears
            txtTextNext.text = DateAdd("m", 3, txtTextStart.text)
         Case 9:                              'Half yearly in advance
            txtTextNext.text = txtTextStart.text
         Case 10:                              'Half yearly in arrears
            txtTextNext.text = DateAdd("m", 6, txtTextStart.text)
         Case 11:                             'yearly in advance
            txtTextNext.text = txtTextStart.text
         Case 12:                             'yearly in arrears
            txtTextNext.text = DateAdd("m", 12, txtTextStart.text)
         Case 13:                             'Daily
            txtTextNext.text = ""
         Case 14:                             '4 Weekly in advance
            txtTextNext.text = txtTextStart.text
         Case 15:                             '4 Weekly in arrears
            txtTextNext.text = DateAdd("d", 28, txtTextStart.text)
         Case 16:                             '4 Monrhly in advance
            txtTextNext.text = txtTextStart.text
         Case 17:                             '4 Monrhly in arrears
            txtTextNext.text = DateAdd("m", 4, txtTextStart)
      End Select
Else
         Select Case FrequencyID
            Case 1:                              'Weekly in advance
               txtTextNext.text = txtTextStart.text
            Case 2:                              'Weekly in arrears
               txtTextNext.text = DateAdd("d", 7, txtTextStart.text)
            Case 3:                              'Fortnightly in advance
               txtTextNext.text = txtTextStart.text
            Case 4:                              'Fortnightly in arrears
               txtTextNext.text = DateAdd("d", 14, txtTextStart.text)
            Case 5:                              'Monthly in advance
               txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Monthly)
            Case 6:                              'Monthly in arrears
               txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Monthly)
            Case 7:                              'Quarterly in advance
               txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Quarterly)
            Case 8:                              'Quarterly in arrears
               txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Quarterly)
            Case 9:                              'Half yearly in advance
               txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Half_Yearly)
            Case 10:                              'Half yearly in arrears
               txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Half_Yearly)
            Case 11:                             'yearly in advance
               txtTextNext.text = NextPayingDate(txtTextStart.text, InAdv, Pay_Yearly)
            Case 12:                             'yearly in arrears
               txtTextNext.text = NextPayingDate(txtTextStart.text, InArr, Pay_Yearly)
            Case 13:                             'Daily
               txtTextNext.text = ""
            Case 14:                             '4 Weekly in advance
               txtTextNext.text = txtTextStart.text
            Case 15:                             '4 Weekly in arrears
               txtTextNext.text = DateAdd("d", 28, txtTextStart.text)
            Case 16:                             '4 Monrhly in advance
               txtTextNext.text = txtTextStart.text
            Case 17:                             '4 Monrhly in arrears
               txtTextNext.text = DateAdd("m", 4, txtTextStart)
     End Select
End If
     
End Sub
Public Function NextDueDate1(FrequencyID As Integer, txtTextStart As Control, ByVal PROPERTY_ID As String) As Date
   If PROPERTY_ID = "" Then
       MsgBox "You must select a property!", vbOKOnly + vbCritical, "No property Selected"
       Exit Function
   End If

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not GetFNCGlobalDataPropertyWise(PROPERTY_ID, adoConn, txtChargeType.Tag) Then Exit Function

   adoConn.Close
   Set adoConn = Nothing

  
      Select Case FrequencyID
'         Case 1:                              'Weekly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 2:                              'Weekly in arrears
'            NextDueDate1 = DateAdd("d", 7, txtTextStart.text)
'         Case 3:                              'Fortnightly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 4:                              'Fortnightly in arrears
'            NextDueDate1 = DateAdd("d", 14, txtTextStart.text)
'         Case 5:                              'Monthly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 6:                              'Monthly in arrears
'            NextDueDate1 = DateAdd("m", 1, txtTextStart.text)
'         Case 7:                              'Quarterly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 8:                              'Quarterly in arrears
'            NextDueDate1 = DateAdd("m", 3, txtTextStart.text)
'         Case 9:                              'Half yearly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 10:                              'Half yearly in arrears
'            NextDueDate1 = DateAdd("m", 6, txtTextStart.text)
'         Case 11:                             'yearly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 12:                             'yearly in arrears
'            NextDueDate1 = DateAdd("m", 12, txtTextStart.text)
'         Case 13:                             'Daily
'            NextDueDate1 = ""
'         Case 14:                             '4 Weekly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 15:                             '4 Weekly in arrears
'            NextDueDate1 = DateAdd("d", 28, txtTextStart.text)
'         Case 16:                             '4 Monrhly in advance
'            NextDueDate1 = txtTextStart.text
'         Case 17:                             '4 Monrhly in arrears
'            NextDueDate1 = DateAdd("m", 4, txtTextStart)
            Case 1:                              'Weekly in advance
               NextDueDate1 = txtTextStart.text
            Case 2:                              'Weekly in arrears
              NextDueDate1 = DateAdd("d", 7, txtTextStart.text)
            Case 3:                              'Fortnightly in advance
               NextDueDate1 = txtTextStart.text
            Case 4:                              'Fortnightly in arrears
               NextDueDate1 = DateAdd("d", 14, txtTextStart.text)
            Case 5:                              'Monthly in advance
               NextDueDate1 = NextPayingDate(txtTextStart.text, InAdv, Pay_Monthly)
            Case 6:                              'Monthly in arrears
               NextDueDate1 = NextPayingDate(txtTextStart.text, InArr, Pay_Monthly)
            Case 7:                              'Quarterly in advance
               NextDueDate1 = NextPayingDate(txtTextStart.text, InAdv, Pay_Quarterly)
            Case 8:                              'Quarterly in arrears
               NextDueDate1 = NextPayingDate(txtTextStart.text, InArr, Pay_Quarterly)
            Case 9:                              'Half yearly in advance
               NextDueDate1 = NextPayingDate(txtTextStart.text, InAdv, Pay_Half_Yearly)
            Case 10:                              'Half yearly in arrears
               NextDueDate1 = NextPayingDate(txtTextStart.text, InArr, Pay_Half_Yearly)
            Case 11:                             'yearly in advance
               NextDueDate1 = NextPayingDate(txtTextStart.text, InAdv, Pay_Yearly)
            Case 12:                             'yearly in arrears
               NextDueDate1 = NextPayingDate(txtTextStart.text, InArr, Pay_Yearly)
            Case 13:                             'Daily
               NextDueDate1 = ""
            Case 14:                             '4 Weekly in advance
               NextDueDate1 = txtTextStart.text
            Case 15:                             '4 Weekly in arrears
               NextDueDate1 = DateAdd("d", 28, txtTextStart.text)
            Case 16:                             '4 Monrhly in advance
               NextDueDate1 = txtTextStart.text
            Case 17:                             '4 Monrhly in arrears
               NextDueDate1 = DateAdd("m", 4, txtTextStart)
      End Select

     
End Function
Private Sub buttonmmodeOnGridClick()
    cmdAgrTopSave.Enabled = False
    tabAgreement.Enabled = True
    cmdAgmntAddNew.Enabled = True
    cmdAgrTopEdit.Enabled = True
'    cboDmdPropertyList.Enabled = True
'    cmdGSEdit.Enabled = True
    cmdAddEditBankCode.Enabled = False
    cmdAddNewBank.Enabled = True
    
'    cmdEditBank.Enabled = True
End Sub

Private Sub LoadFormByClient()
   Dim sSQLQuery_ As String, sFilter As String

   txtClientID.text = LOAD_CLINT_CLIENTID

  ' MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh
   Dim adoConn As New ADODB.Connection
   Dim rsClient As New ADODB.Recordset
   adoConn.Open getConnectionString
'   sSQLQuery_ = "SELECT ClientID, ClientName, ClientAddressLine1, ClientAddressLine2, " & _
'                     "ClientAddressLine3, ClientPostCode, ClientOfficeEmail, ClientPersonalEmail, " & _
'                     "ClientHomeTel, ClientMobile, ClientOfficeAddressLine1, " & _
'                     "ClientOfficeAddressLine2, ClientOfficeAddressLine3, ClientOfficePostCode, " & _
'                     "ClientOfficeTel, ClientMemo, LandLordSageCustAC, LandLordSageSuppAC, " & _
'                     "BANK_ID, CommissionType, CommissionAmt, BGRPayable, VATReg, AcBalance, " & _
'                     "Residency, YearEndDate, PaymentMethod, BacsRef, HomeOfficeAdd " & _
'                "FROM  CLIENT " & _
'                "WHERE CLIENT.ClientID = '" & LOAD_CLINT_CLIENTID & "';"
    sSQLQuery_ = "SELECT ClientID, ClientName, ClientAddressLine1, ClientAddressLine2, ClientAddressLine4, " & _
                  "ClientAddressLine3, ClientPostCode, ClientOfficeEmail, ClientPersonalEmail, " & _
                  "ClientHomeTel, ClientMobile, '', ClientOfficeAddressLine1, ClientOfficeAddressLine4, " & _
                  "ClientOfficeAddressLine2, ClientOfficeAddressLine3, ClientOfficePostCode, " & _
                  "ClientOfficeTel, ClientMemo, LandLordSageCustAC, LandLordSageSuppAC, " & _
                  "BANK_ID, CommissionType, CommissionAmt, BGRPayable, VATReg, AcBalance, " & _
                  "Residency, YearEndDate, PaymentMethod, BacsRef, HomeOfficeAdd, CompReg, RegAdd1, " & _
                  "RegAdd2, RegAdd3, RegAdd4, RegPostCode, CT " & _
                "FROM CLIENT " & _
                "WHERE CLIENT.ClientID = '" & txtClientID.text & "';"

'   adoMain.RecordSource = sSQLQuery_
'   adoMain.CommandType = adCmdText
'   adoMain.Refresh
        rsClient.Open sSQLQuery_, adoConn, adOpenKeyset, adLockOptimistic

   If Not Fill_Form() Then 'If Not Fill_Form(Me, adoMain) Then Mark 1
      MsgBox "Error in Database.", vbExclamation
   Else
      LoadClientProperty

      lblLoading.Caption = "Please wait, tree is building..."
      fmeLoading.Refresh
      Debug.Print time
      DrawLandLordTree tvwLandLord, imgList, txtClientID.text, True, NEW_TYPE
      Debug.Print time
      
       Dim rsClientGlobalData As New ADODB.Recordset
      lblLoading.Caption = "Please wait, global data is loading..."

      fmeLoading.Refresh

      sSQLQuery_ = "SELECT Record_ID, ClientID, QuarterlyDueDate1, QuarterlyDueDate2, " & _
                  "QuarterlyDueDate3, QuarterlyDueDate4, HalfYearlyDueDate1, " & _
                  "HalfYearlyDueDate2, MonthlyDueDate1, MonthlyDueDate2, MonthlyDueDate3, " & _
                  "MonthlyDueDate4, MonthlyDueDate5, MonthlyDueDate6, MonthlyDueDate7, " & _
                  "MonthlyDueDate8, MonthlyDueDate9, MonthlyDueDate10, MonthlyDueDate11, " & _
                  "MonthlyDueDate12, YearlyDueDate, FeeIsuDays, PayIsuDays, LettingFee, " & _
                  "LettingAM, LettingFreq, LettingNtDueDt, LettingStDt, LettingChrgType, " & _
                  "MngFee, MngAM, MngFreq, MngNtDueDt, MngStDt, MngChrgType, RentPayble, " & _
                  "RentAM, RentFreq, RentNtDueDt, RentStDt, RentChrgType " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"
      rsClientGlobalData.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
      
      If rsClientGlobalData.RecordCount > 0 Then
         bGlobalData = True
      Else
         bGlobalData = False
      End If
'      RetrieveMemo "Client", "ClientMemo", txtClientID.text, "ClientID", txtNote
   End If

   fmeLoading.Visible = False
   MousePointer = vbDefault

   Frame5.Visible = False
End Sub
Private Sub LoadClientAddressLins()
    Dim sSQLQuery_  As String
    Dim adoConn As New ADODB.Connection
    Dim rsClient As New ADODB.Recordset
    If txtClientID.text = "" Then Exit Sub
    adoConn.Open getConnectionString
    If (NEW_TYPE <> "Landlord") Then
      sSQLQuery_ = "SELECT ClientID, ClientName, ClientAddressLine1, ClientAddressLine2,ClientAddressLine5,ClientOfficeAddressLine5,RegAdd5, " & _
                     "ClientAddressLine3, ClientPostCode, ClientOfficeEmail, ClientPersonalEmail, " & _
                     "ClientHomeTel, ClientMobile, '', ClientOfficeAddressLine1, " & _
                     "ClientOfficeAddressLine2, ClientOfficeAddressLine3, ClientOfficePostCode, " & _
                     "ClientOfficeTel, ClientMemo, LandLordSageCustAC, LandLordSageSuppAC, " & _
                     "BANK_ID, CommissionType, CommissionAmt, BGRPayable, VATReg, AcBalance, " & _
                     "Residency, YearEndDate, PaymentMethod, BacsRef, HomeOfficeAdd, CompReg, RegAdd1, " & _
                     "RegAdd2, RegAdd3, RegPostCode, CT, ClientAddressLine4, ClientOfficeAddressLine4,  " & _
                     "RegAdd4,groupCode,Comments1,Comments2,stClientHomeTel,stClientOfficeTel,stClientMobile ,stClientPersonalEmail,stClientOfficeEmail  " & _
                   "FROM CLIENT " & _
                   "WHERE CLIENT.ClientID = '" & txtClientID.text & "';"
   End If
'Debug.Print sSQLQuery_
    rsClient.Open sSQLQuery_, adoConn, adOpenKeyset, adLockReadOnly
    
'   adoMain.RecordSource = sSQLQuery_
'   adoMain.CommandType = adCmdText
'   adoMain.Refresh


      If rsClient.EOF Then
         'MsgBox "Error in Database.", vbExclamation
         rsClient.Close
         Set rsClient = Nothing
         Exit Sub
      End If

      txtClientAddressLine1(0).text = IIf(IsNull(rsClient.Fields(2).Value), "", rsClient.Fields(2).Value)
      txtClientAddressLine1(1).text = IIf(IsNull(rsClient.Fields(3).Value), "", rsClient.Fields(3).Value)
      txtClientAddressLine1(2).text = IIf(IsNull(rsClient.Fields(4).Value), "", rsClient.Fields(4).Value)
      txtClientAddressLine1(3).text = IIf(IsNull(rsClient.Fields("ClientAddressLine4").Value), "", rsClient.Fields("ClientAddressLine4").Value)
      txtClientAddressLine1(5).text = IIf(IsNull(rsClient.Fields(5).Value), "", rsClient.Fields(5).Value)
      
      txtClientHomeTel(0).text = IIf(IsNull(rsClient.Fields(8).Value), "", rsClient.Fields(8).Value)
      txtClientHomeTel(1).text = IIf(IsNull(rsClient.Fields(15).Value), "", rsClient.Fields(15).Value)
      txtClientHomeTel(2).text = IIf(IsNull(rsClient.Fields(9).Value), "", rsClient.Fields(9).Value)
      txtClientHomeTel(3).text = IIf(IsNull(rsClient.Fields(7).Value), "", rsClient.Fields(7).Value)
      txtClientHomeTel(4).text = IIf(IsNull(rsClient.Fields(6).Value), "", rsClient.Fields(6).Value)
      txtClientHomeTel(5).text = IIf(IsNull(rsClient.Fields("groupCode").Value), "", rsClient.Fields("groupCode").Value)
      
      'You need to load lined for the additional statement lines for the statement
        txtClientHomeTel(6).text = IIf(IsNull(rsClient!StClientHomeTel), "", rsClient!StClientHomeTel) 'rsClient!ClientHomeTel
        txtClientHomeTel(7).text = IIf(IsNull(rsClient!StClientOfficeTel), "", rsClient!StClientOfficeTel) 'rsClient!ClientOfficeTel
        txtClientHomeTel(8).text = IIf(IsNull(rsClient!StClientMobile), "", rsClient!StClientMobile) 'rsClient!ClientMobile
        txtClientHomeTel(9).text = IIf(IsNull(rsClient!StClientPersonalEmail), "", rsClient!StClientPersonalEmail) 'rsClient!ClientPersonalEmail
        txtClientHomeTel(10).text = IIf(IsNull(rsClient!StClientOfficeEmail), "", rsClient!StClientOfficeEmail) ' rsClient!ClientOfficeEm
    
      txtClientHomeTel(15).text = IIf(IsNull(rsClient.Fields(11).Value), "", rsClient.Fields(11).Value)
      txtClientHomeTel(16).text = IIf(IsNull(rsClient.Fields(12).Value), "", rsClient.Fields(12).Value)
      txtClientHomeTel(17).text = IIf(IsNull(rsClient.Fields(13).Value), "", rsClient.Fields(13).Value)
      txtClientHomeTel(18).text = IIf(IsNull(rsClient.Fields("ClientOfficeAddressLine4").Value), "", rsClient.Fields("ClientOfficeAddressLine4").Value)
      txtClientHomeTel(19).text = IIf(IsNull(rsClient.Fields("ClientOfficeAddressLine5").Value), "", rsClient.Fields("ClientOfficeAddressLine5").Value)
      txtClientHomeTel(20).text = IIf(IsNull(rsClient.Fields(14).Value), "", rsClient.Fields(14).Value)
      
'      txtNote.text = IIf(IsNull(rsCLIENT.Fields(16).Value), "", rsCLIENT.Fields(16).Value)
      'txtResidency.text = IIf(IsNull(rsCLIENT.Fields(25).Value), "", rsCLIENT.Fields(25).Value)
'      txtAcBalance(0).text = IIf(IsNull(rsClient.Fields(24).Value), "", rsClient.Fields(24).Value)
      txtVATReg.text = IIf(IsNull(rsClient.Fields(23).Value), "", rsClient.Fields(23).Value)
      txtYearEndDate.text = IIf(IsNull(rsClient.Fields(26).Value), "", rsClient.Fields(26).Value)
      txtClientHomeTel(21).text = IIf(IsNull(rsClient.Fields.Item("CompReg").Value), "", rsClient.Fields.Item("CompReg").Value)
      txtClientHomeTel(22).text = IIf(IsNull(rsClient.Fields.Item("RegAdd1").Value), "", rsClient.Fields.Item("RegAdd1").Value)
      txtClientHomeTel(23).text = IIf(IsNull(rsClient.Fields.Item("RegAdd2").Value), "", rsClient.Fields.Item("RegAdd2").Value)
      txtClientHomeTel(24).text = IIf(IsNull(rsClient.Fields.Item("RegAdd3").Value), "", rsClient.Fields.Item("RegAdd3").Value)
      txtClientHomeTel(25).text = IIf(IsNull(rsClient.Fields.Item("RegAdd4").Value), "", rsClient.Fields.Item("RegAdd4").Value)
      txtClientHomeTel(26).text = IIf(IsNull(rsClient.Fields.Item("RegAdd5").Value), "", rsClient.Fields.Item("RegAdd5").Value)
      txtClientHomeTel(27).text = IIf(IsNull(rsClient.Fields.Item("RegPostCode").Value), "", rsClient.Fields.Item("RegPostCode").Value)
     
End Sub
Private Sub LoadPaymentDatesFromGlobalData(Optional szProp As String)
   Dim szTemp() As String
   Dim i        As Integer
   Dim szSQL    As String
   Dim rsGlobalData As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   szSQL = "SELECT PropertyID, QDueDate1,QDueDate2,QDueDate3,QDueDate4,HYDueDate1,HYDueDate2,YDueDate,NoOfDaysToSendMFB4Due,MDueDate1,MDueDate2," & _
           "MDueDate3,MDueDate4,MDueDate5,MDueDate6,MDueDate7,MDueDate8,MDueDate9,MDueDate10,MDueDate11,MDueDate12 " & _
           "FROM GlobalData " & _
           "WHERE PropertyID = '" & szPropertySelection2 & "' "
  
   rsGlobalData.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
   bGlobalData = IIf(rsGlobalData.RecordCount > 0, True, False)

   If bGlobalData Then
            If IsNull(rsGlobalData("QDueDate1").Value) Or rsGlobalData("QDueDate1").Value = "" Then
                Exit Sub
            End If
        
            If IsNull(rsGlobalData!NoOfDaysToSendMFB4Due) = False Then txtNoOfDaysToSendMFB4Due.text = rsGlobalData!NoOfDaysToSendMFB4Due
            If txtNoOfDaysToSendMFB4Due.text = "0" Then
                txtNoOfDaysToSendMFB4Due.text = ""
            End If
            cboDay(31).text = Left(rsGlobalData!MDueDate1, 2)
            cboDay(32).text = Left(rsGlobalData!MDueDate2, 2)
            cboDay(33).text = Left(rsGlobalData!MDueDate3, 2)
            cboDay(34).text = Left(rsGlobalData!MDueDate4, 2)
            cboDay(35).text = Left(rsGlobalData!MDueDate5, 2)
            cboDay(36).text = Left(rsGlobalData!MDueDate6, 2)
            cboDay(37).text = Left(rsGlobalData!MDueDate7, 2)
            cboDay(38).text = Left(rsGlobalData!MDueDate8, 2)
            cboDay(39).text = Left(rsGlobalData!MDueDate9, 2)
            cboDay(40).text = Left(rsGlobalData!MDueDate10, 2)
            cboDay(41).text = Left(rsGlobalData!MDueDate11, 2)
            cboDay(42).text = Left(rsGlobalData!MDueDate12, 2)
            
            cboQDay(11).text = Left(rsGlobalData!QDueDate1, 2)
            cboQMth(11).text = Right(rsGlobalData!QDueDate1, Len(rsGlobalData!QDueDate1) - 3)
            cboQDay(12).text = Left(rsGlobalData!QDueDate2, 2)
            cboQMth(12).text = Right(rsGlobalData!QDueDate2, Len(rsGlobalData!QDueDate2) - 3)
            cboQDay(13).text = Left(rsGlobalData!QDueDate3, 2)
            cboQMth(13).text = Right(rsGlobalData!QDueDate3, Len(rsGlobalData!QDueDate3) - 3)
            cboQDay(14).text = Left(rsGlobalData!QDueDate4, 2)
            cboQMth(14).text = Right(rsGlobalData!QDueDate4, Len(rsGlobalData!QDueDate4) - 3)
            
            cboHDay(5).text = Left(rsGlobalData!HYDueDate1, 2)
            cboHMth(5).text = Right(rsGlobalData!HYDueDate1, Len(rsGlobalData!HYDueDate1) - 3)
            cboHDay(6).text = Left(rsGlobalData!HYDueDate2, 2)
            cboHMth(6).text = Right(rsGlobalData!HYDueDate2, Len(rsGlobalData!HYDueDate2) - 3)
            
            cboYDay(2).text = Left(rsGlobalData!YDueDate, 2)
            cboYMth(2).text = Right(rsGlobalData!YDueDate, Len(rsGlobalData!YDueDate) - 3)
       End If
   
End Sub
Private Sub ClearPaymentDates()
            txtNoOfDaysToSendMFB4Due.text = ""
            cboDay(31).text = ""
            cboDay(32).text = ""
            cboDay(33).text = ""
            cboDay(34).text = ""
            cboDay(35).text = ""
            cboDay(36).text = ""
            cboDay(37).text = ""
            cboDay(38).text = ""
            cboDay(39).text = ""
            cboDay(40).text = ""
            cboDay(41).text = ""
            cboDay(42).text = ""
            
            cboQDay(11).text = ""
            cboQMth(11).text = ""
            cboQDay(12).text = ""
            cboQMth(12).text = ""
            cboQDay(13).text = ""
            cboQMth(13).text = ""
            cboQDay(14).text = ""
            cboQMth(14).text = ""
            
            cboHDay(5).text = ""
            cboHMth(5).text = ""
            cboHDay(6).text = ""
            cboHMth(6).text = ""
            
            cboYDay(2).text = ""
            cboYMth(2).text = ""
End Sub
Private Sub LoadClientProperty()
   Dim conClient As New ADODB.Connection
   Dim rstProperty As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'cboProperty.Clear

   'Set the RDO Connections to the dataset
   conClient.Open getConnectionString
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt

   If (NEW_TYPE = "Landlord") Then
      szSQL = "SELECT PROPERTY.PropertyID, PROPERTY.PropertyName  " & _
           "FROM PROPERTY, PROPERTYLANDLORD, LANDLORD " & _
           "WHERE PROPERTY.PROPERTYID = PROPERTYLANDLORD.PROPERTYID " & _
           "AND PROPERTYLANDLORD.LANDLORDID=LANDLORD.LANDLORDID " & _
           "AND LANDLORD.LANDLORDID = '" & txtClientID.text & "' " & _
           "ORDER BY PROPERTY.PropertyName;"
   Else
      szSQL = "SELECT PropertyID, PropertyName  " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyName;"
   End If

   rstProperty.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   If rstProperty.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstProperty.EOF
     'cboProperty.AddItem rstProperty!propertyID & " / " & rstProperty!PropertyName
      rstProperty.MoveNext
   Wend

NoRes:
   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
End Sub

Private Sub flxClientList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClientList_Click
    End If
End Sub

Private Sub flxOtherBankDetails_Click()
    HighLightRowFlxGrid flxOtherBankDetails, flxOtherBankDetails.row
    cmdDeleteBank.Enabled = True
    EnableDisableAcText True
    flxOtherBankDetails_RowColChange
End Sub

Private Sub flxOtherBankDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxOtherBankDetails.ToolTipText = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.MouseRow, flxOtherBankDetails.MouseCol)
End Sub

Private Sub flxOtherBankDetails_RowColChange()
   Dim iCol As Integer
'   cmdEditBank.Enabled = True
   cmdEdit.Enabled = True
   iSlectedRow = flxOtherBankDetails.row
   'HighLightRowFlxGrid flxOtherBankDetails, flxOtherBankDetails.row
   'Below line has been added by anol 11 Mar 2015 ,It was giving error when row is 0
   If flxOtherBankDetails.row = 0 Then Exit Sub
   'end of modification
   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 10)
   txtPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)
   If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14) <> "" Then
      txtNCCODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14)
      txtNominal.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 18)
   Else
      txtNCCODE.text = ""
      txtNominal.text = ""
   End If

   chkOverDraft.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 15) = "YES", 1, 0)
   txtOverDraft.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 16)
   'Below line added by anol 04 Apr 2015 implementation of consolidated.
   chkConsolidated.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 17) = "YES", 1, 0)
  
   
   If chkConsolidated.Value = 1 Then
        txtconsolidatedAccountName.text = returnConBanName(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 19))
        If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 19) = "" Then
        Else
        
            intConsolidatedBankID = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 19)
          End If
   Else
         txtconsolidatedAccountName.text = ""
         txtBank_AC_Name.Visible = True
   End If
   'End of modification
   
   Call updateBankBalance
   If txtNCCODE.text <> "" Then
        'all funds are loaded. now clear all flagging and mark the same what is in the BankFundTable
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Dim rsBankFund As New ADODB.Recordset
        Dim iRow As Long
        For iRow = 1 To flxBankAccountFund.Rows - 1
            flxBankAccountFund.TextMatrix(iRow, 0) = ""
        Next iRow
        rsBankFund.Open "Select * from BankFund where clientID='" & txtClientID.text & "' AND BankCode='" & txtNCCODE.text & "'", adoConn, adOpenStatic, adLockReadOnly
        While Not rsBankFund.EOF
            For iRow = 1 To flxBankAccountFund.Rows - 1
                If flxBankAccountFund.TextMatrix(iRow, 1) = rsBankFund("FundID").Value Then
                    flxBankAccountFund.TextMatrix(iRow, 0) = "X"
                End If
            Next iRow
            rsBankFund.MoveNext
        Wend
        rsBankFund.Close
        adoConn.Close
   End If
End Sub
Private Sub updateBankBalance()
        Dim adoConn As New ADODB.Connection
        Dim adoRST As New ADODB.Recordset
        adoConn.Open getConnectionString
        Dim Balance As Double
        Dim szSQL As String
   ' find current Balance for the selected bank account and selected client ID by anol 2023-05-24
   szSQL = " SELECT sum(SWITCH(T ='3',AMT,T ='4',AMT,T ='8',-AMT,T ='9',-AMT,T ='BP',-AMT,T ='BR',AMT,T ='23',-AMT,T ='24',AMT)) as AMTT from (" & _
            "SELECT SUM(R.Amount) AS AMT, Type AS T " & _
           "FROM tlbReceipt AS R, tlbTransactionTypes AS TT, Units AS U, Property AS P, tlbClientBanks AS B " & _
           "WHERE (R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                  "TT.TYPE_ID = R.Type AND R.BankCode = '" & txtNCCODE & "' AND U.UnitNumber = R.UnitID AND " & _
                  "U.PropertyID = P.PropertyID AND P.ClientID = '" & txtClientID & "' AND B.NominalCode = R.BankCode AND " & _
                  "B.CLIENT_ID = P.ClientID group by Type UNION "
                  
        szSQL = szSQL & _
                "SELECT SUM(BP.NET_AMOUNT + BP.VAT) AS AMT, TRANS AS T " & _
                "FROM tlbBankPayment AS BP, tlbTransactionTypes AS TT, tlbClientBanks AS B " & _
                "WHERE (BP.TransactionType = 11 OR BP.TransactionType = 12) AND " & _
                       "BP.BANK_AC = '" & txtNCCODE & "' AND BP.TransactionType = TT.TYPE_ID AND " & _
                       "BP.ClientID = '" & txtClientID & "' AND B.NominalCode = BP.BANK_AC AND B.CLIENT_ID = BP.ClientID  group by TRANS UNION "
        szSQL = szSQL & _
                "SELECT SUM(P.Amount) AS AMT, Type AS T " & _
                "FROM tlbPayment AS P, tlbTransactionTypes AS TT " & _
                "WHERE (P.Type = 8 OR P.Type = 9 OR P.Type = 24) AND P.BankCode = '" & txtNCCODE & "' AND P.Type = TT.TYPE_ID AND " & _
                       "P.ClientID = '" & txtClientID & "'   group by Type )"
                       
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
      txtBankBalance.text = IIf(IsNull(adoRST.Fields.Item("AMTT").Value), 0, adoRST.Fields.Item("AMTT").Value)
      txtBankBalance.text = Format(txtBankBalance.text, "0.00")
   End If
   adoRST.Close
                       
    szSQL = "Select sum(amount) as DAmt from RetentionDetails where isDeleted=false and BankCode='" & txtNCCODE & "' and ClientID='" & txtClientID & "'"
    adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    If Not adoRST.EOF Then
        txtRetention.text = IIf(IsNull(adoRST.Fields.Item("DAmt").Value), 0, adoRST.Fields.Item("DAmt").Value)
        txtRetention.text = Format(txtRetention.text, "0.00")
    End If
    adoRST.Close
    txtAvailableBankBalance.text = Val(txtBankBalance.text) + Val(txtOverDraft) - Val(txtRetention.text)
     txtAvailableBankBalance.text = Format(txtAvailableBankBalance.text, "0.00")
End Sub
Private Function returnConBanName(strID As String) As String
    If strID = "" Then Exit Function
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsConBank As New ADODB.Recordset
    rsConBank.Open "Select BankName from ConsolidatedBankList where conBankID=" & strID & "", adoConn, adOpenStatic, adLockReadOnly
    If Not rsConBank.EOF Then
        returnConBanName = rsConBank("BankName").Value
    End If
    rsConBank.Close
    adoConn.Close
    Set adoConn = Nothing
    
End Function
Private Function returnConBanID(strBankName As String) As String
    If strBankName = "" Then Exit Function
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Dim rsConBank As New ADODB.Recordset
    rsConBank.Open "Select conBankID from ConsolidatedBankList where BankName='" & strBankName & "'", adoConn, adOpenStatic, adLockReadOnly
    If Not rsConBank.EOF Then
        returnConBanID = rsConBank("conBankID").Value
    End If
'    If returnConBanID = "" Then
'            returnConBanID = Null
'    End If
    rsConBank.Close
    adoConn.Close
    Set adoConn = Nothing
    
End Function
Private Sub restoreSelectedValueflxOtherBankDetails()
   Dim iCol As Integer
'   cmdEditBank.Enabled = True
   If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1) <> "" Then
        'flxOtherBankDetails.row = 1
        HighLightRowFlxGrid flxOtherBankDetails, flxOtherBankDetails.row
'         If cmdEditBank.Enabled = False Then
'            cmdDeleteBank.Enabled = True
'            cmdBACS.Enabled = True
'        End If
        EnableDisableAcText True
   Else
        Exit Sub
   End If
   iSlectedRow = flxOtherBankDetails.row
   HighLightRowFlxGrid flxOtherBankDetails, flxOtherBankDetails.row
   'Below line has been added by anol 11 Mar 2015 ,It was giving error when row is 0
   If flxOtherBankDetails.row = 0 Then Exit Sub
   'end of modification
   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5)
   txtconsolidatedAccountName.text = returnConBanName(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 19))
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 10)
   txtPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)
'   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
'   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14) <> "" Then
      txtNCCODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14)
      txtNominal.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 18)
   Else
      txtNCCODE.text = ""
      txtNominal.text = ""
   End If
   chkOverDraft.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 15) = "YES", 1, 0)
   txtOverDraft.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 16)
   'Below line added by anol 04 Apr 2015 implementation of consolidated.
   chkConsolidated.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 17) = "YES", 1, 0)
   'End of modification
  
End Sub
Private Sub LoadFirstValueflxOtherBankDetails()
   Dim iCol As Integer
'   cmdEditBank.Enabled = True
   If flxOtherBankDetails.TextMatrix(1, 1) <> "" Then
        flxOtherBankDetails.row = 1
        HighLightRowFlxGrid flxOtherBankDetails, 1
'         If cmdEditBank.Enabled = False Then
'            cmdDeleteBank.Enabled = True
'            cmdBACS.Enabled = True
'        End If
        EnableDisableAcText True
   Else
        Exit Sub
   End If
   iSlectedRow = flxOtherBankDetails.row
   HighLightRowFlxGrid flxOtherBankDetails, flxOtherBankDetails.row
   'Below line has been added by anol 11 Mar 2015 ,It was giving error when row is 0
   If flxOtherBankDetails.row = 0 Then Exit Sub
   'end of modification
   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 10)
   txtPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 11)
'   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
'   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   If flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14) <> "" Then
      txtNCCODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 14)
      txtNominal.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 18)
   Else
      txtNCCODE.text = ""
      txtNominal.text = ""
   End If
   chkOverDraft.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 15) = "YES", 1, 0)
   txtOverDraft.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 16)
   'Below line added by anol 04 Apr 2015 implementation of consolidated.
   chkConsolidated.Value = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 17) = "YES", 1, 0)
   If chkConsolidated.Value = 1 Then
        txtconsolidatedAccountName.text = returnConBanName(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.row, 19))
'        cboconsolidatedAccountName.Visible = True
'        txtBank_AC_Name.Visible = False
   Else
'         cboconsolidatedAccountName.Visible = False
         'cboconsolidatedAccountName.text = ""
         txtBank_AC_Name.Visible = True
   End If
   
   
   If txtNCCODE.text <> "" Then
        'all funds are loaded. now clear all flagging and mark the same what is in the BankFundTable
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        Dim rsBankFund As New ADODB.Recordset
        Dim iRow As Long
        For iRow = 1 To flxBankAccountFund.Rows - 1
            flxBankAccountFund.TextMatrix(iRow, 0) = ""
        Next iRow
        rsBankFund.Open "Select * from BankFund where clientID='" & txtClientID.text & "' AND BankCode='" & txtNCCODE.text & "'", adoConn, adOpenStatic, adLockReadOnly
        While Not rsBankFund.EOF
            For iRow = 1 To flxBankAccountFund.Rows - 1
                If flxBankAccountFund.TextMatrix(iRow, 1) = rsBankFund("FundID").Value Then
                    flxBankAccountFund.TextMatrix(iRow, 0) = "X"
                End If
            Next iRow
            rsBankFund.MoveNext
        Wend
        rsBankFund.Close
        adoConn.Close
   End If
   Call updateBankBalance
    'rstRec.Close
'    adoConn.Close
       ' Set rstRec = Nothing
'       Set adoConn = Nothing
       
   'End of modification
  
End Sub



Private Sub flxPayable_displaytextboxes()
        Dim rRow As Integer
        rRow = flxPayable.row
        If rRow = 0 Then Exit Sub
        If flxPayable.TextMatrix(rRow, 1) = "" Then Exit Sub
        strtlbPayableID = flxPayable.TextMatrix(rRow, 1) ' = rsPayable.Fields.Item("PAYABLE_ID").Value
        txtPayableType.Tag = flxPayable.TextMatrix(rRow, 3) ' = rsPayable.Fields.Item("CPA_ID").Value
        txtPayableType.text = flxPayable.TextMatrix(rRow, 4)      '= rsPayable.Fields.Item("PAYABLE_TYPE").Value
'        txtPayDemandType.Tag = flxPayable.TextMatrix(rRow, 5)    '= rsPayable.Fields.Item("DemandID").Value
'        txtPayDemandType.text = flxPayable.TextMatrix(rRow, 6)   '= rsPayable.Fields.Item("PAY_DEMAND_TYPE").Value
        txtPayFund.Tag = flxPayable.TextMatrix(rRow, 7)  '= rsPayable.Fields.Item("FundID").Value
        txtPayFund.text = flxPayable.TextMatrix(rRow, 8)  '= rsPayable.Fields.Item("FundName").Value
        txtPayeeType.text = flxPayable.TextMatrix(rRow, 9)  ' = rsPayable.Fields.Item("clientLandlordID").Value
        txtClientLandlord.text = flxPayable.TextMatrix(rRow, 10)  ' = rsPayable.Fields.Item("clientLandlordID").Value
'        txtPAY_START_DATE.text = Format(flxPayable.TextMatrix(rRow, 10), "dd/mm/yyyy")  ' = rsPayable.Fields.Item("PAY_START_DATE").Value
'        txtPayFrequency.Tag = flxPayable.TextMatrix(rRow, 11)   '= rsPayable.Fields.Item("FREQID").Value
'        txtPayFrequency.text = flxPayable.TextMatrix(rRow, 12)   '= rsPayable.Fields.Item("Frequency").Value
'        chkONDD.Value = IIf(flxPayable.TextMatrix(rRow, 13) = False, 0, 1)
'        txtPAY_NtDueDate.text = Format(flxPayable.TextMatrix(rRow, 14), "dd/mm/yyyy") '= rsPayable.Fields.Item("PAY_NtDueDate").Value
        txtPayableBasis.text = IIf(flxPayable.TextMatrix(rRow, 14) = "FA", "Full Amount", "Percentage")   '= IIf(rsPayable.Fields.Item("PAYABLE_BASIS_ ").Value = "FA", "Total Amount", "Percentage")
        If txtPayableBasis.text = "Full Amount" Then
                txtPercentage.Alignment = vbLeftJustify
          Else
                txtPercentage.Alignment = vbRightJustify
          End If
        txtPercentage.text = flxPayable.TextMatrix(rRow, 16)  'Percentage
        txtStopDate.text = IIf(flxPayable.TextMatrix(rRow, 17) = "", "", Format(flxPayable.TextMatrix(rRow, 17), "dd/mm/yyyy")) ' = rsPayable.Fields.Item("StopDate").Value
'        txtPAY_END_DATE.text = IIf(flxPayable.TextMatrix(rRow, 18) = "", "", Format(flxPayable.TextMatrix(rRow, 18), "dd/mm/yyyy")) ' = rsPayable.Fields.Item("PAY_END_DATE").Value
        cmdPayEdit.Enabled = True
        cmdDeleteRentPayable.Enabled = True
End Sub

Private Sub flxPayable_RowColChange()
'   populateControl Me, flxPayable
'   cmdPayEdit.Enabled = True
End Sub

Private Sub flxPropertySelection1_Click()
    If AGREEMENT_ADDNEW_MODE = False Then 'warning for exiting the mode
         If cmdAgmntSave.Enabled Then
            cmdAgmntSave_Click
            AGREEMENT_ADDNEW_MODE = False
         Else
            cmdAgmntSave.Enabled = False
            cmdAgmntEdit.Enabled = False
            cmdAgmntCancel.Enabled = False
            cmdAgmntAddNew.Enabled = True
         End If
    End If
    SelectOnly1RowFlxGrid flxPropertySelection1, flxPropertySelection1.row, 0
    szPropertySelection1 = flxPropertySelection1.TextMatrix(flxPropertySelection1.row, 1)
    AgreementClearMode ClearTextBoxes
'    AgreementClearMode ClearOnlyTextBoxes
    PayableClearMode ClearTextBoxes
    txtClientLandlord.text = ""
    txtClientLandlord.Tag = ""
    txtPayableBasis.text = ""
    txtPayableBasis.Tag = ""
    txtPercentage.text = ""
    txtPercentage.Tag = ""
    'new code by anol 20210823
    txtClientLandlord.text = ""
    txtClientLandlord.Tag = ""
    txtPayeeType.text = ""
    txtPayeeType.Tag = ""
    txtPayableBasis.text = ""
    txtPayableBasis.Tag = ""
    txtPercentage.text = ""
    txtPercentage.Tag = ""
    PayableButtonMode DefaultMode
   PayableClearMode ClearOnlyTextBoxes
   AgreementButtonMode DefaultMode
   AgreementClearMode ClearOnlyTextBoxes
   
    Call LoadflxAgreement(szPropertySelection1)
    
   
'    CPAButtonMode DefaultMode
End Sub

Private Sub flxPropertySelection1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdAgrTopEdit
    End If
End Sub

Private Sub flxPropertySelection2_Click()
    SelectOnly1RowFlxGrid flxPropertySelection2, flxPropertySelection2.row, 0
    szPropertySelection2 = flxPropertySelection2.TextMatrix(flxPropertySelection2.row, 1)
    Call ClearPaymentDates
    Call LoadPaymentDatesFromGlobalData(szPropertySelection1)
    If szPropertySelection2 = "" Then
        TabGlobalSettingSub.Enabled = False
    Else
        TabGlobalSettingSub.Enabled = True
    End If
'    AgreementClearMode ClearTextBoxes
'    PayableClearMode ClearTextBoxes
End Sub

Private Sub flxPropertySelection2_DblClick()
    FocusControl txtNoOfDaysToSendMFB4Due
End Sub

Private Sub flxPropertySelection2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtNoOfDaysToSendMFB4Due
    End If
End Sub

Private Sub Form_Activate()
'     SaveSizes
   If LOAD_CLINT_CLIENTID <> "" Then
      LoadFormByClient
   End If

   CheckBankSpare1Field
   'Find the number of controls in this form
'   Dim i As Integer
'   Dim ctrl As Control
'    For Each ctrl In frmClientNew4.Controls
'        i = i + 1
'    Next
'    MsgBox i
End Sub

Private Sub Label11_Click(Index As Integer)
On Error GoTo Err
    If Index = 0 Then
        Dim a
        Dim strFileName As String
        Dim i As Integer
        Dim j As Integer
        Dim newLine As String
        Dim strFile As String
        Dim FS

        strFileName = BrowseForFolder(Me.hWnd, "Select a Directory")
        If strFileName = "" Then Exit Sub
        strFileName = strFileName & "\ClientAccountHistory" & Format(Now, "yyyyMMddhhmmss") & ".csv"
            
        Dim iFileNo As Integer
        iFileNo = FreeFile
        'open the file for writing
        Open strFileName For Output As #iFileNo
        'please note, if this file already exists it will be overwritten!
         newLine = ""
         'Write to Specified file from Flex
            For i = 0 To flxACHistory.Rows - 1
                For j = 0 To flxACHistory.Cols - 1
                    newLine = newLine + flxACHistory.TextMatrix(i, j) + ","
                Next j
                newLine = newLine + vbCrLf
                 Print #iFileNo, newLine
                newLine = ""
            Next i

        Close #iFileNo
        MsgBox "File has been written"
      End If
    Exit Sub
Err:
    MsgBox Err.description
    
    
End Sub

Private Sub txtClientAddressLine1_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        If Index > 4 Then
                FocusControl txtClientHomeTel(0)
         Else
                FocusControl txtClientAddressLine1(Index + 1)
         End If
    End If
End Sub

'Private Sub txtClientAddressLine5_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl txtClientPostCode
'    End If
'End Sub

Private Sub txtClientHomeTel_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index >= 0 And Index < 5 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
    ElseIf Index >= 6 And Index < 10 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
    ElseIf Index = 5 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(15)
        End If
   ElseIf Index = 11 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
   ElseIf Index = 12 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
  ElseIf Index = 12 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
    
     ElseIf Index = 13 Then
        If KeyAscii = 13 Then
            FocusControl txtClientHomeTel(Index + 1)
        End If
   ElseIf Index = 14 Then
        If KeyAscii = 13 Then
            FocusControl cmdSavePaymentDetails
        End If
    End If
    
End Sub

Private Sub txtClientHomeTel_LostFocus(Index As Integer)
    If Index = 3 Or Index = 4 Then
    Dim szErrMsg As String
       If Trim(txtClientHomeTel(Index).text) <> "" Then
            If Not ValidateEmail(txtClientHomeTel(Index).text, szErrMsg) Then
               MsgBox szErrMsg, vbCritical + vbOKOnly, "Client Email"
               SelTxtInCtrl txtClientHomeTel(Index)
               txtClientHomeTel(Index).SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtClientOfficeAddressLine5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientHomeTel(20)
    End If
End Sub

Private Sub txtLastChargeDate_Change()
   'TextBoxChangeDate txtLastChargeDate
   If (Len(txtLastChargeDate) = 2 Or Len(txtLastChargeDate) = 5) And bBackSp Then Exit Sub
   TextBoxChangeDate txtLastChargeDate
End Sub

Private Sub txtLastChargeDate_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 8 And (Len(txtLastChargeDate) = 3 Or Len(txtLastChargeDate) = 6) Then
            If Len(txtLastChargeDate) = 3 Then
                    txtLastChargeDate.text = Left(txtLastChargeDate.text, 2)
            ElseIf Len(txtLastChargeDate) = 6 Then
                     txtLastChargeDate.text = Left(txtLastChargeDate.text, 5)
             End If
            bBackSp = True
    Else
            bBackSp = False
    End If
    'How to fix the bug on / when you type a date field 1. txtEND_DATE_Change event add conditional exit sub  2. txtEND_DATE_KeyDown e all codes has be written  3. declare bBackSp variable for this form

End Sub

Private Sub txtLastChargeDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtStopDatemngtFee
    End If
    Dim KA As Integer
    KA = KeyAscii
    DigitTextKeyPress txtLastChargeDate, KA
    KeyAscii = KA
End Sub

Private Sub txtLastChargeDate_LostFocus()
   TextBoxFormatDate txtLastChargeDate
End Sub




Private Sub txtAmount_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtTotalAmountPerYear
    End If
    Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtAmount, KA
   KeyAscii = KA
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.text = Format(txtAmount.text, "0.00")
    If txtChargingMethod.text = "FIXED" And txtChargeBasis.text = UCase("annual") Then
        txtTotalAmountPerYear.text = txtAmount.text
    ElseIf txtChargingMethod.text = "FIXED" And txtChargeBasis.text = UCase("percentage") Then
         If txtChargeBasis.text = UCase("Percentage") Then RCPercentage
        If txtChargeBasis.text = UCase("Annual") Then RCAnnual
    End If
     If txtFrequecymngtFee.text <> "" Then
        If txtFrequecymngtFee.text <> "N/A" Then
                txtTotalAmountPerYear.text = Format(txtTotalAmountPerYear.text, "0.00")
             If txtChargeBasis.Tag = "PC" Then RCPercentage
                If txtChargeBasis.Tag = "AN" Then RCAnnual
            End If
        End If
      If Val(txtAmount.text) > 0 And cmdAgmntSave.Enabled = True Then
            txtLastChargeDate.Locked = False
       End If
End Sub

Private Sub txtCapAmount_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtEND_DATE
    End If
    Dim KA As Integer
    KA = KeyAscii
    DigitTextKeyPress txtCapAmount, KA
    KeyAscii = KA
    
End Sub

Private Sub txtEND_DATE_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdAgmntSave
    End If
    Dim KA As Integer
    KA = KeyAscii
    'DigitTextKeyPress txtEND_DATE, KA
    TextBoxKeyPrsDate txtEND_DATE, KA
    KeyAscii = KA
End Sub

Private Sub txtNtDueDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtAmount
    End If
    Dim KA As Integer
    KA = KeyAscii
    DigitTextKeyPress txtNtDueDate, KA
    KeyAscii = KA
End Sub

Private Sub txtNtDueDate_LostFocus()
   TextBoxFormatDate txtNtDueDate
   If IsDate(txtNtDueDate.text) And IsDate(txtSTART_DATE.text) Then
        If DateDiff("d", txtSTART_DATE.text, txtNtDueDate.text) < 0 Then
            MsgBox "Next Due Date must be greater than or equal to the start date", vbInformation, "Warning"
            txtNtDueDate.text = ""
            FocusControl txtNtDueDate
        End If
    End If
End Sub


Private Sub txtNtDueDate_Change()
   TextBoxChangeDate txtNtDueDate
End Sub


Private Sub txtPAY_END_DATE_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl cmdPaySave
    End If
End Sub

Private Sub txtPAY_NtDueDate_Change()
'   TextBoxChangeDate txtPAY_NtDueDate
End Sub

Private Sub txtPaymentTerms_KeyPress(KeyAscii As MSForms.ReturnInteger)
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtPaymentTerms, KA
   KeyAscii = KA
End Sub

'Private Sub txtPAY_NtDueDate_GotFocus()
'   If txtPAY_NtDueDate.text = "" Then Exit Sub
'   SelTxtInCtrl txtPAY_NtDueDate
'End Sub



'Private Sub txtPAY_NtDueDate_KeyPress(KeyAscii As Integer)
'
'     If KeyAscii = 13 Then
'        FocusControl txtStopDate
'    End If
'   TextBoxKeyPrsDate txtPAY_NtDueDate, KeyAscii
'
'End Sub

'Private Sub txtPAY_NtDueDate_LostFocus()
'   TextBoxFormatDate txtPAY_NtDueDate
'End Sub



Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtStopDate
    End If
    DigitTextKeyPress txtPercentage, KeyAscii
End Sub

Private Sub txtPercentage_LostFocus()
    txtPercentage.text = Format(txtPercentage.text, "0.00")
End Sub

Private Sub txtPeriod_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtLastChargeDate
    End If
    Dim KA As Integer
    KA = KeyAscii
    DigitTextKeyPress txtPeriod, KA
    KeyAscii = KA
End Sub

Private Sub txtPropertySearchSel1_Change()
   'Updated by anol 2020-12-19
   Dim i As Integer
   For i = flxPropertySelection1.Rows - 1 To 1 Step -1
            flxPropertySelection1.RowHeight(i) = 240
            If InStr(1, UCase(flxPropertySelection1.TextMatrix(i, 2)), UCase(txtPropertySearchSel1.text), vbTextCompare) = 0 And txtPropertySearchSel1.text <> "" Then
                flxPropertySelection1.RowHeight(i) = 0
            End If
       
            If flxPropertySelection1.RowHeight(i) = 240 Then
                  flxPropertySelection1.row = i
            End If
   Next i
End Sub



Private Sub txtRegAdd5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientHomeTel(27).SetFocus
    End If
End Sub

Private Sub txtSearchClientID_GotFocus()
    'txtSearchClientName.Width = 500
End Sub

Private Sub txtSearchNo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            txtSearchFromD.text = ""
            txtSearchToD.text = ""
            txtSearchRef.text = ""
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            If Len(txtSearchNo.text) > 0 Then
                Call LoadFlxACHistory(adoConn, "1")
            Else
                Call LoadFlxACHistory(adoConn, "")
            End If
            If Len(txtSearchNo.text) > 0 Then
                cmdSearch.Caption = "Clear Sea&rch"
            Else
                cmdSearch.Caption = "Sea&rch"
            End If
             
            
             adoConn.Close
             Set adoConn = Nothing
             FocusControl txtSearchRef
    End If
End Sub

Private Sub txtSearchProperties2_Change()
       'Updated by anol 2020-12-19
   Dim i As Integer
   For i = flxPropertySelection2.Rows - 1 To 1 Step -1
            flxPropertySelection2.RowHeight(i) = 240
            If InStr(1, UCase(flxPropertySelection2.TextMatrix(i, 2)), UCase(txtSearchProperties2.text), vbTextCompare) = 0 And txtSearchProperties2.text <> "" Then
                flxPropertySelection2.RowHeight(i) = 0
            End If
       
            If flxPropertySelection2.RowHeight(i) = 240 Then
                  flxPropertySelection2.row = i
            End If
   Next i
End Sub

Private Sub txtSearchRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSearchFromD.text = ""
        txtSearchToD.text = ""
        txtSearchNo.text = ""
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        If Len(txtSearchRef.text) > 0 Then
            Call LoadFlxACHistory(adoConn, "2")
        Else
            Call LoadFlxACHistory(adoConn, "")
        End If
        If Len(txtSearchRef.text) > 0 Then
            cmdSearch.Caption = "Clear Sea&rch"
        Else
            cmdSearch.Caption = "Sea&rch"
        End If
            
        
        adoConn.Close
        Set adoConn = Nothing
        FocusControl txtSearchFromD
    End If
End Sub
Private Sub txtSearchFromD_Change()
    TextBoxChangeDate txtSearchFromD
    txtSearchNo.text = ""
    txtSearchRef.text = ""
End Sub
Private Sub txtSearchFromD_GotFocus()
    SelTxtInCtrl txtSearchFromD
End Sub
Private Sub txtSearchFromD_LostFocus()
    If txtSearchFromD.text <> "" Then
        TextBoxFormatDate txtSearchFromD
        txtSearchToD.text = txtSearchFromD.text
        SelTxtInCtrl txtSearchToD
     End If
End Sub





Private Sub txtSearchRef1_KeyPress(KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 13 'enter key
            MsgBox "13"
        Case Else 'who cares
            'MsgBox "Hello"
    End Select

'    If KeyAscii = 13 Then
'            Dim adoConn As New ADODB.Connection
'            adoConn.Open getConnectionString
'            Call LoadFlxACHistory(adoConn, "5")
'            adoConn.Close
'            Set adoConn = Nothing
'    End If

End Sub

Private Sub txtSearchRef1_LostFocus()
    Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            Call LoadFlxACHistory(adoConn, "5")
            adoConn.Close
            Set adoConn = Nothing
End Sub

Private Sub txtSearchToD_Change()
     TextBoxChangeDate txtSearchToD
     txtSearchNo.text = ""
     txtSearchRef.text = ""
End Sub
Private Sub txtSearchToD_GotFocus()
    SelTxtInCtrl txtSearchToD
End Sub
Private Sub txtSearchToD_LostFocus()
    If txtSearchToD.text <> "" Then TextBoxFormatDate txtSearchToD
End Sub
Private Sub txtSearchToD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearchOK.SetFocus
    End If
    TextBoxKeyPrsDate txtSearchToD, KeyAscii
End Sub

Private Sub cmdSearchOK_Click()
    fraSearch.Visible = False
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
        If Trim(txtSearchNo.text) = "" And Trim(txtSearchRef.text) = "" And Trim(txtSearchFromD.text) = "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "")
             cmdSearch.Caption = "Sea&rch"
        ElseIf Trim(txtSearchNo.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchRef.text) <> "" Then
            'do nothing
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) = "" Then
             Call LoadFlxACHistory(adoConn, "3")
             cmdSearch.Caption = "Clear Sea&rch"
        ElseIf Trim(txtSearchFromD.text) <> "" And Trim(txtSearchToD.text) <> "" Then
             cmdSearch.Caption = "Clear Sea&rch"
             Call LoadFlxACHistory(adoConn, "4")
        End If

    
    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub cmdSearchCancel_Click()
        fraSearch.Visible = False
End Sub
Private Sub txtSTART_DATE_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl cmdCommandArray(12)
    End If
    'DigitTextKeyPress txtSTART_DATE, KeyAscii
    TextBoxKeyPrsDate txtSTART_DATE, KeyAscii
End Sub

Private Sub txtStopDate_Change()
   TextBoxChangeDate txtStopDate
End Sub

Private Sub txtStopDate_GotFocus()
   If txtStopDate.text = "" Then Exit Sub
   SelTxtInCtrl txtStopDate
End Sub

Private Sub txtStopDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdPaySave
    End If
End Sub

'Private Sub txtStopDate_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        FocusControl txtPAY_END_DATE
'    End If
'   TextBoxKeyPrsDate txtStopDate, KeyAscii
'End Sub
'Private Sub txtPAY_START_DATE_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'        FocusControl cmdCommandArray(4)
'   End If
'   TextBoxKeyPrsDate txtPAY_END_DATE, KeyAscii
'End Sub
'Private Sub txtPAY_START_DATE_Change()
'   TextBoxChangeDate txtPAY_START_DATE
'End Sub

'Private Sub txtPAY_START_DATE_LostFocus()
'   TextBoxFormatDate txtPAY_END_DATE
'End Sub
Private Sub txtStopDate_LostFocus()
   TextBoxFormatDate txtStopDate
End Sub

Private Sub configflxPropertySelection1()
    flxPropertySelection1.RowHeight(0) = 0
    flxPropertySelection1.Cols = 4
    flxPropertySelection1.ColWidth(0) = 230
    flxPropertySelection1.ColWidth(1) = 950
    flxPropertySelection1.ColWidth(2) = 3200
    flxPropertySelection1.ColWidth(3) = 0 'ClientID
    flxPropertySelection1.Clear
    flxPropertySelection1.Rows = 1
    flxPropertySelection1.ColAlignment(0) = vbLeftJustify
    flxPropertySelection1.ColAlignment(1) = vbLeftJustify
    flxPropertySelection1.ColAlignment(2) = vbLeftJustify
End Sub
Private Sub configflxPropertySelection2()
    flxPropertySelection2.RowHeight(0) = 0
    flxPropertySelection2.Cols = 4
    flxPropertySelection2.ColWidth(0) = 230
    flxPropertySelection2.ColWidth(1) = 950
    flxPropertySelection2.ColWidth(2) = 3200
    flxPropertySelection2.ColWidth(3) = 0 'ClientID
    flxPropertySelection2.Clear
    flxPropertySelection2.Rows = 1
    flxPropertySelection2.ColAlignment(0) = vbLeftJustify
    flxPropertySelection2.ColAlignment(1) = vbLeftJustify
    flxPropertySelection2.ColAlignment(2) = vbLeftJustify
End Sub
'Private Sub configflxPropertySelection3()
'    flxPropertySelection3.RowHeight(0) = 0
'    flxPropertySelection3.Cols = 4
'    flxPropertySelection3.ColWidth(0) = 230
'    flxPropertySelection3.ColWidth(1) = 950
'    flxPropertySelection3.ColWidth(2) = 3200
'    flxPropertySelection3.ColWidth(3) = 0 'ClientID
'    flxPropertySelection3.Clear
'    flxPropertySelection3.Rows = 1
'    flxPropertySelection3.ColAlignment(0) = vbLeftJustify
'    flxPropertySelection3.ColAlignment(1) = vbLeftJustify
'    flxPropertySelection3.ColAlignment(2) = vbLeftJustify
'End Sub

Private Sub LoadConsolidatedBank(adoConn As ADODB.Connection)
   lblClientID(2).Visible = False
   TextBox1.Visible = False
   Dim adoRST     As New ADODB.Recordset
   Dim szSQL      As String
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer


   Dim rstRec As New ADODB.Recordset
' szSQL = "SELECT N.* " & _
'      "FROM NominalLedger AS N " & _
'      "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
'      "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
'      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
'      " ORDER BY N.Code;"
    szSQL = "SELECT BankName,BankACNumber,conBankID " & _
           "FROM ConsolidatedBankList " & _
           "ORDER BY BankName asc;"
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


    Dim rRow As Integer
   'Dim szSQL As String

  ' Dim adoConn As New ADODB.Connection
   
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 4
   flxClientList.ColWidth(0) = 80
   flxClientList.ColWidth(1) = 2000
   flxClientList.ColWidth(2) = 3000
   flxClientList.ColWidth(3) = 0
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 2000
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   Frame5.Width = 6540
   'Frame5.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID(0).Caption = "BANK NAME"
   lblClientID(1).Caption = "BANK ACCOUNT NUMBER"
   lblClientID(0).Width = 1400
   lblClientID(0).Left = 50
   lblClientID(1).Width = 2600
   lblClientID(1).Left = lblClientID(0).Left + flxClientList.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   Frame5.Height = 4095
   flxClientList.Height = 3345
  
           rRow = 1
        While Not rstRec.EOF
           flxClientList.row = 1
           flxClientList.RowSel = 1
           flxClientList.ColSel = 1
           flxClientList.TextMatrix(rRow, 0) = ""
           flxClientList.TextMatrix(rRow, 1) = rstRec.Fields("BankName").Value 'Code
           flxClientList.TextMatrix(rRow, 2) = rstRec.Fields("BankACNumber").Value 'Name
           flxClientList.TextMatrix(rRow, 3) = rstRec.Fields("conBankID").Value 'Name
           flxClientList.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClientList.AddItem ""
           rRow = rRow + 1
        Wend
 
   rstRec.Close
   Set rstRec = Nothing

End Sub

Private Sub LoadCmbNC(adoConn As ADODB.Connection)
   lblClientID(2).Visible = False
   TextBox1.Visible = False
   Dim adoRST     As New ADODB.Recordset
   Dim szSQL      As String
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer


   Dim rstRec As New ADODB.Recordset
' szSQL = "SELECT N.* " & _
'      "FROM NominalLedger AS N " & _
'      "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
'      "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
'      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
'      " ORDER BY N.Code;"
    szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger " & _
           "WHERE ClientID = '" & txtClientID.text & "' " & _
           "ORDER BY Code asc;"
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


    Dim rRow As Integer
   'Dim szSQL As String

  ' Dim adoConn As New ADODB.Connection
   
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 80
   flxClientList.ColWidth(1) = 1500
   flxClientList.ColWidth(2) = 3600
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   Frame5.Width = 6540
   'Frame5.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID(0).Caption = "Nominal Code"
   lblClientID(1).Caption = "Nominal Name"
   lblClientID(0).Width = 1400
   lblClientID(0).Left = 50
   lblClientID(1).Width = 2600
   lblClientID(1).Left = lblClientID(0).Left + flxClientList.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   Frame5.Height = 4095
   flxClientList.Height = 3345
  
           rRow = 1
        While Not rstRec.EOF
           flxClientList.row = 1
           flxClientList.RowSel = 1
           flxClientList.ColSel = 1
           flxClientList.TextMatrix(rRow, 0) = ""
           flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value 'Code
           flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value 'Name
           flxClientList.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClientList.AddItem ""
           rRow = rRow + 1
        Wend
 
   rstRec.Close
   Set rstRec = Nothing

End Sub
Private Sub LoadflxPaymentMethod(adoConn As ADODB.Connection)
   lblClientID(2).Visible = False
   TextBox1.Visible = False
   Dim adoRST     As New ADODB.Recordset
   Dim szSQL      As String
   Dim Data()     As String
   Dim TotalRow   As Integer
   Dim TotalCol   As Integer
   Dim i          As Integer
   Dim j          As Integer


   Dim rstRec As New ADODB.Recordset
' szSQL = "SELECT N.* " & _
'      "FROM NominalLedger AS N " & _
'      "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
'      "Posting AND (ISNULL(CAType) OR CAType='') AND CODE NOT IN " & _
'      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientID.text & "')" & _
'      " ORDER BY N.Code;"
'    szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code asc;"
 szSQL = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'RAT'"
      
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly


   Dim rRow As Integer
   'Dim szSQL As String

  ' Dim adoConn As New ADODB.Connection
   
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 80
   flxClientList.ColWidth(1) = 1500
   flxClientList.ColWidth(2) = 3600
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientName.Visible = True
   Frame5.Width = 6540
   'Frame5.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID(0).Caption = "Short Name"
   lblClientID(1).Caption = "Payment Method"
   lblClientID(0).Width = 1400
   lblClientID(0).Left = 50
   lblClientID(1).Width = 2600
   lblClientID(1).Left = lblClientID(0).Left + flxClientList.ColWidth(1)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   Frame5.Height = 4095
   flxClientList.Height = 3345
  
           rRow = 1
        While Not rstRec.EOF
           flxClientList.row = 1
           flxClientList.RowSel = 1
           flxClientList.ColSel = 1
           flxClientList.TextMatrix(rRow, 0) = ""
           flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value 'Code
           flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value 'Name
           flxClientList.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClientList.AddItem ""
           rRow = rRow + 1
        Wend
 
   rstRec.Close
   Set rstRec = Nothing

End Sub
Private Sub Form_Load()
    If WS_Name = "PCM-DEV2" Then
        Command1.Visible = True
        Command2.Visible = True
    End If
    
   If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
        cmdAdvanceProgr.Visible = True
   Else
        cmdAdvanceProgr.Visible = False
   End If
    sessionID = GetTimeStamp
    reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
    fraAllMemo.Top = 180
    fraAllMemo.Left = 44
    tabMain.Tab = 0
    Me.Height = 10680
    Me.Width = 18390
    chkConsolidatedStatement.Enabled = False
    chkClientAddress.Enabled = False
    chkStatementAddress.Enabled = False
    Me.BackColor = MODULEBACKCOLOR
    fraPaymentDate(15).BackColor = MODULEBACKCOLOR
    tabMain.BackColor = Me.BackColor
    Fraagreement.BackColor = Me.BackColor
    picMain.BackColor = Me.BackColor
    fraPaymentDate(16).BackColor = Me.BackColor
    fraPaymentDate(17).BackColor = Me.BackColor
    chkConsolidatedStatement.BackColor = Me.BackColor
    chkClientAddress.BackColor = Me.BackColor
    chkStatementAddress.BackColor = Me.BackColor

    
    Frame3.BackColor = Me.BackColor
    Frame4.BackColor = Me.BackColor
    Frame1(0).BackColor = Me.BackColor
    Frame1(1).BackColor = Me.BackColor
    Frame1(2).BackColor = Me.BackColor
'    optClientAdd.BackColor = Me.BackColor
'    optStAdd.BackColor = Me.BackColor
    
'    cboPaymentMethod.AddItem "CHEQUE"
'    cboPaymentMethod.AddItem "ONLINE"
'    cboPaymentMethod.AddItem "BACS"
'    cboPaymentMethod.AddItem "DIRECT DEBIT"
'    cboPaymentMethod.AddItem "Bank TRANSFER"
'    cboPaymentMethod.AddItem "TT"
'    cboPaymentMethod.AddItem "CHAPS"
'    cboPaymentMethod.AddItem "STANDING ORDER"
    
    PAYABLE_ADDNEW_MODE = False

    Call ConfigFlxPayable
    ConfigFlxOtherBankDetails

    cboBank_ID.Locked = True
'    txtREVIEW_DATE.Locked = True
    cmdAgrTopSave.Enabled = False
    txtNoOfDaysToSendMFB4Due.Locked = True
    cmdAgrTopSave.Enabled = False
    tabAgreement.Enabled = True
    cmdAgmntAddNew.Enabled = False
    cmdAgrTopEdit.Enabled = False
    
    cmdPayAddNew.Enabled = False
    cmdSetDefaultAC.Enabled = False
    cmdBACS.Enabled = False

    cmdSaveBank.Enabled = False
    chkOverDraft.Enabled = False
    chkConsolidated.Enabled = False
'    cmdGSEdit.Enabled = False
    
    
    cmdVAMemo.Enabled = False
    cmdUnitMemoNew.Enabled = False
    cmdUnitMemoEdit.Enabled = False
    cmdUnitMemoSave.Enabled = False
    cmdDelete.Enabled = False
    cmdUnitMemoCancel.Enabled = False
    cmdClientAddAtch(0).Enabled = False
    cmdClientAddAtch(1).Enabled = False
    cmdClientAddAtch(2).Enabled = False
    cmdCloseMemo.Enabled = False

    cmdImgLeftMove.Enabled = False
    cmdUploadImageAdd.Enabled = False
    cmdImgDelete.Enabled = False
    cmdDeleteClient.Enabled = False
    tabAgreement.Enabled = False
    chkOptedtoTax.Enabled = False
    
    
    
    cmdPaymentType.Enabled = False
    cmdPaymentTypeNew(0).Enabled = False
    cmdBrowseTemplate.Enabled = False
    txtPaymentTerms.Locked = True
    cmdSavePaymentDetails.Enabled = False
   
    Call FillDaysMonths
'    LoadConsolidatedBank
    'If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
            Call WheelHook(Me.hWnd)
    'End If
 '****

End Sub

Public Sub FillDaysMonths()
    Dim i As Integer, j As Integer

    For i = 31 To 42
      For j = 1 To 31
         cboDay(i).AddItem Format(j, "00")
      Next j
'      Month Name is hard coded

   Next i

'Quartarly Payment dates
   For i = 11 To 14
      For j = 1 To 31
         cboQDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboQMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
   
'Half Yearly Payment dates
   For i = 5 To 6
      For j = 1 To 31
         cboHDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboHMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
   
'Yearly Payment Dates
   For j = 1 To 31
      cboYDay(2).AddItem Format(j, "00")
   Next j
   For j = 1 To 12
      cboYMth(2).AddItem Format("1/" & j & "/2000", "MMMM")
   Next j


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   LOAD_CLINT_CLIENTID = ""
   'frmMMain.fraCmdButton.Enabled = True
'   Call WheelUnHook(Me.hWnd)

    Dim X As Integer

   If (cmdSaveBank.Enabled) Then
      X = MsgBox("Do you wish to save changes to Bank Details before closing?", vbQuestion + vbYesNo, "Bank Details")
          If X = vbYes Then
                cmdSaveBank_Click
          ElseIf X = vbCancel Then
                Cancel = 1
                'cmdCancelChange_Click
          Else
                'cancel bank details
                flxOtherBankDetails.Enabled = True
                bOverdraftWarning = False
                CommandButtonEnabled True
                EnableDisableAcText True
                NewBankText True, True
                flxOtherBankDetails_RowColChange
                cmdSetDefaultAC.Enabled = True
                'added by anol 11 Mar 2015
                cboBank_ID.Locked = True
                cmdSaveBank.Enabled = False
                cmdSetDefaultAC.Enabled = False
            '    cmdDeleteBank.Enabled = False
                cmdBACS.Enabled = False
          End If
          Exit Sub
   End If
   If cmdSaveClient.Enabled Then
             X = MsgBox("Do you wish to save changes to the Client's Details?", vbQuestion + vbYesNoCancel, "Client Details")
             If X = vbYes Then
                 cmdSaveClient_Click
                 cmdCancelChange_Click
             ElseIf X = vbCancel Then
                Cancel = 1
             Else
                cmdCancelChange_Click
             End If
             Exit Sub
   End If
   If cmdUnitMemoSave.Enabled Then
             X = MsgBox("Do you wish to save changes to the Client's Memo?", vbQuestion + vbYesNoCancel, "Client's Memo")
             If X = vbYes Then
                 cmdUnitMemoSave_Click
                 cmdUnitMemoCancel_Click
             ElseIf X = vbCancel Then
                Cancel = 1
             Else
                cmdUnitMemoCancel_Click
             End If
             Exit Sub
   End If
   
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<ID|<Name|<Type"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 200        'Solid column
   conFlxGrid.ColWidth(1) = 1200        'Client ID
   conFlxGrid.ColAlignment(2) = vbLeftJustify
   conFlxGrid.ColWidth(2) = 3900       'Client Name
   conFlxGrid.ColWidth(3) = 800       'Balance
   conFlxGrid.RowHeight(0) = 0
   conFlxGrid.Rows = 2

   'conFlxGrid.RowHeightMin = 300
End Sub

Private Sub ClientDetails(szID As String)
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String

   Conn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM Client " & _
           "WHERE ClientID='" & szID & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   txtClientHomeTel(28).text = IIf(IsNull(rst!ClientName), "", rst!ClientName)
   txtTVInfoAdd(0).text = IIf(IsNull(rst!ClientAddressLine1), "", rst!ClientAddressLine1)
   txtTVInfoAdd(1).text = IIf(IsNull(rst!ClientAddressLine2), "", rst!ClientAddressLine2)
   txtTVInfoAdd(2).text = IIf(IsNull(rst!ClientAddressLine3), "", rst!ClientAddressLine3)
   txtTVInfoPC.text = IIf(IsNull(rst!ClientPostCode), "", rst!ClientPostCode)

   rst.Close
   Set rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Sub PropertyDetails(szID As String)
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String

   Conn.Open getConnectionString

   szStr = "SELECT * " & _
         "FROM PROPERTY " & _
         "WHERE PROPERTY.PROPERTYID='" & szID & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   txtClientHomeTel(28).text = IIf(IsNull(rst!PropertyName), "", rst!PropertyName)
   txtTVInfoAdd(0).text = IIf(IsNull(rst!ProAddressLine1), "", rst!ProAddressLine1)
   txtTVInfoAdd(1).text = IIf(IsNull(rst!ProAddressLine2), "", rst!ProAddressLine2)
   txtTVInfoAdd(2).text = IIf(IsNull(rst!ProAddressLine3), "", rst!ProAddressLine3)
   txtTVInfoPC.text = IIf(IsNull(rst!PROPOSTCODE), "", rst!PROPOSTCODE)

   rst.Close
   Set rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Function TenantDetails(szID As String) As Boolean
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String, szaTemp() As String

   szaTemp = Split(szID, "$")
   szID = szaTemp(0)

   Conn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM UNITS " & _
           "WHERE UNITS.UnitNumber='" & szID & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   If rst.EOF Then
     ' MsgBox "Error in Database, Please contact PCM Consulting Ltd.", vbCritical, "Unit Not Found"
      rst.Close
      Conn.Close
      Set rst = Nothing
      Set Conn = Nothing
   Else
      If rst!UnitName <> "" Then txtClientHomeTel(28).text = rst!UnitName
      txtTVInfoAdd(0).text = IIf(rst!UnitAddressLine1 <> "", rst!UnitAddressLine1, "")
      txtTVInfoAdd(1).text = IIf(rst!UnitAddressLine2 <> "", rst!UnitAddressLine2, "")
      txtTVInfoAdd(2).text = IIf(rst!UnitAddressLine3 <> "", rst!UnitAddressLine3, "")
      txtTVInfoPC.text = IIf(rst!UnitPostCode <> "", rst!UnitPostCode, "")
      If rst!OCCUPIED = "N" Then
         TenantDetails = False
         rst.Close
         Conn.Close
         Set rst = Nothing
         Set Conn = Nothing

         Exit Function
      Else
         szaTemp = Split(LLTenant(szID, Conn), " // ")
         txtClientHomeTel(29).text = szaTemp(0)
         txtClientHomeTel(30).text = IIf(IsNull(szaTemp(1)), "", szaTemp(1))
         rst.Close
         szStr = LeaseDetails(szID, Conn)

         Conn.Close
         Set rst = Nothing
         Set Conn = Nothing

         If szStr = "NULL" Then
'            MsgBox "Please update lease information of this unit.", vbInformation + vbOKOnly, "Error"
            TenantDetails = False
         Else
            TenantDetails = True
            szaTemp = Split(szStr, " # ")

            txtClientHomeTel(31).text = szaTemp(0)
            txtClientHomeTel(32).text = szaTemp(1)
            txtClientHomeTel(33).text = szaTemp(2)
            txtClientHomeTel(34).text = szaTemp(3)
         End If
      End If
   End If
End Function

Private Function UnitDetails(szID As String) As Boolean
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String, szaTemp() As String
'        txtTVInfoAdd(0).text = ""
'        txtTVInfoAdd(1).text = ""
'        txtTVInfoAdd(2).text = ""
'        txtTVInfoPC.text = ""
   Conn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM UNITS " & _
           "WHERE UNITS.UnitNumber='" & szID & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   If rst.EOF Then
'      MsgBox "Error in Database, Please contact with PCM Consulting Ltd.", vbCritical, "Unit Not Found"
      rst.Close
      Conn.Close
      Set rst = Nothing
      Set Conn = Nothing
   Else
    'if 2nd grid visible false then load this Unit
       If fraOccupied.Visible = False Then
            If rst!UnitName <> "" Then txtClientHomeTel(28).text = rst!UnitName
            txtTVInfoAdd(0).text = IIf(rst!UnitAddressLine1 <> "", rst!UnitAddressLine1, "")
            txtTVInfoAdd(1).text = IIf(rst!UnitAddressLine2 <> "", rst!UnitAddressLine2, "")
            txtTVInfoAdd(2).text = IIf(rst!UnitAddressLine3 <> "", rst!UnitAddressLine3, "")
            txtTVInfoPC.text = IIf(rst!UnitPostCode <> "", rst!UnitPostCode, "")
      End If
      If rst!OCCUPIED = "N" Then
         UnitDetails = False
         rst.Close
         Conn.Close
         Set rst = Nothing
         Set Conn = Nothing

         Exit Function
      Else
         szaTemp = Split(LLTenant(szID, Conn), " // ")
         txtClientHomeTel(29).text = szaTemp(0)
         txtClientHomeTel(30).text = IIf(IsNull(szaTemp(1)), "", szaTemp(1))
         rst.Close
         szStr = LeaseDetails(szID, Conn)

         Conn.Close
         Set rst = Nothing
         Set Conn = Nothing

         If szStr = "NULL" Then
'            MsgBox "Please update lease information of this unit.", vbInformation + vbOKOnly, "Error"
            UnitDetails = False
         Else
            UnitDetails = True
            szaTemp = Split(szStr, " # ")

            txtClientHomeTel(31).text = szaTemp(0)
            txtClientHomeTel(32).text = szaTemp(1)
            txtClientHomeTel(33).text = szaTemp(2)
            txtClientHomeTel(34).text = szaTemp(3)
         End If
      End If
   End If
End Function

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
Private Sub LoadAllClientFlxGrd(conClient As ADODB.Connection)
'   Dim conClient As New ADODB.Connection
   Dim rstClient As New ADODB.Recordset, rstLandlord As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 360
   lblClientID(1).Left = 1635
   lblClientID(2).Left = 5175

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 4
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
   flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Client ID"
   lblClientID(1).Caption = "Client Name"
   lblClientID(2).Caption = "Balance"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
   

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"

   rstClient.Open szSQL, conClient, adOpenStatic, adLockReadOnly
'
'   szSQL = "SELECT LANDLORDID, LANDLORDNAME,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM LANDLORD " & _
'           "ORDER BY LANDLORDID;"
'
'   rstLandlord.Open szSQL, conClient, adOpenStatic, adLockReadOnly

   If rstClient.EOF Then GoTo NoRes

   Dim iRow As Integer

   iRow = 1

   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!ClientID
      flxClientList.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(iRow, 3) = Format(GetClientBalance(rstClient!ClientID), "0.00")  '
      flxClientList.RowHeight(iRow) = 280
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      iRow = iRow + 1
   Wend
    flxClientList.row = 1

NoRes:
   rstClient.Close
   Set rstClient = Nothing
   Exit Sub
ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   rstClient.Close
   Set rstClient = Nothing
End Sub

Private Sub lstCT_DblClick()
   txtCT.text = lstCT.text
   lstCT.Visible = False
End Sub

Private Sub lstCT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then lstCT_DblClick

   If KeyAscii = 27 Then lstCT_LostFocus
End Sub

Private Sub lstCT_LostFocus()
   lstCT.Visible = False
End Sub

'Private Sub lstResidency_DblClick()
'   txtResidency.text = lstResidency.text
'   lstResidency.Visible = False
'   txtVATReg.SetFocus
'End Sub

Private Sub lstResidency_KeyPress(KeyAscii As Integer)
   'If KeyAscii = 13 Then lstResidency_DblClick
   'If KeyAscii = 27 Then lstResidency_LostFocus
End Sub

'Private Sub lstResidency_LostFocus()
'   lstResidency.Visible = False
'   cmdCTSec.SetFocus
'End Sub

'Private Sub lstResidency_Scroll()
'   lstResidency.ListIndex = GetScrollPos(lstResidency.hWnd, SB_VERT)
'End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picMain.MousePointer = vbDefault
End Sub

'Private Sub LoadConsolidatedBank()
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    Dim sqlSQL As String
'    sqlSQL = "Select BankName,conBankID from consolidatedBanklist"
'    Dim rsConsolidatedBanks As New ADODB.Recordset
'    rsConsolidatedBanks.Open sqlSQL, adoConn, adOpenStatic, adLockReadOnly
'    cboconsolidatedAccountName.Clear
'    Dim TotalRow As Integer
'    TotalRow = rsConsolidatedBanks.RecordCount
'    Dim TotalCol As Integer
'    TotalCol = rsConsolidatedBanks.Fields.Count - 1
'    Dim i, j As Integer
'    Dim Data() As String
'    If rsConsolidatedBanks.EOF Then Exit Sub
'    ReDim Data(TotalCol, TotalRow) As String
'
'
'   For i = 0 To TotalRow 'end of modification
'       For j = 0 To TotalCol - 1
'           Data(j, i) = IIf(IsNull(rsConsolidatedBanks.Fields(j).Value), "", rsConsolidatedBanks.Fields(j).Value)
'       Next j
'       rsConsolidatedBanks.MoveNext
'       If rsConsolidatedBanks.EOF Then Exit For
'   Next i
'   cboconsolidatedAccountName.Column() = Data()
'   rsConsolidatedBanks.Close
'   Set rsConsolidatedBanks = Nothing
'
'
'End Sub



Private Sub tabMain_Click(PreviousTab As Integer)
  ' MousePointer = vbHourglass.
  'use PreviousTab for warning if any unsaved item are there,
  'use tab for programming with current tab ....anol 2020-12-13
    If tabMain.Tab = 2 Then
        Me.Width = 23130
    End If
    If tabMain.Tab <> 2 Then
        PayableButtonMode DefaultMode
        PayableClearMode ClearOnlyTextBoxes
        AgreementButtonMode DefaultMode
        AgreementClearMode ClearOnlyTextBoxes
    End If
    If tabMain.Tab = 0 Or tabMain.Tab = 1 Or tabMain.Tab = 3 Or tabMain.Tab = 4 Or tabMain.Tab = 5 Then
        Me.Width = 18390
    End If
    If tabMain.Tab = 3 Then
        Me.Width = 20985
    End If
   If PreviousTab = 3 Then
        If cmdSaveBank.Enabled Then
             If MsgBox("Do you wish to save changes to Bank Details?", vbQuestion + vbYesNo, "Bank Details") = vbYes Then
                tabMain.Tab = PreviousTab
                FocusControl cmdSaveBank
            Else
                Frame14.Enabled = True
                flxOtherBankDetails.Enabled = True
                bOverdraftWarning = False
                CommandButtonEnabled True
                EnableDisableAcText True
                NewBankText True, True
                flxOtherBankDetails_RowColChange
                cmdSetDefaultAC.Enabled = True
                'added by anol 11 Mar 2015
                cboBank_ID.Locked = True
                cmdSaveBank.Enabled = False
                cmdSetDefaultAC.Enabled = False
            '    cmdDeleteBank.Enabled = False
                cmdBACS.Enabled = False
            End If
        End If
   End If
    If PreviousTab = 0 Then
        If cmdSaveClient.Enabled Then
             If MsgBox("Do you wish to save changes to the Client's Details?", vbQuestion + vbYesNo, "Client Details") = vbYes Then
                tabMain.Tab = PreviousTab
                FocusControl cmdSaveClient
            Else
                cmdCancelChange_Click
            End If
        End If
   End If
   If PreviousTab = 6 Then
        If cmdUnitMemoSave.Enabled And txtClientID.text <> "" Then
             If MsgBox("Do you wish to save changes to the Client's Memo?", vbQuestion + vbYesNo, "Client's Memo") = vbYes Then
                 cmdUnitMemoSave_Click
                 cmdUnitMemoCancel_Click
             Else
                cmdUnitMemoCancel_Click
             End If
        End If
   End If
   
   Select Case tabMain.Tab

   Case 1:                    'Property
      tvwLandLord.SetFocus
      If txtClientID.text <> "" Then
              cmdUploadImageAdd.Enabled = True
              cmdImgLeftMove.Enabled = True
              cmdImgDelete.Enabled = True
      End If
   Case 2:                    'Agreement
        AgreementButtonMode DefaultMode
'        cmdAgmntAddNew.Enabled = False
        PayableButtonMode DefaultMode
        If txtClientID.text <> "" Then
              cmdUploadImageAdd.Enabled = True
              cmdImgLeftMove.Enabled = True
              cmdImgDelete.Enabled = True
             ' PrepareList4Property cboProperty ' added by anol 2020-07-02   old
              loadflxPropertySelection1 ' added by anol 2020-12-13   New
              
              Call LoadflxAgreement(szPropertySelection1)
              cmdAgrTopEdit.Enabled = True
              tabAgreement.Tab = 0
        '      If cboProperty.ListCount < 1 Then
        '         If Len(txtClientID.text) > 0 Then
        '            MsgBox "No property has been entered for this client.", vbCritical + vbOKOnly, "No Property"
        '         End If
        '      Else
                 'cboProperty.ListIndex = 0
                 'AgreementButtonMode DefaultMode
                 'PayableButtonMode DefaultMode
                 PopulateCodes
        '      End If
        End If

   Case 3:                    'Bank Account details
      If txtClientID.text <> "" Then
            'added by anol 04 May 2015 implementation of consolidated.
            chkConsolidated.Value = 0
            'End of addition
            'If LoadNCinCombo Then
            ConfigFlxOtherBankDetails
            LoadFlxOtherBankDetails
'            LoadConsolidatedBank
            Call LoadFirstValueflxOtherBankDetails
            'flxOtherBankDetails.row = 0
            flxOtherBankDetails.col = 0
           
            'End If
      End If

   Case 5:                    'Global setting
      If txtClientID.text <> "" Then
            'Call LoadGlobalData
            Call loadflxPropertySelection2
'            Call loadflxPropertySelection3
      End If
'      EnableGlobalControl False
'      tabDates.Tab = 0
'      If bGlobalData Then
'         cmdGSEdit.Caption = "Edit"
'      Else
'         cmdGSEdit.Caption = "Add New"
'      End If

      PopulateDataCombo

   Case 6:                    'Memo and File attachment
    'start edit mode on click of the memo tab
    'this code is for normal mode
        cmdVAMemo.Enabled = True
        cmdUnitMemoNew.Enabled = True
        cmdUnitMemoEdit.Enabled = False
        cmdUnitMemoSave.Enabled = False
        cmdDelete.Enabled = False
        cmdUnitMemoCancel.Enabled = False
        cmdClientAddAtch(0).Enabled = True
        cmdClientAddAtch(1).Enabled = True
        cmdClientAddAtch(2).Enabled = True
        cmdCloseMemo.Enabled = True
    
      If txtClientID.text <> "" Then
         If (NEW_TYPE = "Landlord") Then
            Call LoadAttachmentFiles(cmbFiles, txtClientID.text, "Landlord")
         Else
            Call LoadAttachmentFiles(cmbFiles, txtClientID.text, "Client")
         End If
      End If
   End Select

'   MousePointer = vbDefault
End Sub

'Public Function LoadNCinCombo() As Boolean
'   Dim adoConn As New ADODB.Connection
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String, TotalRow As Integer
'   Dim Data() As String, i As Integer
'
'   LoadNCinCombo = True
'
'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE ClientID = '" & txtClientID.text & "' " & _
'           "ORDER BY Code asc;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Or adoRst.BOF Then
'      LoadNCinCombo = False
'      adoRst.Close
'      Set adoRst = Nothing
'      adoConn.Close
'      Set adoConn = Nothing
'      Exit Function
'   End If
'
'   TotalRow = adoRst.RecordCount
'   ReDim Data(1, TotalRow - 1) As String
'
'   i = 0
'   While Not adoRst.EOF
'      Data(0, i) = adoRst.Fields.Item("Code").Value
'      Data(1, i) = adoRst.Fields.Item("Name").Value
'      i = i + 1
'      adoRst.MoveNext
'   Wend
'
'   cboNC.Clear
'   cboNC.Column() = Data()
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoConn = Nothing
'End Function

Private Sub LoadFlxACHistory_Sales()
   Dim adoConn    As New ADODB.Connection
   Dim adoRpt     As New ADODB.Recordset
   Dim adoRptDtl  As New ADODB.Recordset
   Dim szSQL      As String
   Dim iKount     As Integer
   Dim iChild     As Integer
   
   adoConn.Open getConnectionString
   
   szSQL = "SELECT Rpt.*, TT.DESCRIPTION AS TT_DES, " & _
                  "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
           "FROM tlbReceipt AS Rpt, tlbTransactionTypes AS TT " & _
           "WHERE Rpt.SageAccountNumber = '" & txtClientID.text & "' And " & _
               "Rpt.Type = TT.TYPE_ID " & _
           "ORDER BY Rpt.RDate;"
'Debug.Print szSQL
   adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = flxACHistory.Rows

   With flxACHistory
      .AddItem ""

      While Not adoRpt.EOF
'If adoRpt.Fields.Item("TransactionID").Value = 568 Then
'MsgBox ""
'End If
         If adoRpt!Type = 1 Then
            szSQL = "SELECT MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, SQ.* " & _
                    "FROM (tlbReceipt AS R INNER JOIN " & _
                    "(" & _
                        "SELECT RT.SlNumber, RT.FromTran, RT.ReceiptAmount, R.DemandRef " & _
                        "FROM (RptTransactions AS RT INNER JOIN " & _
                              "tlbReceipt AS R ON RT.ToTran = R.TransactionID) INNER JOIN " & _
                              "tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                        "Where RT.ToTran = " & adoRpt.Fields.Item("TransactionID").Value & " " & _
                    ") AS SQ ON R.TransactionID = SQ.FromTran) " & _
                        "INNER JOIN tlbTransactionTypes AS T ON R.Type = T.TYPE_ID;"
         Else
            szSQL = "SELECT RT.*, R.SlNumber AS RefID, " & _
                        "MID(TT.CONSTANT, 4, LEN(TT.CONSTANT)-3) AS PF " & _
                    "FROM (RptTransactions AS RT INNER JOIN tlbReceipt AS R ON RT.ToTran = R.TransactionID) " & _
                        "INNER JOIN tlbTransactionTypes AS TT ON R.Type = TT.TYPE_ID " & _
                    "WHERE RT.FromTran = " & adoRpt.Fields.Item("TransactionID").Value & ";"
         End If
'Debug.Print szSQL
         adoRptDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         iChild = 0
         If adoRptDtl.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(iKount, 0) = "+"
            iChild = iKount + 1
            While Not adoRptDtl.EOF
               .TextMatrix(iChild, 0) = "-"
               If adoRpt!Type = 1 Then
                  .TextMatrix(iChild, 5) = "Receipt from " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 5) = "Receipt to " & adoRptDtl.Fields.Item("PF").Value & adoRptDtl.Fields.Item("RefID").Value
               End If
               .TextMatrix(iChild, 6) = Format(adoRptDtl.Fields.Item("ReceiptAmount").Value, "0.00")
               .RowHeight(iChild) = 0
               iChild = iChild + 1
               adoRptDtl.MoveNext
               If Not adoRptDtl.EOF Then .AddItem ""
            Wend
         Else
            .TextMatrix(iKount, 0) = ""
         End If
'1:DemandRef, 2:Invoice, 3:Date, 4:Details, 5:Amount, 6:Amount (OS), 7:Amount (Dr)
         adoRptDtl.Close
'*************
         .TextMatrix(iKount, 1) = adoRpt.Fields.Item("PF").Value & adoRpt.Fields.Item("SlNumber").Value
'*************
         .TextMatrix(iKount, 2) = IIf(UCase(Left(adoRpt.Fields.Item("TT_DES").Value, 5)) = "SALES", Mid(adoRpt.Fields.Item("TT_DES").Value, 7), adoRpt.Fields.Item("TT_DES").Value)
         .TextMatrix(iKount, 3) = IIf(IsNull(adoRpt.Fields.Item("RDate").Value), "", adoRpt.Fields.Item("RDate").Value)
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item("ExtRef").Value), "", adoRpt.Fields.Item("ExtRef").Value)
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item("OSAmount").Value, "0.00")
         If adoRpt!Type = 1 Or adoRpt!Type = 23 Then
            .TextMatrix(iKount, 8) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 9) = Format(adoRpt.Fields.Item("Amount").Value, "0.00")            'Credit
         End If
         adoRpt.MoveNext
         iKount = IIf(iChild = 0, iKount + 1, iChild)
         If Not adoRpt.EOF Then .AddItem ""
      Wend

      adoRpt.Close
'############################## if there any unposted demand in the demand then its picking up here ##################
      szSQL = "SELECT D.DemandID, D.IssueDate, D.Details, SUM(DS.TotalAmount) " & _
              "FROM DemandRecords AS D, DemandSplitRecords AS DS " & _
              "WHERE D.DemandID = DS.DemandID AND DS.TrfReceipt = FALSE AND " & _
                  "D.SageAccountNumber = '" & txtClientID.text & "' " & _
              "GROUP BY D.DemandID, D.Details, D.IssueDate;"

      adoRpt.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRpt.EOF
         .AddItem ""
         .TextMatrix(iKount, 1) = adoRpt.Fields.Item(0).Value
         .TextMatrix(iKount, 2) = "Invoice"
         .TextMatrix(iKount, 3) = adoRpt.Fields.Item(1).Value
         .TextMatrix(iKount, 4) = IIf(IsNull(adoRpt.Fields.Item(2).Value), "", adoRpt.Fields.Item(2).Value)
         .TextMatrix(iKount, 5) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 6) = Format(adoRpt.Fields.Item(3).Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoRpt.Fields.Item(3).Value, "0.00")            'Debit
         adoRpt.MoveNext
         iKount = iKount + 1
         If Not adoRpt.EOF Then .AddItem ""
      Wend
      adoRpt.Close
   End With

   Set adoRpt = Nothing
   Set adoRptDtl = Nothing
   adoConn.Close
   Set adoConn = Nothing
   flxACHistory.row = 0
   flxACHistory.row = 0
End Sub

Private Sub LoadFlxACHistory_Old1()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoPty As New ADODB.Recordset, adoPtyDtl As New ADODB.Recordset

   ConfigFlxACHistory

   adoConn.Open getConnectionString

   szSQL = "SELECT P.*, TT.DESCRIPTION AS TT_DES, PI.SlNumber AS INV_REF, TT.CONSTANT " & _
           "FROM  (tlbPayment AS P INNER JOIN tlbTransactionTypes AS TT ON  " & _
                  "P.Type = TT.TYPE_ID) LEFT JOIN tblPurInv AS PI ON P.PI = PI.MY_ID " & _
           "WHERE  P.SageAccountNumber = '" & txtClientID.text & "' " & _
           "ORDER BY P.TransactionID;"
'Debug.Print szSQL
   adoPty.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iKount = 1

   With flxACHistory
      While Not adoPty.EOF
         If adoPty!Type = 6 Or adoPty!Type = 7 Then
            szSQL = "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type, P.SlNumber " & _
                    "FROM PayTransactions AS PT, tlbPayment AS P " & _
                    "WHERE PT.ToTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID;"
         Else
            szSQL = "SELECT SQ.*, P.SlNumber " & _
                    "FROM tlbPayment AS P, (" & _
                     "SELECT PT.FromTran, PT.ToTran, PT.AllocDate, PT.PaymentAmount, P.Type " & _
                     "FROM PayTransactions AS PT, tlbPayment AS P " & _
                     "WHERE PT.FromTran = " & adoPty.Fields.Item("TransactionID").Value & " AND " & _
                        "PT.FromTran = P.TransactionID) SQ " & _
                    "WHERE SQ.ToTran = P.TransactionID; "
         End If
'Debug.Print szSQL

         adoPtyDtl.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
         iChild = 0
         If adoPtyDtl.RecordCount > 0 Then
            .AddItem ""
            .TextMatrix(iKount, 0) = "+"
            iChild = iKount + 1
            While Not adoPtyDtl.EOF
               .TextMatrix(iChild, 0) = "-"
               .TextMatrix(iChild, 1) = IIf(adoPty.Fields.Item("Type").Value = 6, _
                                            adoPty.Fields.Item("INV_REF").Value, _
                                            adoPty.Fields.Item("SlNumber").Value)
               If adoPty!Type = 6 Then
                  .TextMatrix(iChild, 5) = "Payment from: PP" & adoPtyDtl.Fields.Item("SlNumber").Value
               Else
                  .TextMatrix(iChild, 5) = "Payment to: PI" & adoPtyDtl.Fields.Item("SlNumber").Value
               End If
               .TextMatrix(iChild, 6) = Format(adoPtyDtl.Fields.Item("PaymentAmount").Value, "0.00")
               .RowHeight(iChild) = 0
               iChild = iChild + 1
               adoPtyDtl.MoveNext
               If Not adoPtyDtl.EOF Then .AddItem ""
            Wend
         Else
            .TextMatrix(iKount, 0) = ""
         End If
'1:No,        2:Type,    3:Date, 4:Reference, 5:Description, 6:Amount, 7:Balance, 8:Dr, 9:Cr
         adoPtyDtl.Close
         .TextMatrix(iKount, 1) = Mid(adoPty.Fields.Item("CONSTANT").Value, 4, Len(adoPty.Fields.Item("CONSTANT").Value) - 3)
         .TextMatrix(iKount, 1) = .TextMatrix(iKount, 1) & _
                                      IIf(adoPty.Fields.Item("Type").Value = 6, _
                                      adoPty.Fields.Item("INV_REF").Value, _
                                      adoPty.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 2) = IIf(UCase(Left(adoPty.Fields.Item("TT_DES").Value, 5)) = "SALES", _
                                    Mid(adoPty.Fields.Item("TT_DES").Value, 7), _
                                    adoPty.Fields.Item("TT_DES").Value)
         If InStr(.TextMatrix(iKount, 2), "Payment") > 0 And InStr(.TextMatrix(iKount, 2), "Account") = 0 Then .TextMatrix(iKount, 2) = "Payment"
         If InStr(.TextMatrix(iKount, 2), "Account") > 0 Then .TextMatrix(iKount, 2) = "Payment on A/C"
         If InStr(.TextMatrix(iKount, 2), "Invoice") > 0 Then .TextMatrix(iKount, 2) = "Invoice"
         
         .TextMatrix(iKount, 3) = IIf(IsNull(adoPty.Fields.Item("PDate").Value), "", adoPty.Fields.Item("PDate").Value)
         .TextMatrix(iKount, 4) = IIf(IsNull(adoPty.Fields.Item("Ref").Value), "", adoPty.Fields.Item("Ref").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoPty.Fields.Item("Details").Value), "", adoPty.Fields.Item("Details").Value)
         .TextMatrix(iKount, 6) = Format(adoPty.Fields.Item("Amount").Value, "0.00")
         .TextMatrix(iKount, 7) = Format(adoPty.Fields.Item("OSAmount").Value, "0.00")
         If adoPty!Type = 6 Or adoPty!Type = 24 Then
            .TextMatrix(iKount, 9) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Debit
         Else
            .TextMatrix(iKount, 8) = Format(adoPty.Fields.Item("Amount").Value, "0.00")            'Credit
         End If
         .TextMatrix(iKount, 10) = Format(adoPty.Fields.Item("TransactionID").Value, "0.00")
         adoPty.MoveNext
         iKount = IIf(iChild = 0, iKount + 1, iChild)
         If Not adoPty.EOF Then .AddItem ""
      Wend
   End With

   adoPty.Close
   Set adoPty = Nothing
   Set adoPtyDtl = Nothing
   adoConn.Close
   Set adoConn = Nothing

   flxACHistory.row = 0
   flxACHistory.row = 0
End Sub

Private Sub PopulateDataCombo()
'   Dim sSQLQuery_ As String
'
'   adoFreq.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT CODE, VALUE " & _
'              "FROM SECONDARYCODE " & _
'              "WHERE SECONDARYCODE.PRIMARYCODE = 'FREQ' " & _
'              "ORDER BY VALUE;"
'
'   adoFreq.RecordSource = sSQLQuery_
'   adoFreq.CommandType = adCmdText
'   adoFreq.Refresh
'
'   adoFeeTypes.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT CODE, VALUE " & _
'              "FROM SECONDARYCODE " & _
'              "WHERE SECONDARYCODE.PRIMARYCODE = 'CFT' " & _
'              "ORDER BY VALUE;"
'
'   adoFeeTypes.RecordSource = sSQLQuery_
'   adoFeeTypes.CommandType = adCmdText
'   adoFeeTypes.Refresh
'   Dim adoconn As New ADODB.Connection
'   Dim rsChargeTypes As New ADODB.Recordset
'   adoconn.Open getConnectionString
'
'   sSQLQuery_ = "SELECT ID, FeeType " & _
'              "FROM ChargeTypes;"
'
'   rsChargeTypes.Open sSQLQuery_, adoconn, adOpenStatic, adLockReadOnly
'   rsChargeTypes.Close
'   Set rsChargeTypes = Nothing
   
End Sub

Public Sub SetFlxAgreementHeader(flxControl As MSHFlexGrid, ByVal rstAgreement As ADODB.Recordset)
   Dim sSQLQuery_ As String, iCol As Integer
   Dim iRecCol As Integer, iRow As Integer

   SetControlStyle flxControl
   rstAgreement.MoveFirst

   For iRow = 1 To rstAgreement.RecordCount
      For iRecCol = 0 To flxControl.Cols - 1
         For iCol = 0 To flxControl.Cols - 1
            If flxControl.TextMatrix(0, iCol) = rstAgreement.Fields.Item(iRecCol).Name Then Exit For
         Next iCol
         flxControl.TextMatrix(iRow, iCol) = IIf(IsNull(rstAgreement.Fields.Item(iRecCol).Value), "", rstAgreement.Fields.Item(iRecCol).Value)
      Next iRecCol
      rstAgreement.MoveNext
      If Not rstAgreement.EOF Then flxControl.AddItem ""
   Next iRow
End Sub

Public Sub ConfigFlxAgreement(flxControl As MSHFlexGrid)
   Dim szHeader As String, iCol As Integer

   flxControl.Clear
   flxControl.Rows = 2
   flxControl.Cols = 12

   szHeader$ = "<CHARGE_TYPE|<DEMAND_TYPE|<Fund|<Handling" & _
               "|<CHARGE_METHOD|<CHARGE_BASIS|>AnnualCharge" & _
               "|<START_DATE|<END_DATE|<Frequency|<NtDueDate|<AGREEMENT_ID"
   flxControl.FormatString = szHeader$

   For iCol = 0 To flxControl.Cols - 3
      flxControl.ColWidth(iCol) = 2000 'Label7(iCol + 1).Left - Label7(iCol).Left
   Next iCol
   flxControl.ColWidth(10) = txtNtDueDate.Width - 55
   flxControl.ColWidth(11) = 0
   flxControl.RowHeight(0) = 0
   flxControl.row = 0
   flxControl.col = 0
End Sub

Public Sub SetFlxPayableHeader(flxControl As MSHFlexGrid, ByVal rstPayable As ADODB.Recordset)
   Dim sSQLQuery_ As String, iCol As Integer
   Dim iRecCol As Integer, iRow As Integer

   SetControlStyle flxControl
   rstPayable.MoveFirst

   For iRow = 1 To rstPayable.RecordCount
      For iRecCol = 0 To flxControl.Cols - 1
         For iCol = 0 To flxControl.Cols - 1
            If flxControl.TextMatrix(0, iCol) = rstPayable.Fields.Item(iRecCol).Name Then Exit For
         Next iCol
         If iCol <> flxControl.Cols Then _
            flxControl.TextMatrix(iRow, iCol) = IIf(IsNull(rstPayable.Fields.Item(iRecCol).Value), "", rstPayable.Fields.Item(iRecCol).Value)
      Next iRecCol
      rstPayable.MoveNext
      If Not rstPayable.EOF Then flxControl.AddItem ""
   Next iRow
End Sub

Public Sub ConfigFlxPayable()
   Dim szHeader As String, iCol As Integer

   flxPayable.Clear
   flxPayable.Rows = 2
   flxPayable.Cols = 19

'   szHeader$ = "<PAYABLE_TYPE|<PAY_DEMAND_TYPE|<PAY_Fund|<PAY_HANDLING" & _
'               "|<PAYABLE_METHOD|>PAY_AnnualCharge" & _
'               "|<PAY_START_DATE|<PAY_END_DATE|<PAY_FREQUENCY|<PAY_NtDueDate|<PAYABLE_ID"
'   FlxPayable.FormatString = szHeader$

    flxPayable.ColWidth(0) = 0 'selection column
    flxPayable.ColWidth(1) = 0  'PAYABLE_ID
    flxPayable.ColWidth(2) = 0  'CPA_ID
    flxPayable.ColWidth(3) = 0 'PAYABLE_TYPE_ID
    flxPayable.ColWidth(4) = Label4(2).Left - Label4(0).Left  'PAYABLE_TYPE
    flxPayable.ColWidth(5) = 0  'DemandID
    flxPayable.ColWidth(6) = 0 'Label4(2).Left - Label4(1).Left  'PAY_DEMAND_TYPE
    flxPayable.ColWidth(7) = 0  'FundID
    flxPayable.ColWidth(8) = 2850 'Label4(4).Left - Label4(2).Left - 150 'FundName
    flxPayable.ColWidth(9) = 2100 'Label4(7).Left - Label4(3).Left - 250 'payeetype
    flxPayable.ColAlignment(9) = vbRightJustify
    flxPayable.ColWidth(10) = 2100 'Label4(7).Left - Label4(3).Left - 300 'clientLandlordID 'PAY_START_DATE
    flxPayable.ColWidth(11) = 0  'FREQID
    flxPayable.ColWidth(12) = 0 'Label4(6).Left - Label4(5).Left - 1000 'Frequency
    flxPayable.ColWidth(13) = 0 '1000 'ONDD
    flxPayable.ColWidth(14) = 0 'Label4(7).Left - Label4(6).Left 'PAYABLE_BASIS_ Short name
    flxPayable.ColWidth(15) = Label4(8).Left - Label4(7).Left  'PAYABLE_BASIS_ full name
    flxPayable.ColWidth(16) = Label4(9).Left - Label4(8).Left 'Percentage
    flxPayable.ColWidth(17) = Label4(9).Left - Label4(8).Left 'StopDate
    flxPayable.ColWidth(18) = 0 'Label4(9).Left - Label4(8).Left - 140 'PAY_END_DATE
    flxPayable.RowHeight(0) = 0
    
End Sub
Public Sub ConfigflxManagementFee()
   Dim szHeader As String, iCol As Integer

   flxManagementFee.Clear
   flxManagementFee.Rows = 2
   flxManagementFee.Cols = 26

'   szHeader$ = "<PAYABLE_TYPE|<PAY_DEMAND_TYPE|<PAY_Fund|<PAY_HANDLING" & _
'               "|<PAYABLE_METHOD|>PAY_AnnualCharge" & _
'               "|<PAY_START_DATE|<PAY_END_DATE|<PAY_FREQUENCY|<PAY_NtDueDate|<PAYABLE_ID"
'   flxManagementFee.FormatString = szHeader$

    flxManagementFee.ColWidth(0) = 0 'selection column
    flxManagementFee.ColWidth(1) = 0  'ManagementFee_ID
    flxManagementFee.ColWidth(2) = 0  'CPA_ID
    flxManagementFee.ColWidth(3) = 0 'ManagementFee_ID
    flxManagementFee.ColWidth(4) = Label4(18).Left - Label4(17).Left - 120 'ManagementFee_TYPE/Charge Type
    flxManagementFee.ColWidth(5) = 0  'DemandID
    flxManagementFee.ColWidth(6) = 0 'Label4(19).Left - Label4(18).Left  'ManagementFee_DEMAND_TYPE
    flxManagementFee.ColWidth(7) = 0  'FundID
    flxManagementFee.ColWidth(8) = Label4(20).Left - Label4(19).Left + 150 'FundName
    flxManagementFee.ColWidth(9) = Label4(21).Left - Label4(20).Left - 100 'Charge Method ID
    flxManagementFee.ColWidth(10) = 0 'Label4(22).Left - Label4(21).Left - 200 'Charge Method
    flxManagementFee.ColWidth(11) = Label4(22).Left - Label4(21).Left  'Charge Basis
    flxManagementFee.ColWidth(12) = 1275 'Label4(23).Left - Label4(22).Left - 120 'ManagingAgentID
    flxManagementFee.ColWidth(13) = 975 'Label4(24).Left - Label4(23).Left  'ManagementFee_START_DATE
    flxManagementFee.ColWidth(14) = 0  'FREQID
    flxManagementFee.ColWidth(15) = 2035  'Frequency  - 1000
    flxManagementFee.ColWidth(16) = 0 ' Blank for future use
    flxManagementFee.ColWidth(17) = 1200 'Label4(26).Left - Label4(25).Left 'ManagementFee_NtDueDate
    flxManagementFee.ColWidth(18) = 1470 'Label4(28).Left - Label4(27).Left  'amount
    flxManagementFee.ColWidth(19) = 1440 'Label4(29).Left - Label4(28).Left + 500 'Total amount
    flxManagementFee.ColWidth(20) = 1050 'Label4(30).Left - Label4(29).Left - 120 'Each Period
    flxManagementFee.ColWidth(21) = 995 'Label4(31).Left - Label4(30).Left 'last charge date
    flxManagementFee.ColWidth(22) = 1045 'Label4(31).Left - Label4(30).Left + 100 'StopDate
    flxManagementFee.ColWidth(23) = 995 'Label4(31).Left - Label4(30).Left  'Cap Amount
    flxManagementFee.ColWidth(24) = 995 'Label4(31).Left - Label4(30).Left 'End Date
     flxManagementFee.ColWidth(25) = 0 'Charge type ID
    flxManagementFee.RowHeight(0) = 0
    
End Sub
'Private Sub LoadTypes()
'   Dim sSQLQuery_ As String
'
'   adoMain.ConnectionString = getConnectionString
'
'   sSQLQuery_ = "SELECT ID, FeeType, FeeIC, FeeSagePrefix, FeeNCAmt, FeeNNAmt, " & _
'                  "FeeNCVat, FeeNNVat, FeeNCTotal, FeeNNTotal, TransactionType, " & _
'                  "CategoryCode, PaymentDates " & _
'                "FROM ChargeTypes ORDER BY ID"
'
'   adoMain.RecordSource = sSQLQuery_
'   adoMain.CommandType = adCmdText
'   adoMain.Refresh
'End Sub

Private Sub CheckBankSpare1Field()
   Dim conBank As New ADODB.Connection
   Dim szSQL As String

   conBank.Open getConnectionString

   szSQL = "Update tlbClientBanks " & _
           "Set spare1 = MY_ID " & _
           "Where ISNULL(spare1) OR spare1='';"
   conBank.Execute szSQL

   conBank.Close
   Set conBank = Nothing
End Sub

Private Function ReturnFundCode(conBank As ADODB.Connection, szBankCode As String)
    Dim szSQL As String
    Dim rsFundCodes As New ADODB.Recordset
    szSQL = "SELECT * from BankFund B, Fund F where B.FundID=F.FundID AND ClientID='" & txtClientID.text & "' " & _
           "AND BankCode='" & szBankCode & "'"
    rsFundCodes.Open szSQL, conBank, adOpenStatic, adLockReadOnly
    While Not rsFundCodes.EOF
        If ReturnFundCode = "" Then
            ReturnFundCode = rsFundCodes("FundCode").Value
        Else
            ReturnFundCode = ReturnFundCode & "," & rsFundCodes("FundCode").Value
        End If
        rsFundCodes.MoveNext
    Wend
    rsFundCodes.Close
End Function
Private Sub LoadFlxOtherBankDetails()
   Dim conBank As New ADODB.Connection
   Dim rstBank As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   conBank.Open getConnectionString

   iTotalBankAC = 0

   szSQL = "SELECT * " & _
           "FROM ((tlbClientBanks C LEFT JOIN tlbBank ON tlbBank.BANK_ID = C.BANK_ID)  LEFT JOIN NominalLedger N ON N.Code= C.nominalCode AND N.ClientID=C.CLIENT_ID) " & _
           "WHERE CLIENT_ID = '" & txtClientID.text & "' " & _
           "ORDER BY Bank_AC_Name;"
'Debug.Print szSQL
   rstBank.Open szSQL, conBank, adOpenStatic, adLockReadOnly

   If rstBank.EOF Then GoTo NoResult

   Dim iRow As Integer
   iRow = 1

   While Not rstBank.EOF
      flxOtherBankDetails.TextMatrix(iRow, 1) = IIf(IsNull(rstBank!BANK_ID), "", rstBank!BANK_ID) 'rstBank!BANK_ID
      flxOtherBankDetails.TextMatrix(iRow, 2) = IIf(IsNull(rstBank!BANK_NAME), "", rstBank!BANK_NAME) 'rstBank!BANK_NAME
      flxOtherBankDetails.TextMatrix(iRow, 3) = ReturnFundCode(conBank, IIf(IsNull(rstBank!nominalCode), "", rstBank!nominalCode))
      flxOtherBankDetails.TextMatrix(iRow, 4) = rstBank!Bank_AC_Name
      flxOtherBankDetails.TextMatrix(iRow, 5) = rstBank!BANK_AC_NUM
      flxOtherBankDetails.TextMatrix(iRow, 6) = rstBank!BANK_SC
      flxOtherBankDetails.TextMatrix(iRow, 7) = IIf(rstBank!DEFAULT_AC, "YES", "NO")
      lDefaultBankID = IIf(lDefaultBankID = 0, IIf(rstBank!DEFAULT_AC, rstBank!My_ID, 0), lDefaultBankID)
      flxOtherBankDetails.TextMatrix(iRow, 8) = IIf(IsNull(rstBank!BANK_ADDRESS1), "", rstBank!BANK_ADDRESS1)
      flxOtherBankDetails.TextMatrix(iRow, 9) = IIf(IsNull(rstBank!BANK_ADDRESS2), "", rstBank!BANK_ADDRESS2)
      flxOtherBankDetails.TextMatrix(iRow, 10) = IIf(IsNull(rstBank!BANK_ADDRESS3), "", rstBank!BANK_ADDRESS3)
      flxOtherBankDetails.TextMatrix(iRow, 11) = IIf(IsNull(rstBank!PaymentMethod), "CHEQUE", rstBank!PaymentMethod)
      flxOtherBankDetails.TextMatrix(iRow, 12) = IIf(IsNull(rstBank!BacsRef), "", rstBank!BacsRef)
      flxOtherBankDetails.TextMatrix(iRow, 13) = rstBank!My_ID
      flxOtherBankDetails.TextMatrix(iRow, 14) = IIf(IsNull(rstBank!nominalCode), "", rstBank!nominalCode)
      flxOtherBankDetails.TextMatrix(iRow, 15) = IIf(rstBank!AllowOverDraft, "YES", "NO")
      flxOtherBankDetails.TextMatrix(iRow, 16) = IIf(IsNull(rstBank!OverdraftLimit), "0.00", Format(rstBank!OverdraftLimit, "0.00"))
      'added by anol 04 May 2015
      flxOtherBankDetails.TextMatrix(iRow, 17) = IIf(rstBank!Consolidated, "YES", "NO")
      flxOtherBankDetails.TextMatrix(iRow, 18) = IIf(IsNull(rstBank!Name), "", rstBank!Name)    'IIf(rstBank!Name, "YES", "NO")
      'addding consolidated Bank features'
      flxOtherBankDetails.TextMatrix(iRow, 19) = IIf(IsNull(rstBank!ConsolidatedBankID), "", rstBank!ConsolidatedBankID) 'from tlbClientBanks
      flxOtherBankDetails.TextMatrix(iRow, 20) = IIf(IsNull(rstBank!ConsBankACNumber), "", rstBank!ConsBankACNumber)
      flxOtherBankDetails.TextMatrix(iRow, 21) = IIf(IsNull(rstBank!ConsSortCode), "", rstBank!ConsSortCode)
      flxOtherBankDetails.TextMatrix(iRow, 22) = IIf(IsNull(rstBank!conBankReadOnly), "", rstBank!conBankReadOnly)
'      flxOtherBankDetails.TextMatrix(iRow, 23) = IIf(IsNull(rstBank!FundID), "", rstBank!FundID)
'      flxOtherBankDetails.TextMatrix(iRow, 24) = IIf(IsNull(rstBank!FundName), "", rstBank!FundName)
      'End of modification
      rstBank.MoveNext
      If Not rstBank.EOF Then flxOtherBankDetails.AddItem ""
      iRow = iRow + 1
      iTotalBankAC = iTotalBankAC + 1
   Wend
   If iRow > 1 Then
        flxOtherBankDetails.row = 1
   End If
   flxOtherBankDetails.col = 0
    Dim rsClient As New ADODB.Recordset
    szSQL = "SELECT * FROM CLIENT WHERE CLIENT.ClientID = '" & txtClientID.text & "';"
   rsClient.Open szSQL, conBank, adOpenStatic, adLockReadOnly
   If rsClient.EOF Then
      conBank.Close
      Set conBank = Nothing
      Exit Sub
   End If
   If IsNull(rsClient!ShowBankAccountFunds) = True Then
        chkShowFundBankAccount.Value = 0
   ElseIf rsClient!ShowBankAccountFunds = False Then
        chkShowFundBankAccount.Value = 0
   ElseIf rsClient!ShowBankAccountFunds = True Then
         chkShowFundBankAccount.Value = 1
   End If
   rsClient.Close
   
NoResult:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub ConfigFlxOtherBankDetails()
   Dim szHeader As String, i As Integer

   flxOtherBankDetails.Clear
   flxOtherBankDetails.Cols = 25
   flxOtherBankDetails.Rows = 2
   flxOtherBankDetails.RowHeight(0) = 0

   szHeader = "<BANK_ID|<BANK_NAME|<BANK_POST_CODE|<BANK_AC_NAME|<BANK_AC_NUM|<BANK_SC|<DEFAULT_AC|MY_ID|NC|Overdraft|>Limit"
   flxOtherBankDetails.FormatString = szHeader

   flxOtherBankDetails.ColWidth(0) = 0
   ' For i = 2 To flxOtherBankDetails.Cols - 10
   For i = 2 To 9
      flxOtherBankDetails.ColWidth(i - 1) = Label6(i - 1).Left - Label6(i - 2).Left
   Next i

   flxOtherBankDetails.ColWidth(7) = Label6(7).Left - Label6(6).Left    'i=8
   flxOtherBankDetails.ColWidth(8) = 0
   flxOtherBankDetails.ColWidth(9) = 0
   flxOtherBankDetails.ColWidth(10) = 0
   flxOtherBankDetails.ColWidth(11) = 0      'PaymentMethod
   flxOtherBankDetails.ColWidth(12) = 0      'BacsRef
   flxOtherBankDetails.ColWidth(13) = 0      'MY_ID
   flxOtherBankDetails.ColWidth(14) = 0      'NC
   flxOtherBankDetails.ColWidth(15) = Label6(8).Left - Label6(7).Left    'i=8
   flxOtherBankDetails.ColWidth(16) = flxOtherBankDetails.Width + flxOtherBankDetails.Left - Label6(8).Left - 300
   flxOtherBankDetails.ColWidth(17) = 0      'Consolidated
   flxOtherBankDetails.ColWidth(18) = 0      'NL Code's Name
   flxOtherBankDetails.ColWidth(19) = 0      'Consolidated Bank ID
   flxOtherBankDetails.ColWidth(20) = 0      'Consolidated Bank AC
   flxOtherBankDetails.ColWidth(21) = 0      'Consolidated Bank Sort Code
   flxOtherBankDetails.ColWidth(22) = 0 'conBankReadOnly
    flxOtherBankDetails.ColWidth(23) = 0 'fund ID
    flxOtherBankDetails.ColWidth(24) = 0 'fund Name
End Sub

Private Sub tabMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbDefault
End Sub
Private Function FuncLevel(XNode As Node) As Integer
    On Error GoTo Err
    If XNode.Parent Is Nothing Then
        FuncLevel = 1
    End If
    If XNode.Parent.Parent Is Nothing Then
        FuncLevel = 2
    End If
    If XNode.Parent.Parent.Parent Is Nothing Then
        FuncLevel = 3
    End If
    If XNode.Parent.Parent.Parent.Parent Is Nothing Then
        FuncLevel = 4
    End If
    If XNode.Parent.Parent.Parent.Parent.Parent Is Nothing Then
        FuncLevel = 5
    End If
    Exit Function
Err:
End Function

Private Sub tvwLandLord_Click()
     If tvwLandLord.Nodes.Count = 0 Then Exit Sub
    txtClientHomeTel(31).text = ""
    txtClientHomeTel(32).text = ""
    txtClientHomeTel(33).text = ""
    txtClientHomeTel(34).text = ""
    Dim intLevel As Integer
    intLevel = FuncLevel(tvwLandLord.SelectedItem)
    'MsgBox intLevel
   'If Button = 1 Then
      szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
      fraType.Caption = szaPremisisIDType(1)

      IMAGE_FILE_NAME_ = ImageLoader(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), lblImageName)

      txtClientHomeTel(28).text = tvwLandLord.SelectedItem.text
      txtTVInfoAdd(0).text = ""
      txtTVInfoAdd(1).text = ""
      txtTVInfoAdd(2).text = ""
      txtTVInfoPC.text = ""
      
        txtClientHomeTel(31).text = ""
        txtClientHomeTel(32).text = ""
        txtClientHomeTel(33).text = ""
        txtClientHomeTel(34).text = ""
         

      If szaPremisisIDType(1) = "CLIENT" Then
         fraOccupied.Visible = False
         ClientDetails szaPremisisIDType(0)
         txtClientHomeTel(31).text = ""
         txtClientHomeTel(32).text = ""
         txtClientHomeTel(33).text = ""
         txtClientHomeTel(34).text = ""
      End If
      If szaPremisisIDType(1) = "PROPERTY" Then
         fraOccupied.Visible = False
         PropertyDetails szaPremisisIDType(0)
         txtClientHomeTel(31).text = ""
         txtClientHomeTel(32).text = ""
         txtClientHomeTel(33).text = ""
         txtClientHomeTel(34).text = ""
      End If
      If szaPremisisIDType(1) = "UNITS" Then
         If UnitDetails(szaPremisisIDType(0)) Then
            fraOccupied.Visible = True
         Else
            fraOccupied.Visible = False
         End If
      End If
      If szaPremisisIDType(1) = "TENANT" Then
         LesseeDetails szaPremisisIDType(0)
         
         fraOccupied.Visible = False
         
         If txtClientHomeTel(31).text = "" Then
           ' TenantDetails szaPremisisIDType(0)
         End If
         fraType.Visible = True
      End If
      If szaPremisisIDType(1) = "LEASE" Then
         fraOccupied.Visible = True
         fraType.Visible = True
         LeaseInfo szaPremisisIDType(0)
         
         szaPremisisIDType = Split(tvwLandLord.SelectedItem.Parent.key, "@")
         szaPremisisIDType = Split(szaPremisisIDType(0), "$")
         'MsgBox szaPremisisIDType(0)
         UnitDetails szaPremisisIDType(0)
         'LesseeDetails szaPremisisIDType(0)
         'TenantDetails szaPremisisIDType(0)
      End If
   
End Sub

Private Sub tvwLandLord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Exit Sub
   If tvwLandLord.Nodes.Count = 0 Then Exit Sub
    txtClientHomeTel(31).text = ""
    txtClientHomeTel(32).text = ""
    txtClientHomeTel(33).text = ""
    txtClientHomeTel(34).text = ""
    Dim intLevel As Integer
    intLevel = FuncLevel(tvwLandLord.SelectedItem)
    'MsgBox intLevel
   If Button = 1 Then
      szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
      fraType.Caption = szaPremisisIDType(1)

      IMAGE_FILE_NAME_ = ImageLoader(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), lblImageName)

      txtClientHomeTel(28).text = tvwLandLord.SelectedItem.text
      txtTVInfoAdd(0).text = ""
      txtTVInfoAdd(1).text = ""
      txtTVInfoAdd(2).text = ""
      txtTVInfoPC.text = ""
      
        txtClientHomeTel(31).text = ""
        txtClientHomeTel(32).text = ""
        txtClientHomeTel(33).text = ""
        txtClientHomeTel(34).text = ""
         

      If szaPremisisIDType(1) = "CLIENT" Then
         fraOccupied.Visible = False
         ClientDetails szaPremisisIDType(0)
         txtClientHomeTel(31).text = ""
         txtClientHomeTel(32).text = ""
         txtClientHomeTel(33).text = ""
         txtClientHomeTel(34).text = ""
      End If
      If szaPremisisIDType(1) = "PROPERTY" Then
         fraOccupied.Visible = False
         PropertyDetails szaPremisisIDType(0)
         txtClientHomeTel(31).text = ""
         txtClientHomeTel(32).text = ""
         txtClientHomeTel(33).text = ""
         txtClientHomeTel(34).text = ""
      End If
      If szaPremisisIDType(1) = "UNITS" Then
         If UnitDetails(szaPremisisIDType(0)) Then
            fraOccupied.Visible = True
         Else
            fraOccupied.Visible = False
         End If
      End If
      If szaPremisisIDType(1) = "TENANT" Then
         LesseeDetails szaPremisisIDType(0)
         
         fraOccupied.Visible = False
         
         If txtClientHomeTel(31).text = "" Then
           ' TenantDetails szaPremisisIDType(0)
         End If
         fraType.Visible = True
      End If
      If szaPremisisIDType(1) = "LEASE" Then
         fraOccupied.Visible = True
         fraType.Visible = True
         LeaseInfo szaPremisisIDType(0)
         
         szaPremisisIDType = Split(tvwLandLord.SelectedItem.Parent.key, "@")
         'szaPremisisIDType = Split(szaPremisisIDType(0), "$")
         LesseeDetails szaPremisisIDType(0)
         'TenantDetails szaPremisisIDType(0)
      End If
   End If
End Sub

Private Sub LeaseInfo(szID As String)
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String, szaTemp() As String

   Conn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM Tenants AS T INNER JOIN LeaseDetails AS L ON T.SageAccountNumber = L.SageAccountNumber " & _
           "WHERE L.LeaseID = '" & szID & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   txtClientHomeTel(28).text = rst!Name
   If rst!InvoiceTo = "B" Then
      txtTVInfoAdd(0).text = IIf(IsNull(rst!BillAddressLine1), "", rst!BillAddressLine1)
      txtTVInfoAdd(1).text = IIf(IsNull(rst!BillAddressLine2), "", rst!BillAddressLine2)
      txtTVInfoAdd(2).text = IIf(IsNull(rst!BillAddressLine3), "", rst!BillAddressLine3)
      txtTVInfoAdd(3).text = IIf(IsNull(rst!BillAddressLine4), "", rst!BillAddressLine4)
      txtTVInfoPC.text = IIf(IsNull(rst!BillPostCode), "", rst!BillPostCode)
   Else
      txtTVInfoAdd(0).text = IIf(IsNull(rst!HOAddressLine1), "", rst!HOAddressLine1)
      txtTVInfoAdd(1).text = IIf(IsNull(rst!HOAddressLine2), "", rst!HOAddressLine2)
      txtTVInfoAdd(2).text = IIf(IsNull(rst!HOAddressLine3), "", rst!HOAddressLine3)
      txtTVInfoAdd(3).text = IIf(IsNull(rst!HOAddressLine4), "", rst!HOAddressLine4)
      txtTVInfoPC.text = IIf(IsNull(rst!HOPostCode), "", rst!HOPostCode)
   End If
    txtClientHomeTel(31).text = IIf(IsNull(rst!StartDate), "", rst!StartDate)
    Dim strTemp As String
    If rst!OLED = True Then
           strTemp = "Override lease end date"
    Else
            strTemp = IIf(IsNull(rst!EndDate), "", rst!EndDate)
    End If
    txtClientHomeTel(32).text = strTemp
    
    txtClientHomeTel(33).text = IIf(IsNull(rst!TYPEOFSTORE), "", rst!TYPEOFSTORE)
    txtClientHomeTel(34).text = IIf(IsNull(rst!RentReviewDate), "", rst!RentReviewDate)
   rst.Close
   Set rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Sub LesseeDetails(szID As String)
   Dim Conn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim szStr As String, szaTemp() As String

   szaTemp = Split(szID, "$")

   Conn.Open getConnectionString

   szStr = "SELECT * " & _
           "FROM Tenants " & _
           "WHERE SageAccountNumber = '" & szaTemp(1) & "';"
   rst.Open szStr, Conn, adOpenStatic, adLockReadOnly

   txtClientHomeTel(28).text = rst!Name
   If rst!InvoiceTo = "B" Then
      txtTVInfoAdd(0).text = IIf(IsNull(rst!BillAddressLine1), "", rst!BillAddressLine1)
      txtTVInfoAdd(1).text = IIf(IsNull(rst!BillAddressLine2), "", rst!BillAddressLine2)
      txtTVInfoAdd(2).text = IIf(IsNull(rst!BillAddressLine3), "", rst!BillAddressLine3)
      txtTVInfoAdd(3).text = IIf(IsNull(rst!BillAddressLine4), "", rst!BillAddressLine4)
      txtTVInfoPC.text = IIf(IsNull(rst!BillPostCode), "", rst!BillPostCode)
   Else
      txtTVInfoAdd(0).text = IIf(IsNull(rst!HOAddressLine1), "", rst!HOAddressLine1)
      txtTVInfoAdd(1).text = IIf(IsNull(rst!HOAddressLine2), "", rst!HOAddressLine2)
      txtTVInfoAdd(2).text = IIf(IsNull(rst!HOAddressLine3), "", rst!HOAddressLine3)
      txtTVInfoAdd(3).text = IIf(IsNull(rst!HOAddressLine4), "", rst!HOAddressLine4)
      txtTVInfoPC.text = IIf(IsNull(rst!HOPostCode), "", rst!HOPostCode)
   End If

   rst.Close
   Set rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub
Private Sub LockAcText(bLock As Boolean)
   cboBank_ID.Locked = bLock
   txtPaymentMethod.Locked = bLock
   cmdNC.Enabled = Not bLock
   cmdPaymentTypeNew(1).Enabled = Not bLock
   cmdPaymentTypeNew(2).Enabled = Not bLock
   txtBank_AC_Name.Locked = bLock
'   txtBANK_SC.Locked = bLock
   txtBANK_AC_NUM.Locked = bLock
   txtOverDraft.Locked = bLock
   txtBankBalance.Locked = bLock
    txtRetention.Locked = bLock
    txtAvailableBankBalance.Locked = bLock
   'chkOverDraft.Locked = bLock
   'chkConsolidated.Locked = bLock
End Sub
Private Sub EnableDisableAcText(bLock As Boolean)
   'cboBank_ID.Locked = bLock 'added by anol 2020-06-24
   cboBank_ID.Enabled = Not bLock
'   cboPaymentMethod.Locked = bLock
'   cboNC.Locked = bLock
'   txtBank_AC_Name.Locked = bLock
'   txtBANK_SC.Locked = bLock
'   txtBANK_AC_NUM.Locked = bLock
'   txtOverDraft.Locked = bLock
'   chkOverDraft.Enabled = Not bLock
    flxBankAccountFund.Enabled = Not bLock
    chkShowFundBankAccount.Enabled = Not bLock

   txtPaymentMethod.Enabled = Not bLock
   txtPaymentMethod.Locked = bLock
    cmdNC.Enabled = Not bLock
    cmdPaymentTypeNew(1).Enabled = Not bLock
    cmdPaymentTypeNew(2).Enabled = Not bLock
     
   
   txtBank_AC_Name.Enabled = Not bLock
   txtNCCODE.Enabled = Not bLock
   txtNominal.Enabled = Not bLock
   
   txtBank_AC_Name.Locked = bLock
   txtBANK_SC.Enabled = Not bLock
   cmdconsolidatedAccountName.Enabled = Not bLock
   txtBANK_SC.Locked = bLock
  ' cboconsolidatedAccountName.Locked = bLock
   txtBANK_AC_NUM.Enabled = Not bLock
   txtBankBalance.Enabled = Not bLock
    txtRetention.Enabled = Not bLock
    txtAvailableBankBalance.Enabled = Not bLock
   txtconsolidatedAccountName.Enabled = Not bLock
   txtBANK_AC_NUM.Locked = bLock
   txtOverDraft.Enabled = Not bLock
   txtOverDraft.Locked = bLock
   chkOverDraft.Enabled = Not bLock
'   txtFundName.Enabled = Not bLock
'   txtFundCode.Enabled = Not bLock
   'chkOverDraft.Locked = bLock
   
'added by anol 04 Apr 2015,Implementation of consolidated
   chkConsolidated.Enabled = Not bLock
'End if
 '  If Not bBankNewEdit Then Exit Sub

'   txtBank_AC_Name.text = ""
'   txtBANK_SC.text = ""
'   txtBANK_AC_NUM.text = ""
'   txtBacsRef.text = ""
End Sub

Private Sub LockingAllTextClientAddress(bLock As Boolean)
   txtClientName.Locked = bLock
'   lstResidency.Enabled = Not bLock
   txtVATReg.Locked = bLock
   txtYearEndDate.Locked = bLock
   txtClientAddressLine1(0).Locked = bLock
   chkConsolidatedStatement.Enabled = Not bLock
   chkClientAddress.Enabled = Not bLock
   chkStatementAddress.Enabled = Not bLock
   txtClientAddressLine1(1).Locked = bLock
   txtClientAddressLine1(2).Locked = bLock
   txtClientAddressLine1(3).Locked = bLock
   txtClientAddressLine1(4).Locked = bLock
   txtClientAddressLine1(5).Locked = bLock
   txtClientHomeTel(0).Locked = bLock
   txtClientHomeTel(1).Locked = bLock
   txtClientHomeTel(2).Locked = bLock
   txtClientHomeTel(3).Locked = bLock
   txtClientHomeTel(4).Locked = bLock
   txtClientHomeTel(5).Locked = bLock
   txtClientHomeTel(6).Locked = bLock
   txtClientHomeTel(7).Locked = bLock
   txtClientHomeTel(8).Locked = bLock
   txtClientHomeTel(9).Locked = bLock
   txtClientHomeTel(10).Locked = bLock
   txtClientHomeTel(15).Locked = bLock
   txtClientHomeTel(16).Locked = bLock
   txtClientHomeTel(17).Locked = bLock
   txtClientHomeTel(18).Locked = bLock
   txtClientHomeTel(19).Locked = bLock
   txtClientHomeTel(20).Locked = bLock
   txtClientHomeTel(22).Locked = bLock
   txtClientHomeTel(23).Locked = bLock
   txtClientHomeTel(24).Locked = bLock
   txtClientHomeTel(25).Locked = bLock
   txtClientHomeTel(26).Locked = bLock
   txtClientHomeTel(21).Locked = bLock
   txtClientHomeTel(27).Locked = bLock
End Sub

'Private Sub txtAnnualCharge_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   'Added By Samrat. 12/10/2006
'   Dim KA As Integer
'   KA = KeyAscii
'   DigitTextKeyPress txtAnnualCharge, KA
'   KeyAscii = KA
'End Sub

Private Sub txtBANK_AC_NUM_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtBANK_ADDRESS1_DblClick()
   'MsgBox "To edit the bank details, please go to Bank through Tools menu!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_ADDRESS2_DblClick()
   'MsgBox "To edit the bank details, please go to Bank through Tools menu!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_ADDRESS3_DblClick()
   'MsgBox "To edit the bank details, please go to Bank through Tools menu!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_NAME_DblClick()
   'MsgBox "To edit the bank details, please go to Bank through Tools menu!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_POST_CODE_DblClick()
   'MsgBox "To edit the bank details, please go to Bank through Tools menu!.", vbInformation + vbOKOnly, "Bank Details"
End Sub

Private Sub txtBANK_SC_KeyPress(KeyAscii As Integer)
   'If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 8 Then KeyAscii = 0
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtBANK_SC, KA
   KeyAscii = KA
End Sub





'Private Sub txtClientAddressLine2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl txtClientAddressLine3
'    End If
'End Sub

'Private Sub txtClientAddressLine4_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        FocusControl txtClientAddressLine5
'    End If
'End Sub



Private Sub txtClientID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then
        Exit Sub
   End If
   If KeyAscii = 13 Then
        FocusControl txtVATReg
   End If
   If (KeyAscii >= 65 And KeyAscii <= 90) Or _
         (KeyAscii >= 97 And KeyAscii <= 122) Or _
         (KeyAscii >= 48 And KeyAscii <= 57) Then
      If (KeyAscii >= 97 And KeyAscii <= 122) Then
         KeyAscii = KeyAscii - 32
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txtClientID_LostFocus()
'   If txtClientID.Locked Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'   Dim szSQL   As String
'   Dim szID    As String
'
'   adoConn.Open getConnectionString
'
'   szID = txtClientID.text
'
'   If (IsAccountExist(szID, adoConn)) Then
'      If (Not (txtClientID.text = szID)) Then
'         MsgBox "The ID is already in use. Possible suggestion is '" & szID & "' and you may chose different ID"
'         txtClientID.text = szID
'         SelTxtInCtrl txtClientID
'      End If
'   End If
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

Private Sub txtClientName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtClientID
    End If
    If KeyAscii = 44 Then KeyAscii = 0
End Sub

Private Sub txtClientName_LostFocus()
   Dim szChoice As String, szaChoice() As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString
   szSQL = "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRST.EOF Then
      szChoice = adoRST.Fields.Item("Value").Value
      szaChoice = Split(szChoice, "#")
   End If

   adoRST.Close
   Set adoRST = Nothing
   adoConn.Close
   Set adoConn = Nothing

'   If UBound(szaChoice) > 0 Then
'      If szaChoice(1) <> "" Then
'         If InStr(szaChoice(1), "CL") > 0 Then
'            If ADD_NEW_CLIENT And txtClientID.text = "" And txtClientName.text <> "" Then txtClientID.text = CreateClientId(txtClientName.text)
'         End If
'      End If
'   End If
End Sub

Private Sub txtClientOfficeAddressLine2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientHomeTel(17).SetFocus
    End If
End Sub

Private Sub txtClientOfficeAddressLine4_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl txtClientHomeTel(19)
    End If
End Sub



'Private Sub txtClientHomeTel_LostFocus()
'   Dim szErrMsg As String
'
'   If Trim(txtClientHomeTel(4).text) <> "" Then
'      If Not ValidateEmail(txtClientHomeTel(4).text, szErrMsg) Then
'         MsgBox szErrMsg, vbCritical + vbOKOnly, "Client Email"
'         SelTxtInCtrl txtClientHomeTel(4)
'         txtClientHomeTel(4).SetFocus
'      End If
'   End If
'End Sub



'Private Sub txtClientHomeTel_LostFocus()
'   Dim szErrMsg As String
'
'   If Trim(txtClientHomeTel(3).text) <> "" Then
'      If Not ValidateEmail(txtClientHomeTel(3).text, szErrMsg) Then
'         MsgBox szErrMsg, vbCritical + vbOKOnly, "Client Email"
'         SelTxtInCtrl txtClientHomeTel(3)
'         txtClientHomeTel(3).SetFocus
'      End If
'   End If
'   If Trim(txtClientHomeTel(4).text) <> "" Then
'      If Not ValidateEmail(txtClientHomeTel(4).text, szErrMsg) Then
'         MsgBox szErrMsg, vbCritical + vbOKOnly, "Client Email"
'         SelTxtInCtrl txtClientHomeTel(4)
'         txtClientHomeTel(4).SetFocus
'      End If
'   End If
'End Sub

Private Sub txtCompReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientHomeTel(22).SetFocus
    End If
End Sub

Private Sub txtEND_DATE_Change()
   If (Len(txtEND_DATE) = 2 Or Len(txtEND_DATE) = 5) And bBackSp Then Exit Sub
   TextBoxChangeDate txtEND_DATE
End Sub

Private Sub txtEND_DATE_GotFocus()
   If txtSTART_DATE.text = "" Then
      MsgBox "Please enter start date before the end date.", vbCritical + vbOKOnly, "End Date"
      FocusControl txtSTART_DATE
   End If
   txtEND_DATE.SelStart = 0
   txtEND_DATE.SelLength = Len(txtStopDatemngtFee)
End Sub

Private Sub txtEND_DATE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 8 And (Len(txtEND_DATE) = 3 Or Len(txtEND_DATE) = 6) Then
            If Len(txtEND_DATE) = 3 Then
                    txtEND_DATE.text = Left(txtEND_DATE.text, 2)
            ElseIf Len(txtEND_DATE) = 6 Then
                     txtEND_DATE.text = Left(txtEND_DATE.text, 5)
             End If
            bBackSp = True
    Else
            bBackSp = False
    End If
    'How to fix the bug on / when you type a date field 1. txtEND_DATE_Change event add conditional exit sub  2. txtEND_DATE_KeyDown e all codes has be written  3. declare bBackSp variable for this form
End Sub

Private Sub txtEND_DATE_LostFocus()
   TextBoxFormatDate txtEND_DATE
End Sub

Private Sub txtFeeIsuDays_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

'Private Sub txtClientHomeTel(5)_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtClientOfficeAddressLine1.SetFocus
'    End If
'End Sub

Private Sub txtNoOfDaysToSendMFB4Due_GotFocus()
   If txtNoOfDaysToSendMFB4Due.text = "" Then Exit Sub

   SelTxtInCtrl txtNoOfDaysToSendMFB4Due
End Sub

Private Sub txtNoOfDaysToSendMFB4Due_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
        FocusControl cboDay(2)
    End If
   DigitTextKeyPress txtNoOfDaysToSendMFB4Due, KeyAscii
End Sub

Private Sub txtOverDraft_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPAY_AnnualCharge_GotFocus()
'   If cboPAYABLE_METHOD.text = "RECEIVABLE" Or cboPAYABLE_METHOD.text = "RECEIVED" Then
'      txtPAY_AnnualCharge.text = ""
'      txtPAY_AnnualCharge.Locked = True
'   Else
'      txtPAY_AnnualCharge.Locked = False
'   End If
End Sub

'Private Sub txtPAY_AnnualCharge_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   'Added By Samrat. 12/10/2006
'   Dim KA As Integer
'
'   KA = KeyAscii
'   DigitTextKeyPress txtPAY_AnnualCharge, KA
'   KeyAscii = KA
'End Sub

'Private Sub txtPAY_END_DATE_Change()
'   TextBoxChangeDate txtPAY_END_DATE
'End Sub

'Private Sub txtPAY_END_DATE_LostFocus()
'   TextBoxFormatDate txtPAY_END_DATE
'End Sub





Private Sub txtPayIsuDays_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtRegAdd2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientHomeTel(24).SetFocus
    End If
End Sub

Private Sub txtRegAdd4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientHomeTel(26).SetFocus
    End If
End Sub

Private Sub txtRegPostCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSaveClient
    End If
End Sub

Private Sub txtREVIEW_DATE_Change()
   TextBoxChangeDate txtREVIEW_DATE
End Sub

Private Sub txtREVIEW_DATE_GotFocus()
   If txtREVIEW_DATE.text = "" Then Exit Sub

   SelTxtInCtrl txtREVIEW_DATE
End Sub

Private Sub txtREVIEW_DATE_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        FocusControl txtAgreementStartDate
'    End If
    If KeyAscii = 13 Then
        FocusControl cmdAgrTopSave
    End If
   TextBoxKeyPrsDate txtREVIEW_DATE, KeyAscii
End Sub

Private Sub txtREVIEW_DATE_LostFocus()
   TextBoxFormatDate txtREVIEW_DATE
End Sub

Private Sub txtAgreementStartDate_Change()
   TextBoxChangeDate txtAgreementStartDate
End Sub

Private Sub txtAgreementStartDate_GotFocus()
   If txtAgreementStartDate.text = "" Then Exit Sub
   SelTxtInCtrl txtAgreementStartDate
End Sub

Private Sub txtAgreementStartDate_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        FocusControl txtAgreementEndDate
    End If
   TextBoxKeyPrsDate txtAgreementStartDate, KeyAscii
End Sub

Private Sub txtAgreementStartDate_LostFocus()
   TextBoxFormatDate txtAgreementStartDate
End Sub

Private Sub txtAgreementEndDate_Change()
   TextBoxChangeDate txtAgreementEndDate
End Sub

Private Sub txtAgreementEndDate_GotFocus()
   If txtAgreementEndDate.text = "" Then Exit Sub
   SelTxtInCtrl txtAgreementEndDate
End Sub

Private Sub txtAgreementEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtREVIEW_DATE 'cmdAgrTopSave
    End If
   TextBoxKeyPrsDate txtAgreementEndDate, KeyAscii
End Sub

Private Sub txtAgreementEndDate_LostFocus()
   TextBoxFormatDate txtAgreementEndDate
End Sub
Private Sub txtRRPA_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSearchClientID_Change()
       
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240

      If InStr(1, UCase(flxClientList.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
'    Call cmdSupplierFilter_Click
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientName_Change()
       'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
        flxClientList.RowHeight(i) = 240

        If InStr(1, UCase(flxClientList.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
        End If

      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
        'cmdSupplierFilter_Click
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        flxClientList.SetFocus
    End If
End Sub

Private Sub txtSTART_DATE_Change()
   TextBoxChangeDate txtSTART_DATE
End Sub
'
'Private Sub txtSTART_DATE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'   TextBoxKeyPrsDate txtSTART_DATE, KeyCode
'End Sub

Private Sub txtSTART_DATE_LostFocus()
   TextBoxFormatDate txtSTART_DATE
   
   
   
'    TextBoxFormatDate txtSTART_DATE
'   If txtSTART_DATE.Enabled Then Exit Sub

'   If txtSTART_DATE.text <> "" And txtChargeType.Tag <> "" Then  'And txtNtDueDate.text = "" removed by anol 04/04/2016
'      If TextBoxFormatDate(txtSTART_DATE) And txtNtDueDate.text <> "" Then 'I have added this part And txtNtDueDate.text <> ""  20180405
'         If txtNtDueDate.text <> "N/A" Then
'                NextDueDate txtFrequecymngtFee.Tag, txtSTART_DATE, txtNtDueDate, szPropertySelection1
'         End If
'      End If
'   End If
       If txtSTART_DATE.text <> "N/A" And cmdAgmntSave.Enabled = True And txtFrequecymngtFee.Tag <> "" Then
                        If txtSTART_DATE.text <> "" And Trim(txtNtDueDate.text) = "" Then
                            Call NextDueDate(CInt(txtFrequecymngtFee.Tag), txtSTART_DATE, txtNtDueDate, szPropertySelection1)
                            'Now set FDD for this charge type
                            dtFDD = NextDueDate1(CInt(txtFrequecymngtFee.Tag), txtNtDueDate, szPropertySelection1)
                        Else
                            txtComparenextDueDate1 = txtNtDueDate.text
                            Call NextDueDate(CInt(txtFrequecymngtFee.Tag), txtSTART_DATE, txtComparenextDueDate1, szPropertySelection1)
                            If txtComparenextDueDate1 <> txtNtDueDate.text Then
                                    If MsgBox("Do you wish to update the Next Due Date with the calculated Next Due Date of '" & txtComparenextDueDate1 & "' ?", vbYesNo, "Please confirm?") = vbYes Then
                                              txtNtDueDate = txtComparenextDueDate1
                                     Else
                                        FocusControl txtNtDueDate
                                        txtNtDueDate.SelStart = 0
                                        txtNtDueDate.SelLength = Len(txtNtDueDate)
                                    End If
                            End If
                        End If
                End If

'   If txtDemandTypemngtFee.text = "" Then
'      MsgBox "Please select a demand type.", vbCritical + vbOKOnly, "Rent Charges"
'      FocusControl cmdCommandArray(1)
'   End If
   
   
   
End Sub
'Private Function Fill_Form(frmCurrent As Form, ByVal adoConnector As Adodc) As Boolean
'   Dim iFieldsCount As Integer, iControlCount As Integer, i As Integer, j As Integer
'   Dim sNextField As String, cControl As Control
'
'   On Error Resume Next
'
'   If adoConnector.Recordset.RecordCount < 1 Then
'      Fill_Form = False
'      Exit Function
'   End If
'   iFieldsCount = adoConnector.Recordset.Fields.count
'   iControlCount = frmCurrent.Controls.count
'
'   For i = 0 To iFieldsCount - 1
'       sNextField = adoConnector.Recordset.Fields(i).Name
'       For Each cControl In frmCurrent.Controls
'           If UCase(sNextField) = UCase(Mid(CStr(cControl.Name), 4)) Then
'               Select Case TypeName(cControl)
'                   Case "TextBox"
'                        If cControl.Name <> "txtYearEndDate" Then
'                            cControl.text = IIf(IsNull(adoConnector.Recordset.Fields(i).Value), "", adoConnector.Recordset.Fields(i).Value)
'                            Debug.Print cControl.Name & "=" & adoConnector.Recordset.Fields(i).Name
'                        End If
'                        Exit For
'
'                   Case "CheckBox"
'                        cControl.Value = adoConnector.Recordset.Fields(i).Value
'                         Debug.Print cControl.Name & "=" & adoConnector.Recordset.Fields(i).Name
'                        Exit For
'
'                   Case "ComboBox"
'                           cControl.text = adoConnector.Recordset.Fields(i).Value
'                            Debug.Print cControl.Name & "=" & adoConnector.Recordset.Fields(i).Name
'                        Exit For
'
'               End Select
'           End If
'       Next cControl
'   Next i
'
'   adoConnector.Recordset.Update
'   adoConnector.Refresh
'
'   Fill_Form = True
'End Function
Private Function Fill_Form() As Boolean
   Dim adoConn As New ADODB.Connection
   Dim rsClient As New ADODB.Recordset
   Dim szSQL As String
   adoConn.Open getConnectionString
'   szSQL = "SELECT ClientID, ClientName, ClientAddressLine1, ClientAddressLine2, ClientAddressLine3,ClientAddressLine5,ClientOfficeAddressLine5,RegAdd5, ClientPostCode,ClientOfficeEmail, ClientPersonalEmail,ClientHomeTel,ClientMobile, '', " & _
'                "ClientOfficeAddressLine1, ClientOfficeAddressLine2, ClientOfficeAddressLine3, ClientOfficePostCode, ClientOfficeTel, ClientMemo, LandLordSageCustAC, LandLordSageSuppAC, BANK_ID," & _
'                "CommissionType, CommissionAmt, BGRPayable, VATReg, AcBalance, Residency,  PaymentMethod, BacsRef, HomeOfficeAdd, CompReg, RegAdd1, RegAdd2, RegAdd3, RegPostCode, CT, " & _
'                "ClientAddressLine4, ClientOfficeAddressLine4, RegAdd4,groupCode,Comments1,Comments2,RentSummaryTemplate,RemittanceTemplate,ConsolidatedStatement FROM CLIENT WHERE CLIENT.ClientID = '" & txtClientID.text & "';"
'rem by anol 20210830
 szSQL = "SELECT * FROM CLIENT WHERE CLIENT.ClientID = '" & txtClientID.text & "';"


   rsClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   adoConn.Close
'   Set adoConn = Nothing
   If rsClient.EOF Then
      Fill_Form = False
      adoConn.Close
      Set adoConn = Nothing
      Exit Function
   End If
   If IsNull(rsClient!ShowBankAccountFunds) = True Then
        chkShowFundBankAccount.Value = 0
   ElseIf rsClient!ShowBankAccountFunds = False Then
        chkShowFundBankAccount.Value = 0
   ElseIf rsClient!ShowBankAccountFunds = True Then
         chkShowFundBankAccount.Value = 1
   End If
    'chkShowFundBankAccount.Value = (IIf(IsNull(rsClient!ShowBankAccountFunds), False, rsClient!ShowBankAccountFunds))
    If chkShowFundBankAccount.Value Then
        flxBankAccountFund.Visible = True
    Else
        flxBankAccountFund.Visible = False
    End If
    txtClientID.text = IIf(IsNull(rsClient!ClientID), "", rsClient!ClientID)
    txtClientName.text = IIf(IsNull(rsClient!ClientName), "", rsClient!ClientName) 'rsClient!ClientName
    txtClientAddressLine1(0).text = IIf(IsNull(rsClient!ClientAddressLine1), "", rsClient!ClientAddressLine1) 'rsClient!ClientAddressLine1
    txtClientAddressLine1(1).text = IIf(IsNull(rsClient!ClientAddressLine2), "", rsClient!ClientAddressLine2) 'rsClient!ClientAddressLine2
    txtClientAddressLine1(2).text = IIf(IsNull(rsClient!ClientAddressLine3), "", rsClient!ClientAddressLine3) 'rsClient!ClientAddressLine3
    txtClientAddressLine1(5).text = IIf(IsNull(rsClient!ClientPostCode), "", rsClient!ClientPostCode) 'rsClient!ClientPostCode
    
    
    txtClientHomeTel(0).text = IIf(IsNull(rsClient!ClientHomeTel), "", rsClient!ClientHomeTel) 'rsClient!ClientHomeTel
    txtClientHomeTel(1).text = IIf(IsNull(rsClient!ClientOfficeTel), "", rsClient!ClientOfficeTel) 'rsClient!ClientOfficeTel
    txtClientHomeTel(2).text = IIf(IsNull(rsClient!ClientMobile), "", rsClient!ClientMobile) 'rsClient!ClientMobile
    txtClientHomeTel(3).text = IIf(IsNull(rsClient!ClientPersonalEmail), "", rsClient!ClientPersonalEmail) 'rsClient!ClientPersonalEmail
    txtClientHomeTel(4).text = IIf(IsNull(rsClient!ClientOfficeEmail), "", rsClient!ClientOfficeEmail) ' rsClient!ClientOfficeEmail
    txtClientHomeTel(5).text = IIf(IsNull(rsClient!groupCode), "", rsClient!groupCode) 'rsClient!groupCode
    
    txtClientHomeTel(6).text = IIf(IsNull(rsClient!StClientHomeTel), "", rsClient!StClientHomeTel) 'rsClient!ClientHomeTel
    txtClientHomeTel(7).text = IIf(IsNull(rsClient!StClientOfficeTel), "", rsClient!StClientOfficeTel) 'rsClient!ClientOfficeTel
    txtClientHomeTel(8).text = IIf(IsNull(rsClient!StClientMobile), "", rsClient!StClientMobile) 'rsClient!ClientMobile
    txtClientHomeTel(9).text = IIf(IsNull(rsClient!StClientPersonalEmail), "", rsClient!StClientPersonalEmail) 'rsClient!ClientPersonalEmail
    txtClientHomeTel(10).text = IIf(IsNull(rsClient!StClientOfficeEmail), "", rsClient!StClientOfficeEmail) ' rsClient!ClientOfficeEmail
    
    
    
    txtClientHomeTel(15).text = IIf(IsNull(rsClient!ClientOfficeAddressLine1), "", rsClient!ClientOfficeAddressLine1) 'rsClient!ClientOfficeAddressLine1
    txtClientHomeTel(16).text = IIf(IsNull(rsClient!ClientOfficeAddressLine2), "", rsClient!ClientOfficeAddressLine2) 'rsClient!ClientOfficeAddressLine2
    txtClientHomeTel(17).text = IIf(IsNull(rsClient!ClientOfficeAddressLine3), "", rsClient!ClientOfficeAddressLine3) 'rsClient!ClientOfficeAddressLine3
    txtClientHomeTel(20).text = IIf(IsNull(rsClient!ClientOfficePostCode), "", rsClient!ClientOfficePostCode) 'rsClient!ClientOfficePostCode
    
    cboBank_ID.text = IIf(IsNull(rsClient!BANK_ID), "", rsClient!BANK_ID) 'rsClient!BANK_ID
    txtVATReg.text = IIf(IsNull(rsClient!VATReg), "", rsClient!VATReg) 'rsClient!VATReg
    'txtAcBalance(0).text = IIf(IsNull(rsClient!AcBalance), "0.00", rsClient!AcBalance) ' rsClient!AcBalance
    txtPaymentMethod.text = IIf(IsNull(rsClient!PaymentMethod), "", rsClient!PaymentMethod) 'rsClient!PaymentMethod
    txtPaymentMethod.Tag = IIf(IsNull(rsClient!PaymentMethod), "", rsClient!PaymentMethod)
    txtClientHomeTel(21).text = IIf(IsNull(rsClient!CompReg), "", rsClient!CompReg) 'rsClient!CompReg
    txtClientHomeTel(22).text = IIf(IsNull(rsClient!RegAdd1), "", rsClient!RegAdd1) 'rsClient!RegAdd1
    txtClientHomeTel(23).text = IIf(IsNull(rsClient!RegAdd2), "", rsClient!RegAdd2) 'rsClient!RegAdd2
    txtClientHomeTel(24).text = IIf(IsNull(rsClient!RegAdd3), "", rsClient!RegAdd3) 'rsClient!RegAdd3
    txtClientHomeTel(27).text = IIf(IsNull(rsClient!RegPostCode), "", rsClient!RegPostCode) 'rsClient!RegPostCode
    txtCT.text = IIf(IsNull(rsClient!CT), "", rsClient!CT) 'rsClient!CT
    txtClientAddressLine1(3).text = IIf(IsNull(rsClient!ClientAddressLine4), "", rsClient!ClientAddressLine4) ' rsClient!ClientAddressLine4
    txtClientAddressLine1(4).text = IIf(IsNull(rsClient!ClientAddressLine5), "", rsClient!ClientAddressLine5) ' rsClient!ClientAddressLine5
    txtClientHomeTel(18).text = IIf(IsNull(rsClient!ClientOfficeAddressLine4), "", rsClient!ClientOfficeAddressLine4) 'rsClient!ClientOfficeAddressLine4
    txtClientHomeTel(19).text = IIf(IsNull(rsClient!ClientOfficeAddressLine5), "", rsClient!ClientOfficeAddressLine5) 'rsClient!ClientOfficeAddressLine4
    txtClientHomeTel(25).text = IIf(IsNull(rsClient!RegAdd4), "", rsClient!RegAdd4) ' rsClient!RegAdd4
    txtClientHomeTel(26).text = IIf(IsNull(rsClient!RegAdd5), "", rsClient!RegAdd5) ' rsClient!RegAdd4
    
    txtComments1.text = IIf(IsNull(rsClient!Comments1), "", rsClient!Comments1) 'rsClient!Comments1
    txtComments2(0).text = IIf(IsNull(rsClient!Comments2), "", rsClient!Comments2) 'rsClient!Comments2
    txtComments2(1).text = IIf(IsNull(rsClient!LesseeTemplate), "", rsClient!LesseeTemplate)
    txtComments2(2).text = IIf(IsNull(rsClient!LesseeAccTemplate), "", rsClient!LesseeAccTemplate)
    txtComments2(3).text = IIf(IsNull(rsClient!CSPreviewTemplate), "", rsClient!CSPreviewTemplate)
    txtComments2(4).text = IIf(IsNull(rsClient!CSTemplate), "", rsClient!CSTemplate)
    
    txtRenSummaryStatement.text = IIf(IsNull(rsClient!RentSummaryTemplate), "", rsClient!RentSummaryTemplate)
    txtRemittanceTemplate.text = IIf(IsNull(rsClient!RemittanceTemplate), "", rsClient!RemittanceTemplate) 'rsClient!Comments2
    chkConsolidatedStatement.Value = IIf(rsClient!ConsolidatedStatement = 0, 0, 1)
    chkClientAddress.Value = IIf(rsClient!StToClientAddress = 0, 0, 1)
    chkStatementAddress.Value = IIf(rsClient!StToStatementAddress = 0, 0, 1)
    rsClient.Close
    adoConn.Close
    Set adoConn = Nothing
    Fill_Form = True
End Function

Private Sub txtStopDatemngtFee_GotFocus()
        txtStopDatemngtFee.SelStart = 0
        txtStopDatemngtFee.SelLength = Len(txtStopDatemngtFee)
End Sub

Private Sub txtStopDatemngtFee_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If KeyAscii = 13 Then
        FocusControl txtCapAmount
    End If
    Dim KA As Integer
    KA = KeyAscii
    TextBoxKeyPrsDate txtStopDatemngtFee, KA
   ' DigitTextKeyPress txtStopDatemngtFee, KA
    KeyAscii = KA
End Sub

Private Sub txtStopDatemngtFee_LostFocus()
   TextBoxFormatDate txtStopDatemngtFee
End Sub
Private Sub txtStopDatemngtFee_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
''   MsTextBoxKeyPrsDate txtStopDatemngtFee, KeyCode
'    If KeyCode = 13 Then
''        FocusControl txtCapAmount
'        SendKeys vbTab
'    End If
    If KeyCode = 8 And (Len(txtStopDatemngtFee) = 3 Or Len(txtStopDatemngtFee) = 6) Then
            If Len(txtStopDatemngtFee) = 3 Then
                     txtStopDatemngtFee.text = Left(txtStopDatemngtFee.text, 2)
            ElseIf Len(txtStopDatemngtFee) = 6 Then
                      txtStopDatemngtFee.text = Left(txtStopDatemngtFee.text, 5)
             End If
            bBackSp = True
    Else
            bBackSp = False
    End If
    'How to fix the bug on / when you type a date field 1. txtEND_DATE_Change event add conditional exit sub  2. txtEND_DATE_KeyDown e all codes has be written  3. declare bBackSp variable for this form

End Sub

Private Sub txtStopDatemngtFee_Change()
'   TextBoxChangeDate txtStopDatemngtFee
   If (Len(txtStopDatemngtFee) = 2 Or Len(txtStopDatemngtFee) = 5) And bBackSp Then Exit Sub
   TextBoxChangeDate txtStopDatemngtFee
End Sub

Private Sub txtTotalAmountPerYear_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtPeriod
    End If
    Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtTotalAmountPerYear, KA
   KeyAscii = KA
End Sub



Private Sub RCAnnual()
           Dim Area As String, Total As Double
           Dim Rst1 As New ADODB.Recordset
           Dim Conn1 As New ADODB.Connection
           Dim szSQL  As String
           txtAmount.text = Format(IIf(txtAmount = "", 0, txtAmount), "0.00")
        
           Total = CDbl(Val(txtAmount))
           txtTotalAmountPerYear.text = Format(Total, "0.00")
        
           Conn1.Open getConnectionString
           If txtFrequecymngtFee.Tag <> "" Then
              szSQL = "SELECT PARTOFYEAR " & _
                        "FROM FREQUENCIES " & _
                        "WHERE ID = " & txtFrequecymngtFee.Tag & ";"
           Else
              szSQL = "SELECT PARTOFYEAR " & _
                        "FROM FREQUENCIES " & _
                        "WHERE ID = " & txtFrequecymngtFee.Tag & ";"
           End If
           Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
        
           txtPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")
        
           Rst1.Close
           Conn1.Close
           Set Rst1 = Nothing
           Set Conn1 = Nothing
End Sub
Private Sub RCPercentage()
        Dim Total As Double, TotalRentCharge As String ', temp() As String
        Dim Conn1 As New ADODB.Connection
        Dim TotalRentCharge1 As String
        Dim TotalRentCharge2 As String
        Dim TotalRentCharge3 As String
        Dim Rst1 As New ADODB.Recordset
        Dim szSQL As String
           ' The following code is to calculate fees and charges payable according to percentage
           Conn1.Open getConnectionString

           TotalRentCharge3 = GetGlobalTotalRC(Conn1, szPropertySelection1, txtFundMngtFee.Tag)
            If TotalRentCharge3 <= 0 Then
                TotalRentCharge3 = GetGlobalTotalSC(Conn1, szPropertySelection1, txtFundMngtFee.Tag)
                If TotalRentCharge3 <= 0 Then
                     TotalRentCharge3 = GetGlobalTotalIC(Conn1, szPropertySelection1, txtFundMngtFee.Tag)
                End If
            End If
          
           If TotalRentCharge3 <= 0 Then
                     MsgBox "   A budget has not been set for '" & strSelectedFundName & "' fund." & (Chr(13) + Chr(10)) & _
                     "Please enter a budget for '" & strSelectedFundName & "' fund.", vbInformation, "Enter Budget"
           Else
              Total = CDbl(TotalRentCharge3) * (CDbl(Val(txtAmount.text)) / 100)
              txtTotalAmountPerYear.text = Format(Total, "0.00")
        '      temp = Split(cboFreqBR.text, "-")
              szSQL = "SELECT PARTOFYEAR " & _
                        "FROM FREQUENCIES " & _
                        "WHERE ID = " & txtFrequecymngtFee.Tag & ";"
              Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

              txtPeriod.text = Format((Total / CInt(Rst1!PartOfYear)), "0.00")

              Rst1.Close
              Set Rst1 = Nothing

              txtAmount.text = Format(IIf(txtAmount.text = "", 0, txtAmount.text), "0.00")
           End If

           Conn1.Close
           Set Conn1 = Nothing
           Exit Sub

ErrorHander:
           Rst1.Close
           Conn1.Close
           Set Rst1 = Nothing
           Set Conn1 = Nothing
End Sub

Public Function GetGlobalTotalIC(adoConn As ADODB.Connection, szPropertySelection1 As String, szFund As String) As Double
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String
'number 1
   On Error GoTo ErrorHanlder

'Samrat 18/09/2014    Financial year has been implemented
   szSQL = "SELECT GlobalInsurance.Amount as TRC " & _
           "FROM GlobalInsurance,  Property " & _
           "WHERE " & _
                 "Property.PropertyID = '" & szPropertySelection1 & "' AND " & _
                 "Property.CBY = GlobalInsurance.FinancialYear AND " & _
                 "GlobalInsurance.FundType = " & szFund & ";"
'Debug.Print szSQL
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not IsNull(rstRst.Fields.Item("TRC").Value) Then
      GetGlobalTotalIC = CDbl(rstRst.Fields.Item("TRC").Value)
   Else
      GetGlobalTotalIC = -1
   End If

   rstRst.Close
   Set rstRst = Nothing
   Exit Function

ErrorHanlder:

   GetGlobalTotalIC = -1

   rstRst.Close
   Set rstRst = Nothing
End Function

Public Function GetGlobalTotalSC(ByVal rdoConn As ADODB.Connection, szPropertySelection1 As String, szFund As String) As Double
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String
'number 2
   On Error GoTo ErrorHanlder

'Samrat 18/09/2014  Financial year has implemented
   szSQL = "SELECT GlobalSC.TotalBudget as SC " & _
           "FROM GlobalSC, Units, Property " & _
           "WHERE " & _
                 "Property.PropertyID = '" & szPropertySelection1 & "' AND " & _
                 "Property.CBY = GlobalSC.FinancialYear AND " & _
                 "GlobalSC.Fund = " & szFund & ";"
'Debug.Print szSQL
   rstRst.Open szSQL, rdoConn, adOpenStatic, adLockReadOnly

   If Not IsNull(rstRst!SC) Then
      GetGlobalTotalSC = CDbl(rstRst!SC)
   Else
      GetGlobalTotalSC = -1
   End If

   rstRst.Close
   Set rstRst = Nothing
   Exit Function

ErrorHanlder:

   GetGlobalTotalSC = -1

   rstRst.Close
   Set rstRst = Nothing
End Function


Public Function GetGlobalTotalRC(ByVal adoConn As ADODB.Connection, szPropertySelection1 As String, szFund As String) As Double
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String
'number 3
   On Error GoTo ErrorHanlder

   szSQL = "SELECT GlobalRC.TotalBudget as TRC " & _
           "FROM GlobalRC, Property " & _
           "WHERE " & _
                 "Property.PropertyID = '" & szPropertySelection1 & "' AND " & _
                 "Property.CBY = GlobalRC.FinancialYear AND " & _
                 "GlobalRC.Fund = '" & szFund & "';"
'Debug.Print szSQL
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not IsNull(rstRst!TRC) Then
      GetGlobalTotalRC = CDbl(rstRst!TRC)
   Else
      GetGlobalTotalRC = -1
   End If

   rstRst.Close
   Set rstRst = Nothing
   Exit Function

ErrorHanlder:

   GetGlobalTotalRC = -1

   rstRst.Close
   Set rstRst = Nothing
End Function

Private Sub txtTotalAmountPerYear_LostFocus()
    If IsNumeric(txtTotalAmountPerYear) Then
        txtTotalAmountPerYear.text = Format(txtTotalAmountPerYear, "0.00")
    End If
    
End Sub

Private Sub txtVATReg_KeyPress(KeyAscii As Integer)
       'If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 8 Then KeyAscii = 0
   Dim KA As Integer
   If KeyAscii = 13 Then
        FocusControl chkOptedtoTax
   End If
   KA = KeyAscii
   DigitTextKeyPress txtVATReg, KA
   KeyAscii = KA
End Sub

Private Sub txtYearEndDate_Change()
   TextBoxChangeDate txtYearEndDate
End Sub

Private Sub txtYearEndDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtYearEndDate, KeyAscii
End Sub

Private Sub txtYearEndDate_LostFocus()
   TextBoxFormatDate txtYearEndDate
End Sub

Private Sub AgreementButtonMode(ByVal mode As ComponentMode) 'Management Fee
   Select Case mode

   Case ComponentMode.DefaultMode
        cmdCommandArray(9).Enabled = False
        cmdCommandArray(8).Enabled = False
      
        cmdCommandArray(7).Enabled = False
        cmdCommandArray(6).Enabled = False
        cmdCommandArray(10).Enabled = False
        cmdCommandArray(11).Enabled = False
        cmdCommandArray(12).Enabled = False
        txtSTART_DATE.Locked = True
        txtNtDueDate.Locked = True
        txtTotalAmountPerYear.Locked = True
        txtAmount.Locked = True
        txtPeriod.Locked = True
        txtStopDatemngtFee.Locked = True
        txtLastChargeDate.Locked = True
        txtCapAmount.Locked = True
        txtEND_DATE.Locked = True
        cmdAgmntAddNew.Enabled = True
        cmdAgmntEdit.Enabled = False
        cmdAgmntSave.Enabled = False
        cmdAgmntCancel.Enabled = False
        flxManagementFee.Enabled = True

   Case ComponentMode.NewEntryMode
        cmdDeleteMgtFee.Enabled = False
        flxManagementFee.Enabled = True
        cmdCommandArray(9).Enabled = True
        cmdCommandArray(8).Enabled = True
        txtSTART_DATE.Locked = False
        cmdCommandArray(7).Enabled = True
        cmdCommandArray(6).Enabled = True
        cmdCommandArray(10).Enabled = True
        cmdCommandArray(11).Enabled = True
        cmdCommandArray(12).Enabled = True
        
        txtNtDueDate.Locked = False
        txtTotalAmountPerYear.Locked = False
        txtAmount.Locked = False
        txtPeriod.Locked = False
        txtStopDatemngtFee.Locked = False
        txtCapAmount.Locked = False
        txtEND_DATE.Locked = False
        cmdAgmntAddNew.Enabled = False
        cmdAgmntEdit.Enabled = False
        cmdAgmntSave.Enabled = True
        cmdAgmntCancel.Enabled = True
        txtLastChargeDate.Locked = False

   Case ComponentMode.EditMode
        cmdCommandArray(9).Enabled = True
        cmdCommandArray(8).Enabled = True
        txtSTART_DATE.Locked = False
        txtEND_DATE.Locked = False
        cmdCommandArray(7).Enabled = True
        cmdCommandArray(6).Enabled = True
        cmdCommandArray(10).Enabled = True
        cmdCommandArray(11).Enabled = True
        cmdCommandArray(12).Enabled = True
        txtNtDueDate.Locked = False
        txtTotalAmountPerYear.Locked = False
        txtAmount.Locked = False
        txtPeriod.Locked = False
        txtStopDatemngtFee.Locked = False
       
        txtCapAmount.Locked = False
        txtEND_DATE.Locked = False
        cmdAgmntAddNew.Enabled = False
        cmdAgmntEdit.Enabled = False
        cmdAgmntSave.Enabled = True
        cmdAgmntCancel.Enabled = True
        cmdDeleteMgtFee.Enabled = False
   End Select
End Sub

Private Sub CPAButtonMode(ByVal mode As ComponentMode)
   Select Case mode

   Case ComponentMode.DefaultMode
     
        txtNoOfDaysToSendMFB4Due.Locked = True
        
        txtAgreementStartDate.Enabled = False
        txtAgreementEndDate.Enabled = False
        txtREVIEW_DATE.Enabled = False
'      txtAgreementStartDate.Locked = True
'      txtAgreementEndDate.Locked = True
'      txtREVIEW_DATE.Locked = True

'      flxAgreement.Enabled = True
      flxPayable.Enabled = True

      cmdAgrTopSave.Enabled = True
      cmdAgrTopEdit.Enabled = False

   Case ComponentMode.NewEntryMode
'      txtREVIEW_DATE.Locked = False
'      txtAgreementStartDate.Locked = False
'      txtAgreementEndDate.Locked = False

 txtAgreementStartDate.Enabled = True
            txtAgreementEndDate.Enabled = True
            txtREVIEW_DATE.Enabled = True
            
      txtNoOfDaysToSendMFB4Due.Locked = False

'      flxAgreement.Enabled = False
      flxPayable.Enabled = False

      cmdAgrTopSave.Enabled = False
      cmdAgrTopEdit.Enabled = False

   Case ComponentMode.EditMode
'      txtREVIEW_DATE.Locked = False
'      txtAgreementStartDate.Locked = False
'      txtAgreementEndDate.Locked = False

            txtAgreementStartDate.Enabled = True
            txtAgreementEndDate.Enabled = True
            txtREVIEW_DATE.Enabled = True
            
      txtNoOfDaysToSendMFB4Due.Locked = False

'      flxAgreement.Enabled = False
      flxPayable.Enabled = False

      cmdAgrTopSave.Enabled = False
      cmdAgrTopEdit.Enabled = False
   End Select
End Sub

Private Sub PayableButtonMode(ByVal mode As ComponentMode)
   Select Case mode

   Case ComponentMode.DefaultMode
      txtPayableType.Locked = True
'      txtPayDemandType.Locked = True
'      cboPAYABLE_METHOD.Locked = True
'      txtPAY_AnnualCharge.Locked = True
'      txtPAY_START_DATE.Locked = True
'      txtPAY_END_DATE.Locked = True
      txtPayFund.Locked = True
'      cboPAY_HANDLING.Locked = True
'      txtPayFrequency.Locked = True
       txtStopDate.Locked = True
      flxPayable.Enabled = True
'      txtPAY_NtDueDate.Locked = False

      cmdPayAddNew.Enabled = True
      cmdPayEdit.Enabled = False
      cmdPaySave.Enabled = False
      cmdPayCancel.Enabled = False
      cmdCommandArray(0).Enabled = False
      cmdCommandArray(1).Enabled = False
      cmdCommandArray(2).Enabled = False
      cmdCommandArray(3).Enabled = False
'      cmdCommandArray(4).Enabled = False
      cmdCommandArray(5).Enabled = False

   Case ComponentMode.NewEntryMode
      txtPayableType.Locked = False
'      txtPayDemandType.Locked = False
'      cboPAYABLE_METHOD.Locked = False
'      txtPAY_AnnualCharge.Locked = False
'      txtPAY_START_DATE.Locked = False
'      txtPAY_END_DATE.Locked = False
      txtPayFund.Locked = False
'      cboPAY_HANDLING.Locked = False
'      txtPayFrequency.Locked = False
        txtStopDate.Locked = False
      flxPayable.Enabled = False

      cmdPayAddNew.Enabled = False
      cmdPayEdit.Enabled = False
      cmdPaySave.Enabled = True
      cmdPayCancel.Enabled = True
      cmdCommandArray(0).Enabled = True
      cmdCommandArray(1).Enabled = True
      cmdCommandArray(2).Enabled = True
      cmdCommandArray(3).Enabled = True
      'cmdCommandArray(4).Enabled = True
      cmdCommandArray(5).Enabled = True

   Case ComponentMode.EditMode
      txtPayableType.Locked = False
'      txtPayDemandType.Locked = False
'      cboPAYABLE_METHOD.Locked = False
'      txtPAY_AnnualCharge.Locked = False
'      txtPAY_START_DATE.Locked = False
'      txtPAY_END_DATE.Locked = False
      txtPayFund.Locked = False
'      cboPAY_HANDLING.Locked = False
'      txtPayFrequency.Locked = False
      txtStopDate.Locked = False
      flxPayable.Enabled = False
'      txtPAY_NtDueDate.Locked = False
      If txtPayableBasis.Tag = "FA" Then
            txtPercentage.Locked = True
      Else
            txtPercentage.Locked = False
      End If
      

      cmdPayAddNew.Enabled = False
      cmdPayEdit.Enabled = False
      cmdPaySave.Enabled = True
      cmdPayCancel.Enabled = True
      cmdCommandArray(0).Enabled = True
      cmdCommandArray(1).Enabled = True
      cmdCommandArray(2).Enabled = True
      cmdCommandArray(3).Enabled = True
'      cmdCommandArray(4).Enabled = True
      cmdCommandArray(5).Enabled = True
   End Select
End Sub

Private Sub AgreementClearMode(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
        txtChargeType.text = ""
        txtChargeType.Tag = ""
'        txtDemandTypemngtFee.text = ""
        txtSTART_DATE.text = ""
        txtEND_DATE.text = ""
        txtFundMngtFee.text = ""
        txtNtDueDate.text = ""
        strtlbAgreementID = 0
'        txtDemandTypemngtFee.Tag = ""
        txtFundMngtFee.Tag = ""
        txtFundMngtFee.text = ""
        txtManagingAgentAC.Tag = ""
        txtManagingAgentAC.text = ""
        txtChargingMethod.text = ""
        txtChargingMethod.Tag = ""
        txtChargeBasis.Tag = ""
        txtChargeBasis.text = ""
        txtSTART_DATE = ""
        txtFrequecymngtFee.Tag = ""
        txtFrequecymngtFee.text = ""
        txtNtDueDate.text = ""
        txtTotalAmountPerYear.text = ""
        txtPeriod.text = ""
        txtStopDatemngtFee.text = ""
        txtCapAmount.text = ""
        txtEND_DATE.text = ""
        txtAmount.text = ""
        txtLastChargeDate.text = ""
    
    

   Case CearEntryComponents.ClearOnlyComboBoxes
'      cboCHARGE_BASIS.Clear
'      cboCHARGE_METHOD.Clear
''      cboFund.Clear
'      cboHandling.Clear
'      cboFrequency.Clear

   Case CearEntryComponents.ClearBoth
      AgreementClearMode ClearOnlyTextBoxes
      AgreementClearMode ClearOnlyComboBoxes
   End Select
End Sub

Private Sub CPAClearMode(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
      txtREVIEW_DATE.text = ""
      txtNoOfDaysToSendMFB4Due.text = ""

   Case CearEntryComponents.ClearBoth
      AgreementClearMode ClearOnlyTextBoxes
   End Select
End Sub

Private Sub PayableClearMode(ByVal mode As CearEntryComponents)
   Select Case mode

   Case CearEntryComponents.ClearOnlyTextBoxes
      txtPayableType.text = ""
'      txtPayDemandType.text = ""
'      cboPAYABLE_METHOD.text = ""
'      txtPAY_AnnualCharge.text = ""
'      txtPAY_START_DATE.text = ""
'      txtPAY_END_DATE.text = ""
      txtPayFund.text = ""
'      cboPAY_HANDLING.text = ""
'      txtPayFrequency.text = ""
'      txtPAY_NtDueDate.text = ""
 'new code by anol 20210823
        txtClientLandlord.text = ""
        txtClientLandlord.Tag = ""
        txtPayeeType.text = ""
        txtPayeeType.Tag = ""
        txtPayableBasis.text = ""
        txtPayableBasis.Tag = ""
        txtPercentage.text = ""
        txtPercentage.Tag = ""

   Case CearEntryComponents.ClearOnlyComboBoxes
'      cboPAYABLE_METHOD.Clear
'      cboPAY_Fund.Clear
'      cboPAY_HANDLING.Clear
'      txtPayFrequency.Clear

   Case CearEntryComponents.ClearBoth
      PayableClearMode ClearOnlyTextBoxes
      PayableClearMode ClearOnlyComboBoxes
   End Select
End Sub

Private Sub loadpayableTypes()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Payable Type"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
   
   

   adoConn.Open getConnectionString
   szSQL = "SELECT ID, PayType FROM PayableTypes where PropertyID='" & szPropertySelection1 & "';"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("PayType").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub loadClientLandlord()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid column will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = " ID"
   lblClientID(1).Caption = " Name"
   lblClientID(2).Caption = " Type"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
'Load landlord which is only related to the property
   adoConn.Open getConnectionString
   If txtPayeeType.text = "Client" Then
        szSQL = "SELECT SupplierID, SupplierName,'CLIENT' as STYPE FROM Supplier S " & _
            "where Type in('CLIENT') AND SupplierID='" & txtClientID.text & "' ORDER BY TYPE,SupplierID ;"
   Else
        szSQL = "SELECT SupplierID, SupplierName,iif(Type='LLORD','LANDLORD','CLIENT') as STYPE FROM Supplier S, PropertyLandlord P " & _
            "where P.LandlordID=S.SupplierID AND P.PropertyID='" & szPropertySelection1 & "' AND Type in('LLORD') ORDER BY TYPE,SupplierID ;"
   End If
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        rRow = 1
        If rstRec.EOF Then
            MsgBox "Please setup Landlord for the selected property", vbInformation, "Warning "
            rstRec.Close
            adoConn.Close
            Set rstRec = Nothing
            Set adoConn = Nothing
            Exit Sub
        End If
        Dim bResult As Boolean
        While Not rstRec.EOF
'            bResult = InTheList(rstRec.Fields.Item("SupplierID").Value)
'            If Not bResult Then
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("SupplierID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("SupplierName").Value
                    flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("SType").Value
                    flxClientList.RowHeight(rRow) = 280
'            End If
            rstRec.MoveNext
            If Not bResult Then
                If Not rstRec.EOF Then flxClientList.AddItem ""
                rRow = rRow + 1
            End If
         Wend
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Function InTheList(strLandlordID As String) As Boolean
    ' flxPayable.TextMatrix(rRow, 9) = rsPayable.Fields.Item("clientLandlordID").Value
    Dim iRow As Integer
    For iRow = 1 To flxPayable.Rows - 1
        If flxPayable.TextMatrix(iRow, 9) = strLandlordID Then
                InTheList = True
                Exit Function
        End If
    Next
End Function
Private Sub LoadLessee()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid column will try to fit accordingly
 flxClientList.Clear
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Clear
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = " ID"
   lblClientID(1).Caption = " Name"
   lblClientID(2).Caption = " Type"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoConn.Open getConnectionString
   If Len(txtSearchClientName.text) > 0 Then
        szSQL = "SELECT TenantID, Name FROM Tenants T,Units U,Property P and P.clientID='" & txtClientID.text & "' and isnull(Comments) AND " & _
             "' T.currUnit=U.unitNumber AND P.PropertyID=U.PropertyID AND Name like'%" & txtSearchClientName.text & "%' ORDER BY TenantID ;"
   ElseIf Len(txtSearchClientID.text) > 0 Then
         szSQL = "SELECT TenantID, Name FROM T,Units U,Property P and P.clientID='" & txtClientID.text & "' and isnull(Comments) AND " & _
             "' T.currUnit=U.unitNumber AND P.PropertyID=U.PropertyID" & _
             "'AND TenantID like'%" & txtSearchClientID.text & "%' ORDER BY TenantID ;"
   Else
        szSQL = "SELECT TenantID, Name FROM Tenants T,Units U,Property P where P.clientID='" & txtClientID.text & "' and isnull(Comments) AND " & _
             " T.currUnit=U.unitNumber AND P.PropertyID=U.PropertyID ORDER BY TenantID ;"
   End If
   
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        rRow = 1
        flxClientList.TextMatrix(rRow, 1) = "ALL"
        flxClientList.TextMatrix(rRow, 2) = "ALL"
         flxClientList.RowHeight(rRow) = 280
        flxClientList.AddItem ""
        rRow = 2
        While Not rstRec.EOF
            flxClientList.row = 1
            flxClientList.RowSel = 1
            flxClientList.ColSel = 1
            flxClientList.TextMatrix(rRow, 0) = ""
            flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("TenantID").Value
            flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Name").Value
            flxClientList.TextMatrix(rRow, 3) = ""
            flxClientList.RowHeight(rRow) = 280
            rstRec.MoveNext
            If Not rstRec.EOF Then flxClientList.AddItem ""
            rRow = rRow + 1
         Wend
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadSupplier(strType As String)
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid column will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Clear
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = " ID"
   lblClientID(1).Caption = " Name"
   lblClientID(2).Caption = " Type"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoConn.Open getConnectionString
   If Len(txtSearchClientName.text) > 0 Then
        szSQL = "SELECT SupplierID, SupplierName,iif(Type='AGENT','Managing Agent','') as STYPE FROM Supplier where Type in('" & _
            strType & "') AND SupplierName like'%" & txtSearchClientName.text & "%' ORDER BY TYPE,SupplierID ;"
   ElseIf Len(txtSearchClientID.text) > 0 Then
        szSQL = "SELECT SupplierID, SupplierName,iif(Type='AGENT','Managing Agent','') as STYPE FROM Supplier where Type in('" & _
        strType & "') AND SupplierID like'%" & txtSearchClientID.text & "%'  ORDER BY TYPE,SupplierID ;"
   Else
        szSQL = "SELECT SupplierID, SupplierName,iif(Type='AGENT','Managing Agent','') as STYPE FROM Supplier where Type in('" & _
        strType & "') ORDER BY TYPE,SupplierID ;"
   End If
   
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        rRow = 1
        flxClientList.TextMatrix(rRow, 1) = "ALL"
        flxClientList.TextMatrix(rRow, 2) = "ALL"
         flxClientList.RowHeight(rRow) = 280
        flxClientList.AddItem ""
        rRow = 2
        While Not rstRec.EOF
            flxClientList.row = 1
            flxClientList.RowSel = 1
            flxClientList.ColSel = 1
            flxClientList.TextMatrix(rRow, 0) = ""
            flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("SupplierID").Value
            flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("SupplierName").Value
            flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("SType").Value
            flxClientList.RowHeight(rRow) = 280
            rstRec.MoveNext
            If Not rstRec.EOF Then flxClientList.AddItem ""
            rRow = rRow + 1
         Wend
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadManagingAgent()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
 'you just change label position then searchbox and grid column will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 4510
   flxClientList.Cols = 4
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = " ID"
   lblClientID(1).Caption = " Name"
   lblClientID(2).Caption = " Type"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoConn.Open getConnectionString
   szSQL = "SELECT SupplierID, SupplierName,iif(Type='AGENT','Managing Agent','') as STYPE FROM Supplier where Type in('AGENT') ORDER BY TYPE,SupplierID ;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        rRow = 1
        While Not rstRec.EOF
            flxClientList.row = 1
            flxClientList.RowSel = 1
            flxClientList.ColSel = 1
            flxClientList.TextMatrix(rRow, 0) = ""
            flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("SupplierID").Value
            flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("SupplierName").Value
            flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("SType").Value
            flxClientList.RowHeight(rRow) = 280
            rstRec.MoveNext
            If Not rstRec.EOF Then flxClientList.AddItem ""
            rRow = rRow + 1
         Wend
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadPayeeTypes()
       Dim rRow As Integer
       Dim szSQL As String
    
       Dim adoConn As New ADODB.Connection
       Dim rstRec As New ADODB.Recordset
    'This is Ideal gridview Popup which is written by anol 2020-18-12
    'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
       lblClientID(0).Left = 250
       lblClientID(1).Left = 1365
       lblClientID(2).Left = 3510
       flxClientList.Cols = 3
       
       flxClientList.RowHeight(0) = 0
       flxClientList.ColWidth(0) = 200
       flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
       If flxClientList.Cols > 3 Then
            flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
            txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
       ElseIf flxClientList.Cols = 3 Then
            flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
            txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
       End If
       If flxClientList.Cols = 4 Then
            flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
            TextBox1.Visible = True
       ElseIf flxClientList.Cols = 3 Then
            flxClientList.ColWidth(3) = 0
            TextBox1.Visible = False
       End If
       txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
       TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
       txtSearchClientName.Visible = True
    
       
       flxClientList.Clear
       flxClientList.Rows = 2
       flxClientList.ColAlignment(0) = vbLeftJustify
       flxClientList.ColAlignment(1) = vbLeftJustify
       flxClientList.ColAlignment(2) = vbLeftJustify
       If flxClientList.Cols > 3 Then
            flxClientList.ColAlignment(3) = vbLeftJustify
       End If
       
       lblClientID(0).Caption = "PayeeTypes"
       lblClientID(1).Caption = "PayeeTypes"
       lblClientID(2).Caption = ""
       
       txtSearchClientID.Left = lblClientID(0).Left
       txtSearchClientName.Left = lblClientID(1).Left
       TextBox1.Left = lblClientID(2).Left
       TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
       
       txtSearchClientName.text = ""
       txtSearchClientID.text = ""
       TextBox1.text = ""
    
            flxClientList.TextMatrix(1, 1) = "Client"
            flxClientList.TextMatrix(1, 2) = "Client"
            flxClientList.RowHeight(1) = 280
            flxClientList.AddItem ""
            
            flxClientList.TextMatrix(2, 1) = "Landlord"
            flxClientList.TextMatrix(2, 2) = "Landlord"
            flxClientList.RowHeight(2) = 280
'       adoConn.Open getConnectionString
'       szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
'       rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'                    rRow = 1
'                    While Not rstRec.EOF
'                        flxClientList.row = 1
'                        flxClientList.RowSel = 1
'                        flxClientList.ColSel = 1
'                        flxClientList.TextMatrix(rRow, 0) = ""
'                        flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
'                        flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("TYPE").Value
'                        flxClientList.RowHeight(rRow) = 280
'                        rstRec.MoveNext
'                        If Not rstRec.EOF Then flxClientList.AddItem ""
'                        rRow = rRow + 1
'                     Wend
'
'
'       rstRec.Close
'       adoConn.Close
'       Set rstRec = Nothing
'       Set adoConn = Nothing
End Sub
Private Sub LoadDemandTypes()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
'This is Ideal gridview Popup which is written by anol 2020-18-12
'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Demand Type"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoConn.Open getConnectionString
   szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("TYPE").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub loadFrequencies()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
'This is Ideal gridview Popup which is written by anol 2020-18-12
'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Frequency"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoConn.Open getConnectionString
   szSQL = "SELECT ID, Frequency FROM Frequencies;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Frequency").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
      
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadChargeBasis()
    
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
'This is Ideal gridview Popup which is written by anol 2020-18-12
'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "CODE"
   lblClientID(1).Caption = "Charge Basis"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   If txtChargingMethod.text = "FIXED" Then
        
        flxClientList.TextMatrix(1, 0) = ""
        flxClientList.TextMatrix(1, 1) = "AN"
        flxClientList.TextMatrix(1, 2) = "Annual"
        flxClientList.RowHeight(1) = 280
        flxClientList.AddItem ""
        
        flxClientList.TextMatrix(2, 0) = ""
        flxClientList.TextMatrix(2, 1) = "PC"
        flxClientList.TextMatrix(2, 2) = "Percentage"
        flxClientList.RowHeight(2) = 280
        flxClientList.AddItem ""
    
   ElseIf txtChargingMethod.text = UCase("Receivable") Or txtChargingMethod.text = UCase("Received") Then
        flxClientList.TextMatrix(1, 0) = ""
        flxClientList.TextMatrix(1, 1) = "PC"
        flxClientList.TextMatrix(1, 2) = "Percentage"
        flxClientList.RowHeight(1) = 280
        flxClientList.AddItem ""
  
   End If
    
    
    
End Sub
Private Sub loadPayableBasis()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
'This is Ideal gridview Popup which is written by anol 2020-18-12
'you just change label position & cols then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510
   flxClientList.Cols = 3
   
   flxClientList.RowHeight(0) = 0
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   If flxClientList.Cols > 3 Then
        flxClientList.ColAlignment(3) = vbLeftJustify
   End If
   
   lblClientID(0).Caption = "CODE"
   lblClientID(1).Caption = "Payable Basis"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
'check in the grid if any entry then show only one of this item
    Dim iRow As Integer
    Dim strHasvalue As String
'    For iRow = 1 To flxPayable.Rows - 1
'       If flxPayable.TextMatrix(iRow, 14) = "TA" Or flxPayable.TextMatrix(iRow, 14) = "PC" Then
'                strHasvalue = flxPayable.TextMatrix(iRow, 14)
'                Exit For
'       End If
'    Next
'    If strHasvalue = "" Then
'        flxClientList.TextMatrix(1, 0) = ""
'        flxClientList.TextMatrix(1, 1) = "TA"
'        flxClientList.TextMatrix(1, 2) = "Total Amount"
'        flxClientList.RowHeight(1) = 280
'        flxClientList.AddItem ""
'
'        flxClientList.TextMatrix(2, 0) = ""
'        flxClientList.TextMatrix(2, 1) = "PC"
'        flxClientList.TextMatrix(2, 2) = "Percentage"
'        flxClientList.RowHeight(2) = 280
'    ElseIf strHasvalue = "TA" Then
'            flxClientList.AddItem ""
'            flxClientList.TextMatrix(1, 0) = ""
'            flxClientList.TextMatrix(1, 1) = "TA"
'            flxClientList.TextMatrix(1, 2) = "Total Amount"
'            flxClientList.RowHeight(1) = 280
'            flxClientList.RemoveItem 2
'    ElseIf strHasvalue = "PC" Then
            flxClientList.AddItem ""
            flxClientList.TextMatrix(1, 0) = ""
            flxClientList.TextMatrix(1, 1) = "PC"
            flxClientList.TextMatrix(1, 2) = "Percentage"
            flxClientList.RowHeight(1) = 280
            'flxClientList.RemoveItem 2
             flxClientList.AddItem ""
             flxClientList.TextMatrix(2, 0) = ""
            flxClientList.TextMatrix(2, 1) = "FA"
            flxClientList.TextMatrix(2, 2) = "FULL Amount"
            flxClientList.RowHeight(2) = 280
            
'    End If
    
                    
'   adoconn.Open getConnectionString
'   szSQL = "SELECT ID, Frequency FROM Frequencies;"
'   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'                rRow = 1
'                While Not rstRec.EOF
'                    flxClientList.row = 1
'                    flxClientList.RowSel = 1
'                    flxClientList.ColSel = 1
'                    flxClientList.TextMatrix(rRow, 0) = ""
'                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("ID").Value
'                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Frequency").Value
'                    flxClientList.RowHeight(rRow) = 280
'                    rstRec.MoveNext
'                    If Not rstRec.EOF Then flxClientList.AddItem ""
'                    rRow = rRow + 1
'                 Wend
'
'
'   rstRec.Close
'   adoconn.Close
'   Set rstRec = Nothing
'   Set adoconn = Nothing
End Sub
Private Sub LoadFunds()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 4
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "ID"
   lblClientID(1).Caption = "Fund Code"
   lblClientID(2).Caption = "Fund Name"
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoConn.Open getConnectionString
   szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoConn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
   Else
        iSel = 1
        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
                szPropertySelection1 & "' and ClientID='" & txtClientID.text & "' and isDeleted=false"
   End If
   rsFundMatrix.Close
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxClientList.Clear
      flxClientList.Rows = 2
   Else
            rRow = 1
            flxClientList.TextMatrix(rRow, 0) = ""
            flxClientList.TextMatrix(rRow, 1) = ""
            flxClientList.TextMatrix(rRow, 2) = "N/A"
            flxClientList.TextMatrix(rRow, 3) = "N/A"
            flxClientList.RowHeight(rRow) = 280
            flxClientList.AddItem ""
            rRow = 2
                While Not rstRec.EOF
'                    flxClientList.row = 1
'                    flxClientList.RowSel = 1
'                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundID").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundCode").Value
                    flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundName").Value
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadflxBankAccountFund()
    Dim adoConn As New ADODB.Connection
    Dim rsFundMatrix As New ADODB.Recordset
    Dim rstRec As New ADODB.Recordset
    Dim szSQL As String
    Dim iSel As Long
    Dim rRow As Long
        flxBankAccountFund.Clear
        flxBankAccountFund.Rows = 2
    flxBankAccountFund.RowHeight(0) = 0
        flxBankAccountFund.Cols = 4
        flxBankAccountFund.ColWidth(0) = 280
        flxBankAccountFund.ColWidth(1) = 0
        flxBankAccountFund.ColWidth(2) = 1000
        flxBankAccountFund.ColWidth(3) = 2700
    flxBankAccountFund.ColAlignment(0) = vbLeftJustify
        flxBankAccountFund.ColAlignment(1) = vbLeftJustify
        flxBankAccountFund.ColAlignment(2) = vbLeftJustify
        flxBankAccountFund.ColAlignment(3) = vbLeftJustify
    adoConn.Open getConnectionString

'  rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoConn, adOpenStatic, adLockReadOnly
'       If rsFundMatrix("isfundAssign").Value = False Then
'        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
'   Else
'        iSel = 1
'        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
'                szPropertySelection1 & "' and ClientID='" & txtClientID.text & "' and isDeleted=false"
'   End If
'   rsFundMatrix.Close
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxBankAccountFund.Clear
      flxBankAccountFund.Rows = 2
   Else

            rRow = 1
                While Not rstRec.EOF
                    flxBankAccountFund.TextMatrix(rRow, 0) = ""
                    flxBankAccountFund.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundID").Value
                    flxBankAccountFund.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundCode").Value
                    flxBankAccountFund.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundName").Value
                    flxBankAccountFund.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxBankAccountFund.AddItem ""
                    rRow = rRow + 1
                 Wend
   End If
   If txtNCCODE.text <> "" Then
        'all funds are loaded. now clear all flagging and mark the same what is in the BankFundTable
        'Dim adoConn As New adodb.Connection
       ' adoConn.Open getConnectionString
        Dim rsBankFund As New ADODB.Recordset
        Dim iRow As Long
        For iRow = 1 To flxBankAccountFund.Rows - 1
            flxBankAccountFund.TextMatrix(iRow, 0) = ""
        Next iRow
        rsBankFund.Open "Select * from BankFund where clientID='" & txtClientID.text & "' AND BankCode='" & txtNCCODE.text & "'", adoConn, adOpenStatic, adLockReadOnly
        While Not rsBankFund.EOF
            For iRow = 1 To flxBankAccountFund.Rows - 1
                If flxBankAccountFund.TextMatrix(iRow, 1) = rsBankFund("FundID").Value Then
                    flxBankAccountFund.TextMatrix(iRow, 0) = "X"
                End If
            Next iRow
            rsBankFund.MoveNext
        Wend
        rsBankFund.Close
        'adoConn.Close
   End If
    

End Sub

Public Sub PopulateCodes()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
Exit Sub
   adoConn.Open getConnectionString

   ' Charge Type
   sSQLQuery = "SELECT ID, FeeType FROM ChargeTypes;"

'   populateCombo adoconn, sSQLQuery, cboCHARGE_TYPE

   ' Payable Type 'this has been converted to gridview
'   sSQLQuery = "SELECT ID, PayType FROM PayableTypes;"
'
'   populateCombo adoconn, sSQLQuery, cboPAYABLE_TYPE

   ' Demand Type 'this has been converted to gridview
   sSQLQuery = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"

'   populateCombo adoconn, sSQLQuery, cboDEMAND_TYPE
   ' Payable Demand Type
'   populateCombo adoconn, sSQLQuery, cboPAY_DEMAND_TYPE

   ' Fund both - Charge and Payable
'   LoadFund

   ' Charge Basis
   sSQLQuery = "SELECT CODE, VALUE FROM SECONDARYCODE WHERE PRIMARYCODE = 'AMTP'"

'   populateCombo adoconn, sSQLQuery, cboCHARGE_BASIS
   ' Payable Basis
'   populateCombo adoConn, sSQLQuery, cboPAYABLE_BASIS

   ' Amt/%
   sSQLQuery = "SELECT CODE, VALUE FROM SECONDARYCODE WHERE PRIMARYCODE = 'CRGBS'"

'   populateCombo adoconn, sSQLQuery, cboCHARGE_METHOD
   ' Payable method
'   populateCombo adoconn, sSQLQuery, cboPAYABLE_METHOD

   ' Frequency
   sSQLQuery = "SELECT ID, Frequency FROM Frequencies;"

'   populateCombo adoconn, sSQLQuery, cboFrequency
   ' Pay Frequency
'   populateCombo adoconn, sSQLQuery, cboPAY_FREQUENCY

   adoConn.Close
   Set adoConn = Nothing
'   Dim Data(1, 1) As String
'   Data(0, 0) = "Total Amount"
'   Data(1, 0) = "1"
'   Data(0, 1) = "Percentage"
'   Data(1, 1) = "2"
'   cboPayableBasis.Column() = Data
End Sub

'Private Sub LoadFund()
'   ' Error Handler
'   On Error GoTo Error_Handler
'
'   Dim adoconn As ADODB.Connection
'   Dim rRow As Integer, iRec As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   Set adoconn = New ADODB.Connection
'   adoconn.Open getConnectionString
'
'   szSQL = "SELECT FundID, FundName " & _
'           "FROM Fund;"
'
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
'   Else
'      ReDim Data(2, adoRst.RecordCount) As String
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
'
'      cboPAY_Fund.Clear
'      cboPAY_Fund.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoconn = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   MsgBox "Error in Loading fund.", vbExclamation, "Loading Fund"
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoconn = Nothing
'End Sub

Private Sub ConfigFlxACHistory()
   Dim szHeader As String, iCol As Integer
   flxACHistory.Clear
   szHeader$ = "Sign|<No|<Type|<Date|<Reference|<Description|>Amount|>Balance|>Debit|>Credit|>Transaction ID"

   With flxACHistory
      .FormatString = szHeader$
      .Cols = 12
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = 230                                                       'Sign
      .ColWidth(1) = Label11(3).Left - Label11(2).Left - 420                        'No
      .ColWidth(2) = Label11(3).Left - Label11(2).Left - 180                        'SAGEACCOUNTNO
      .ColWidth(3) = Label11(3).Left - Label11(2).Left - 200                        'Type
      .ColWidth(4) = Label11(4).Left - Label11(3).Left                         'Date
      .ColWidth(5) = Label11(5).Left - Label11(4).Left                         'Reference
      .ColWidth(6) = Label11(6).Left - Label11(5).Left                         'Description
      .ColWidth(7) = Label11(7).Left - Label11(6).Left                         'Amount
      .ColWidth(8) = Label11(8).Left - Label11(7).Left                         'Balance
      .ColWidth(9) = Label11(9).Left - Label11(8).Left                         'Debit
      .ColWidth(10) = .ColWidth(8)                                              'Credit
      .ColWidth(11) = 0                                                        'Transaction ID
   End With
   ConfigFlxACHistorySplit
End Sub

Private Sub ConfigFlxACHistorySplit()
   Dim szHeader As String, iCol As Integer

   With flxACHistorySplit
      .Clear
      .Cols = 12
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = Label11(11).Left - Label11(10).Left                         'No
      .ColWidth(1) = Label11(12).Left - Label11(11).Left                         'Type
      .ColWidth(2) = Label11(13).Left - Label11(12).Left                         'Date
      .ColWidth(3) = Label11(14).Left - Label11(13).Left                         'Ref
      .ColWidth(4) = Label11(15).Left - Label11(14).Left                         'N/C
      .ColWidth(5) = Label11(16).Left - Label11(15).Left                         'Job
      .ColWidth(6) = Label11(17).Left - Label11(16).Left                         'Unit
      .ColWidth(7) = Label11(18).Left - Label11(17).Left                         'Fund
      .ColWidth(8) = Label11(19).Left - Label11(18).Left                         'Desc
      .ColWidth(9) = Label11(20).Left - Label11(19).Left                         'Total
      .ColWidth(10) = Label11(21).Left - Label11(20).Left                        'Debit
      .ColWidth(11) = .Width - Label11(21).Left - 100                            'Credit
   End With
End Sub
Private Sub loadflxPropertySelection1() '***************************************Load  PROPERTY Grid ******************************************
   Dim szSQL    As String
   Dim TotalRow As Integer
   Dim TotalCol As Integer
   Dim i        As Integer
   Dim j        As Integer
   

   On Error GoTo ErrorHandler
   configflxPropertySelection1
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rsProperty As New ADODB.Recordset
   
   szSQL = "SELECT PropertyID, PropertyName,ClientID, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyID;"


    rsProperty.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
    i = 1
    While Not rsProperty.EOF
        flxPropertySelection1.AddItem ""
        flxPropertySelection1.TextMatrix(i, 0) = ""
        flxPropertySelection1.TextMatrix(i, 1) = rsProperty("PropertyID").Value
        flxPropertySelection1.TextMatrix(i, 2) = rsProperty("PropertyName").Value
        flxPropertySelection1.TextMatrix(i, 3) = rsProperty("ClientID").Value
        i = i + 1
        rsProperty.MoveNext
    Wend
    If i > 1 Then
        SelectOnly1RowFlxGrid flxPropertySelection1, 1, 0
        szPropertySelection1 = flxPropertySelection1.TextMatrix(1, 1) 'saving the first propertyID in the list in a variable
    End If
    rsProperty.Close
    Set rsProperty = Nothing
    Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
End Sub
Private Sub loadflxPropertySelection2() '***************************************Load  PROPERTY Grid ******************************************
   Dim szSQL    As String
   Dim TotalRow As Integer
   Dim TotalCol As Integer
   Dim i        As Integer
   Dim j        As Integer
   

   On Error GoTo ErrorHandler
   configflxPropertySelection2
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rsProperty As New ADODB.Recordset
   
   szSQL = "SELECT PropertyID, PropertyName,ClientID, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyID;"


    rsProperty.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
    i = 1
    While Not rsProperty.EOF
        flxPropertySelection2.AddItem ""
        flxPropertySelection2.TextMatrix(i, 0) = ""
        flxPropertySelection2.TextMatrix(i, 1) = rsProperty("PropertyID").Value
        flxPropertySelection2.TextMatrix(i, 2) = rsProperty("PropertyName").Value
        flxPropertySelection2.TextMatrix(i, 3) = rsProperty("ClientID").Value
        i = i + 1
        rsProperty.MoveNext
    Wend
    rsProperty.Close
    Set rsProperty = Nothing
    Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
End Sub
'Private Sub loadflxPropertySelection3() '***************************************Load  PROPERTY Grid ******************************************
'   Dim szSQL    As String
'   Dim TotalRow As Integer
'   Dim TotalCol As Integer
'   Dim i        As Integer
'   Dim j        As Integer
'
'
'   On Error GoTo ErrorHandler
'   configflxPropertySelection3
'   Dim adoConn As New adodb.Connection
'   adoConn.Open getConnectionString
'   Dim rsProperty As New adodb.Recordset
'
'   szSQL = "SELECT PropertyID, PropertyName,ClientID, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
'           "ORDER BY PropertyID;"
'
'
'    rsProperty.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic
'    i = 1
'    While Not rsProperty.EOF
'        flxPropertySelection2.AddItem ""
'        flxPropertySelection2.TextMatrix(i, 0) = ""
'        flxPropertySelection2.TextMatrix(i, 1) = rsProperty("PropertyID").Value
'        flxPropertySelection2.TextMatrix(i, 2) = rsProperty("PropertyName").Value
'        flxPropertySelection2.TextMatrix(i, 3) = rsProperty("ClientID").Value
'        i = i + 1
'        rsProperty.MoveNext
'    Wend
'    rsProperty.Close
'    Set rsProperty = Nothing
'    Exit Sub
'
'ErrorHandler:
'   MsgBox Err.description & "::" & Err.Number
'End Sub
Private Sub PrepareList4Property(cboProperty As Control)
   Dim szSQL    As String
   Dim TotalRow As Integer
   Dim TotalCol As Integer
   Dim i        As Integer
   Dim j        As Integer
   Dim Data()   As String

   On Error GoTo ErrorHandler
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rsProperty As New ADODB.Recordset
'*************************************** PROPERTY ******************************************
'Resolved by BOSL
'Issue No: 0000467
'Modified By: Asif. 04 Sep 2014
   
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyID;"
'''''''''''''''''''''''''''''''''

   rsProperty.Open szSQL, adoConn, adOpenKeyset, adLockOptimistic

   If rsProperty.EOF Then Exit Sub

   TotalRow = rsProperty.RecordCount
   TotalCol = rsProperty.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(rsProperty.Fields(j).Value), "", rsProperty.Fields(j).Value)
      Next j
      rsProperty.MoveNext
      If rsProperty.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number
End Sub
'Private Sub PrepareProperty4Agreement()
'   Dim szSQL    As String
'   Dim TotalRow As Integer
'   Dim TotalCol As Integer
'   Dim i        As Integer
'   Dim j        As Integer
'   Dim Data()   As String
'
'   On Error GoTo ErrorHandler
'   Dim adoconn As New ADODB.Connection
'   Dim rsProperty As New ADODB.Recordset
'   adoconn.ConnectionString = getConnectionString
'
'   szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
'           "FROM Property WHERE CLIENTID = '" & txtClientID.text & "' ORDER BY PropertyID;"
'
'
'   rsProperty.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
''   If rsProperty.EOF Then Exit Sub
''
''   TotalRow = rsProperty.RecordCount
''   TotalCol = rsProperty.Fields.count - 1
''
''   ReDim Data(TotalCol, TotalRow) As String
''
''   For i = 0 To TotalRow - 1
''      For j = 0 To TotalCol - 1
''         Data(j, i) = IIf(IsNull(rsProperty.Fields(j).Value), "", rsProperty.Fields(j).Value)
''      Next j
''      rsProperty.MoveNext
''      If rsProperty.EOF Then Exit For
''   Next i
''   cboProperty.Column() = Data()
'    i = 1
'    While Not rsProperty.EOF
'        flxPropertySelection1.AddItem ""
'        flxPropertySelection1.TextMatrix(i, 1) = ""
'        flxPropertySelection1.TextMatrix(i, 1) = ""
'        flxPropertySelection1.TextMatrix(i, 1) = ""
'        flxPropertySelection1.TextMatrix(i, 1) = ""
'        rsProperty.MoveNext
'    Wend
'
'   Exit Sub
'
'ErrorHandler:
'   MsgBox Err.description & "::" & Err.Number
'End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================

