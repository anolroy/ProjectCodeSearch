VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BB5807FE-DBD2-11D3-87C1-4C980CC10374}#1.0#0"; "MyHover.ocx"
Begin VB.Form frmPO 
   Caption         =   "Purchase Orders"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11070
   ScaleWidth      =   15000
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   1530
      ScaleHeight     =   2895
      ScaleWidth      =   4815
      TabIndex        =   82
      Top             =   8190
      Visible         =   0   'False
      Width           =   4845
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
         Index           =   0
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2175
         Index           =   0
         Left            =   15
         TabIndex        =   84
         Top             =   645
         Width           =   4765
         _ExtentX        =   8414
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
         Index           =   0
         Left            =   2115
         TabIndex        =   91
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   90
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   89
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch1 
         Height          =   195
         Left            =   1560
         TabIndex        =   88
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Left            =   3720
         TabIndex        =   87
         Top             =   135
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   86
         Top             =   375
         Width           =   1215
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2143;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1350
         TabIndex        =   85
         Top             =   375
         Width           =   1215
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2143;450"
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
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   4500
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   7605
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   71
      Top             =   8190
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
         TabIndex        =   72
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   73
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
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
         TabIndex        =   79
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   78
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
         Left            =   1875
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   74
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
         Left            =   0
         Top             =   120
         Width           =   5355
      End
   End
   Begin TabDlg.SSTab tabPurExp 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Purchase Orders"
      TabPicture(0)   =   "frmPO.frx":17D2A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape4(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label20(270)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20(19)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label20(16)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20(14)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20(15)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label20(17)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label20(13)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label20(12)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label20(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20(18)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPurchaseSplit(21)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblPurchaseSplit(22)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPurchaseSplit(19)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblPurchaseSplit(28)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblPurchaseSplit(24)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblPurchaseSplit(23)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblPurchaseSplit(25)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblPurchaseSplit(27)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblPurchaseSplit(26)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblPurchaseSplit(29)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label50(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Shape4(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label50(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label50(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdAccSel"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboAccount"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblPurchaseSplit(20)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtClientID"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "flxPurchaseSplit"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "flxPurchase"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkSelectAllDemands"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "fraEditDemand"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdClientSerc"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdTypeList"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtProperty"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Purchase Orders History"
      TabPicture(1)   =   "frmPO.frx":17D46
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSupplierHistory"
      Tab(1).Control(1)=   "cmdClientHistory"
      Tab(1).Control(2)=   "cmdRevHistory"
      Tab(1).Control(3)=   "cmdPrintListHistory"
      Tab(1).Control(4)=   "flxPurchHistory"
      Tab(1).Control(5)=   "flxPurchHistorySplit"
      Tab(1).Control(6)=   "txtSupplierHistory"
      Tab(1).Control(7)=   "txtClientHistory"
      Tab(1).Control(8)=   "Label50(2)"
      Tab(1).Control(9)=   "Shape4(2)"
      Tab(1).Control(10)=   "Label50(0)"
      Tab(1).Control(11)=   "Label50(1)"
      Tab(1).Control(12)=   "Label20(30)"
      Tab(1).Control(13)=   "Label20(31)"
      Tab(1).Control(14)=   "Label20(32)"
      Tab(1).Control(15)=   "Label20(29)"
      Tab(1).Control(16)=   "Label20(38)"
      Tab(1).Control(17)=   "Label20(34)"
      Tab(1).Control(18)=   "Label20(33)"
      Tab(1).Control(19)=   "Label20(35)"
      Tab(1).Control(20)=   "Label20(37)"
      Tab(1).Control(21)=   "Label20(36)"
      Tab(1).Control(22)=   "Label20(7)"
      Tab(1).Control(23)=   "Label20(5)"
      Tab(1).Control(24)=   "Label20(6)"
      Tab(1).Control(25)=   "Label20(8)"
      Tab(1).Control(26)=   "Label20(1)"
      Tab(1).Control(27)=   "Label20(4)"
      Tab(1).Control(28)=   "Label20(3)"
      Tab(1).Control(29)=   "Label20(2)"
      Tab(1).Control(30)=   "Shape4(1)"
      Tab(1).Control(31)=   "Label20(0)"
      Tab(1).Control(32)=   "cmbPropertyHistory"
      Tab(1).Control(33)=   "Label20(9)"
      Tab(1).ControlCount=   34
      Begin VB.CommandButton cmdSupplierHistory 
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
         Left            =   -63480
         TabIndex        =   66
         Top             =   675
         Width           =   255
      End
      Begin VB.CommandButton cmdClientHistory 
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
         Left            =   -71175
         TabIndex        =   63
         Top             =   630
         Width           =   255
      End
      Begin VB.TextBox txtProperty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   675
         Width           =   2550
      End
      Begin VB.CommandButton cmdTypeList 
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
         Left            =   8175
         TabIndex        =   80
         Top             =   690
         Width           =   255
      End
      Begin VB.CommandButton cmdClientSerc 
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
         Left            =   4590
         TabIndex        =   69
         Top             =   675
         Width           =   255
      End
      Begin VB.CommandButton cmdRevHistory 
         BackColor       =   &H00FEFEFE&
         Caption         =   "Reverse History"
         Height          =   400
         Left            =   -73320
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   7440
         Width           =   1440
      End
      Begin VB.CommandButton cmdPrintListHistory 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print List"
         Height          =   400
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   7440
         Width           =   1440
      End
      Begin VB.Frame fraEditDemand 
         Height          =   7500
         Left            =   40
         TabIndex        =   25
         Top             =   360
         Width           =   1215
         Begin VB.CommandButton cmdCreatePI 
            Caption         =   "Create Purchase Invoice"
            Height          =   735
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2400
            Width           =   1080
         End
         Begin VB.CommandButton cmdEmail 
            Caption         =   "&Email"
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4800
            Width           =   1080
         End
         Begin VB.CommandButton cmdPostDemands 
            Caption         =   "Post to Hist."
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   5805
            Width           =   1080
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   375
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1380
            Width           =   1080
         End
         Begin VB.CommandButton cmdPrintPI_List 
            Caption         =   "Print List"
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3780
            Width           =   1080
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Close"
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   6840
            Width           =   1080
         End
         Begin MyHoverButton.Button cmdNew 
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   26
            Top             =   360
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmPO.frx":17D62
            HoverPicture    =   "frmPO.frx":17D7E
            DisabledPicture =   "frmPO.frx":17D9A
            DownPicture     =   "frmPO.frx":17DB6
            MouseIcon       =   "frmPO.frx":17DD2
            Caption         =   "&Add New"
            HoverCaption    =   "Add New"
            DownCaption     =   "&Add New"
         End
      End
      Begin VB.CheckBox chkSelectAllDemands 
         Appearance      =   0  'Flat
         Caption         =   "&Select All"
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
         Height          =   215
         Left            =   1320
         TabIndex        =   1
         Top             =   1200
         Width           =   215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchase 
         Height          =   4515
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7964
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   8421631
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   12
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchaseSplit 
         Height          =   1635
         Left            =   1320
         TabIndex        =   13
         Top             =   6210
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2884
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   10
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistory 
         Height          =   3555
         Left            =   -74880
         TabIndex        =   39
         Top             =   1470
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6271
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   12
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPurchHistorySplit 
         Height          =   2115
         Left            =   -74880
         TabIndex        =   40
         Top             =   5295
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   3731
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   12
      End
      Begin MSForms.TextBox txtSupplierHistory 
         Height          =   255
         Left            =   -66090
         TabIndex        =   93
         Top             =   675
         Width           =   2880
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5080;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientHistory 
         Height          =   255
         Left            =   -73785
         TabIndex        =   92
         Top             =   630
         Width           =   2880
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5080;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientID 
         Height          =   255
         Left            =   1980
         TabIndex        =   70
         Top             =   675
         Width           =   2880
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "5080;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
         Height          =   195
         Index           =   2
         Left            =   -66855
         TabIndex        =   65
         Top             =   675
         Width           =   630
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   660
         Index           =   2
         Left            =   -74880
         Top             =   480
         Width           =   12400
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   64
         Top             =   675
         Width           =   465
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   1
         Left            =   -71010
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop/Unit"
         Height          =   195
         Index           =   30
         Left            =   -74280
         TabIndex        =   58
         Top             =   5085
         Width           =   690
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop/Unit Name"
         Height          =   195
         Index           =   31
         Left            =   -73320
         TabIndex        =   57
         Top             =   5085
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   32
         Left            =   -71460
         TabIndex        =   56
         Top             =   5085
         Width           =   285
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   29
         Left            =   -74760
         TabIndex        =   55
         Top             =   5085
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount £"
         Height          =   195
         Index           =   38
         Left            =   -63480
         TabIndex        =   54
         Top             =   5085
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No."
         Height          =   195
         Index           =   34
         Left            =   -69360
         TabIndex        =   53
         Top             =   5085
         Width           =   510
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   33
         Left            =   -70320
         TabIndex        =   52
         Top             =   5085
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   35
         Left            =   -68280
         TabIndex        =   51
         Top             =   5085
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Vat £"
         Height          =   195
         Index           =   37
         Left            =   -64440
         TabIndex        =   50
         Top             =   5085
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Net £"
         Height          =   195
         Index           =   36
         Left            =   -65520
         TabIndex        =   49
         Top             =   5085
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   -67080
         TabIndex        =   48
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   195
         Index           =   5
         Left            =   -70320
         TabIndex        =   47
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref."
         Height          =   195
         Index           =   6
         Left            =   -68160
         TabIndex        =   46
         Top             =   1245
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount £"
         Height          =   195
         Index           =   8
         Left            =   -64200
         TabIndex        =   45
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   44
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier A/C"
         Height          =   195
         Index           =   4
         Left            =   -71460
         TabIndex        =   43
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   3
         Left            =   -72420
         TabIndex        =   42
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   -74040
         TabIndex        =   41
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop/Unit"
         Height          =   195
         Index           =   20
         Left            =   1920
         TabIndex        =   38
         Top             =   6000
         Width           =   690
      End
      Begin MSForms.ComboBox cboAccount 
         Height          =   285
         Left            =   9135
         TabIndex        =   37
         Top             =   675
         Width           =   2745
         VariousPropertyBits=   1753237535
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4842;503"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411"
      End
      Begin MSForms.CommandButton cmdAccSel 
         Height          =   285
         Left            =   11880
         TabIndex        =   36
         Top             =   675
         Width           =   315
         Caption         =   "; ;"
         Size            =   "556;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   35
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   34
         Top             =   675
         Width           =   885
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   660
         Index           =   3
         Left            =   1320
         Top             =   480
         Width           =   11160
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Account:"
         Height          =   195
         Index           =   3
         Left            =   8520
         TabIndex        =   33
         Top             =   675
         Width           =   855
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Recoverable"
         Height          =   195
         Index           =   29
         Left            =   11520
         TabIndex        =   23
         Top             =   6000
         Width           =   885
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Net £"
         Height          =   195
         Index           =   26
         Left            =   8400
         TabIndex        =   22
         Top             =   6000
         Width           =   390
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Vat £"
         Height          =   195
         Index           =   27
         Left            =   9480
         TabIndex        =   21
         Top             =   6000
         Width           =   360
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   25
         Left            =   6480
         TabIndex        =   20
         Top             =   6000
         Width           =   840
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   23
         Left            =   4800
         TabIndex        =   19
         Top             =   6000
         Width           =   360
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Job No."
         Height          =   195
         Index           =   24
         Left            =   5640
         TabIndex        =   18
         Top             =   6000
         Width           =   510
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount £"
         Height          =   195
         Index           =   28
         Left            =   10440
         TabIndex        =   17
         Top             =   6000
         Width           =   675
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Index           =   19
         Left            =   1440
         TabIndex        =   16
         Top             =   6000
         Width           =   210
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   22
         Left            =   4140
         TabIndex        =   15
         Top             =   6000
         Width           =   285
      End
      Begin VB.Label lblPurchaseSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Prop/Unit Name"
         Height          =   195
         Index           =   21
         Left            =   2640
         TabIndex        =   14
         Top             =   6000
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoiced"
         Height          =   195
         Index           =   18
         Left            =   11160
         TabIndex        =   10
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   11
         Left            =   2040
         TabIndex        =   9
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   12
         Left            =   3060
         TabIndex        =   8
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C"
         Height          =   195
         Index           =   13
         Left            =   4020
         TabIndex        =   7
         Top             =   1215
         Width           =   270
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   6
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount £"
         Height          =   195
         Index           =   17
         Left            =   10080
         TabIndex        =   5
         Top             =   1215
         Width           =   675
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref."
         Height          =   195
         Index           =   15
         Left            =   6840
         TabIndex        =   4
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   14
         Left            =   5160
         TabIndex        =   3
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   16
         Left            =   7920
         TabIndex        =   2
         Top             =   1215
         Width           =   840
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
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
         Index           =   19
         Left            =   1320
         TabIndex        =   11
         Top             =   1230
         Width           =   11175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
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
         Index           =   270
         Left            =   1320
         TabIndex        =   24
         Top             =   6015
         Width           =   11175
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   660
         Index           =   6
         Left            =   1320
         Top             =   480
         Width           =   11160
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   660
         Index           =   1
         Left            =   -74880
         Top             =   480
         Width           =   12400
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   59
         Top             =   1260
         Width           =   12720
      End
      Begin MSForms.ComboBox cmbPropertyHistory 
         Height          =   285
         Left            =   -70305
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   3525
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6218;503"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058"
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   60
         Top             =   5100
         Width           =   12735
      End
   End
End
Attribute VB_Name = "frmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bFormLoaded  As Boolean
Private bTotalPayTyped  As Boolean
Private iSelected    As Integer
Private iPIEdit      As Integer
Private bEditMode    As Boolean        'Is the PI in Edit mode?
Private nTaxCode  As Double             'Tax code for Invoice
Private bHistoryLoaded As Boolean

Dim cGridSPTotal        As Currency
Dim bChangesMade  As Boolean            'This variable is uesed as flag that user has made any changes
Dim bPayAll             As Boolean
Dim bEmailResult  As Boolean
Private Type SendJobsByEmail
        szSuppID      As String
        szSuppEmail   As String
        szClient      As String
        colAtt        As Collection
        lURN          As Long
        szSuppName    As String
End Type
Private uSupplier(0)        As SendJobsByEmail
Dim sTextBox As String



Private Sub cboAccount_Change()
    If Not bFormLoaded Then Exit Sub
    SortTheGrid flxPurchase, txtClientID, txtProperty, cboAccount
    flxPurchaseSplit.Clear
End Sub

Private Sub chkSelectAllDemands_Click()
   Dim iRow As Integer, i As Integer

   If chkSelectAllDemands.Value Then
      For i = 1 To flxPurchase.Rows - 1
         flxPurchase.TextMatrix(i, 1) = ""
      Next i
      For i = 1 To flxPurchase.Rows - 2
         flxPurchase.TextMatrix(i, 1) = "X"
      Next i

      flxPurchase.row = flxPurchase.Rows - 1

      flxPurchase_Click
   Else
      For i = 1 To flxPurchase.Rows - 1
         flxPurchase.TextMatrix(i, 1) = ""
      Next i

      ConfigFlxPurchaseSplit
   End If
End Sub

Private Sub cmdAccSel_Click()
   sTextBox = "A/C"
   LoadSupplierAccount
   txtSearch1.Visible = True
   txtSearch2.Visible = True
   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4500
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = 5800 'txtAc(0).Left + 100
   fraList.Top = 350 'txtAc(0).Top
   fraList.Visible = True
   fraList.ZOrder 0
   
   txtSearch1.SetFocus
   'fraLay.Enabled = False
   'fraControls.Enabled = False
   tabPurExp.Enabled = False
   'fraCmds.Enabled = False
End Sub
Private Sub LoadSupplierAccount()
   Dim adoConn As New ADODB.Connection
   Dim rstRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer

'ConfigFlxSupplier - Configuring flxSupplier grid
   With flxSupplier(0)
      .Cols = 6
      .ColWidth(0) = 1000
      .ColWidth(1) = 2200
      .ColAlignment(1) = vbLeftJustify
      .ColWidth(2) = 0
      .ColWidth(3) = 0
      .ColWidth(4) = 1000
      .ColWidth(5) = 0

      '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
      lblSearch0(0).Width = 700
      lblSearch0(0).Left = 60
      lblSearch1.Width = 2600
      lblSearch1.Left = lblSearch0(0).Left + .ColWidth(0)
      lblSearch2.Width = 750
      lblSearch2.Left = 3220
      lblSearch2.Visible = True

      txtSearch1.Width = 900
      txtSearch1.Left = 70

      txtSearch2.Width = 2200
      txtSearch2.Left = txtSearch1.Left + .ColWidth(0)

      ' Error Handler
      On Error GoTo ErrorHandler

      'Set the RDO Connections to the dataset
      adoConn.Open getConnectionString

      '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(0).Caption = "A/C ID"
      lblSearch1.Caption = "Name"
      lblSearch2.Caption = "A/C Bal"

      
         szSQL = "SELECT SupplierID, SupplierName, NominalCode, VATCode, PaymentTerms, VAT_RATE " & _
                 "FROM Supplier LEFT JOIN tlbVatCode " & _
                     "ON Supplier.VATCode = tlbVatCode.VAT_CODE " & _
                 "WHERE Supplier.TYPE = 'SUPPLIER' " & _
                 "ORDER BY SupplierName;"
'Debug.Print szSQL
            .Clear
            .Rows = 2
            rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            iRow = 1
            
             flxSupplier(0).TextMatrix(1, 0) = "ALL"
           flxSupplier(0).TextMatrix(1, 1) = "All Suppliers"
           flxSupplier(0).AddItem ""
           iRow = 2
           
            While Not rstRst.EOF
               .TextMatrix(iRow, 0) = rstRst!SupplierID
               .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!SupplierName), "", rstRst!SupplierName)
               .TextMatrix(iRow, 2) = IIf(IsNull(rstRst!nominalCode), "", rstRst!nominalCode)
               .TextMatrix(iRow, 3) = IIf(IsNull(rstRst!VatCode) Or rstRst!VatCode = "", "", rstRst!VatCode & "##" & rstRst!VAT_RATE)
'               .TextMatrix(iRow, 3) = IIf(IsNull(rstRst!VatCode) Or rstRst!VatCode = "", "T9##0", rstRst!VatCode & "##" & rstRst!VAT_RATE)
               .TextMatrix(iRow, 5) = IIf(IsNull(rstRst!PaymentTerms), "", rstRst!PaymentTerms)
               rstRst.MoveNext
               If Not rstRst.EOF Then .AddItem ""
               iRow = iRow + 1
            Wend
            
     
            End With

   rstRst.Close
   
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   Exit Sub
   
ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"
   
   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdClientHistory_Click()
    picClient.Visible = True
    picClient.Left = 1980
    picClient.Top = 675
    sTextBox = "3"
    LoadflxClient
    tabPurExp.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClientSerc_Click()
    
    picClient.Left = 1980
    picClient.Top = 675
    sTextBox = "1"
    LoadflxClient
    picClient.Visible = True
    tabPurExp.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 3600
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           flxClient.TextMatrix(1, 0) = "ALL"
           flxClient.TextMatrix(1, 1) = "All Client"
           flxClient.TextMatrix(1, 2) = ""
           flxClient.RowHeight(1) = 280
           flxClient.AddItem ""
           rRow = 2
           
            'rRow = 1
            While Not rstRec.EOF
              ' flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
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
Private Sub cmdGridUnitLookup_Click(Index As Integer)
        fraList.Visible = False
        If sTextBox = "PROPERTY" Then
        End If
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    tabPurExp.Enabled = True
    cmdClientSerc.SetFocus
End Sub

Private Sub cmdSupplierHistory_Click()
    sTextBox = "4"
   LoadSupplierAccount
   txtSearch1.Visible = True
   txtSearch2.Visible = True
   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4500
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = 5800 'txtAc(0).Left + 100
   fraList.Top = 350 'txtAc(0).Top
   fraList.Visible = True
   fraList.ZOrder 0
   
   txtSearch1.SetFocus
   'fraLay.Enabled = False
   'fraControls.Enabled = False
   tabPurExp.Enabled = False
   'fraCmds.Enabled = False
End Sub

Private Sub cmdTypeList_Click()
       sTextBox = "2"
       LoadPropertyList
       txtSearch1.Visible = True
       txtSearch2.Visible = True
    
       txtSearch1.text = ""
       txtSearch2.text = ""
    
       fraList.Width = 4815
       cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
       Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
       flxSupplier(0).Width = fraList.Width - 50
       fraList.Left = txtProperty.Left + txtProperty.Width - fraList.Width '+ fraLay(0).Left
       fraList.Top = txtProperty.Top '+ fraLay(0).Top '+ tabPurExp.Top '+ 380
       fraList.Visible = True
       fraList.ZOrder 0
       
       txtSearch1.SetFocus
End Sub
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxSupplier(0).RowHeight(0) = 0
   flxSupplier(0).Cols = 2
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700

   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   flxSupplier(0).ColAlignment(0) = vbLeftJustify
   flxSupplier(0).ColAlignment(1) = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1.Width = 1400
   txtSearch1.Left = 40
   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   lblSearch0(0).Caption = "Property ID"
   lblSearch1.Caption = "Property Name"
   lblSearch2.Visible = False
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   If txtClientID.text = "ALL" Then
            szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "ORDER BY PropertyID;"
    Else
            szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "WHERE ClientID = '" & txtClientID.text & "' " & _
            "ORDER BY PropertyID;"
    End If

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            flxSupplier(0).TextMatrix(1, 0) = "ALL"
           flxSupplier(0).TextMatrix(1, 1) = "All Properties"
           flxSupplier(0).AddItem ""
           rRow = 2
            
   While Not rstRec.EOF
      flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
      flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
      flxSupplier(0).RowHeight(rRow) = 280
      rstRec.MoveNext
      If Not rstRec.EOF Then flxSupplier(0).AddItem ""
      rRow = rRow + 1
   Wend

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub flxClient_Click()
        tabPurExp.Enabled = True
        If sTextBox = "1" Then
              txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
        ElseIf sTextBox = "2" Then
              txtProperty.text = flxClient.TextMatrix(flxClient.row, 0)
        ElseIf sTextBox = "3" Then
              txtClientHistory.text = flxClient.TextMatrix(flxClient.row, 0)
        ElseIf sTextBox = "4" Then
              txtClientHistory.text = flxClient.TextMatrix(flxClient.row, 0)
        End If
        picClient.Visible = False
        txtProperty.text = "ALL"
        cmdClientSerc.SetFocus
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       tabPurExp.Enabled = True
        If sTextBox = "1" Then
              txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
        ElseIf sTextBox = "2" Then
              txtProperty.text = flxClient.TextMatrix(flxClient.row, 0)
        ElseIf sTextBox = "3" Then
              txtClientHistory.text = flxClient.TextMatrix(flxClient.row, 0)
         ElseIf sTextBox = "4" Then
              txtClientHistory.text = flxClient.TextMatrix(flxClient.row, 0)
        End If
        picClient.Visible = False
        txtProperty.text = "ALL"
        cmdClientSerc.SetFocus
    End If
End Sub

Private Sub flxSupplier_Click(Index As Integer)
    tabPurExp.Enabled = True
     If sTextBox = "A/C" Then
            cboAccount.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
            cmdAccSel.SetFocus
      ElseIf sTextBox = "4" Then
            txtSupplierHistory.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
            cmdSupplierHistory.SetFocus
      Else
            txtProperty.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
            cmdTypeList.SetFocus
      End If
      fraList.Visible = False
      
End Sub

Private Sub flxSupplier_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
                tabPurExp.Enabled = True
                If sTextBox = "A/C" Then
                    cboAccount.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
                    cmdAccSel.SetFocus
                Else
                    txtProperty.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
                    cmdTypeList.SetFocus
                End If
                fraList.Visible = False
      End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
End Sub

Private Sub txtClientID_Change()
     If Not bFormLoaded Then Exit Sub
     SortTheGrid flxPurchase, txtClientID, txtProperty, cboAccount
     flxPurchaseSplit.Clear
End Sub
Private Sub SortTheGrid(flxGrid As MSHFlexGrid, cmbClientCombo As Control, cmbPropCombo As Control, cmbSuppCombo As Control)
   Dim sFlag As Single, iRow As Integer
   Dim szSQL As String

   sFlag = 0
   For iRow = 1 To flxGrid.Rows - 1
      If cmbClientCombo.text = "ALL" Then
         sFlag = 100
      Else
         If flxGrid.TextMatrix(iRow, 14) = cmbClientCombo.text Then sFlag = 100
      End If

      If Len(cmbPropCombo.text) > 0 Then
         If cmbPropCombo.text = "ALL" Then
            sFlag = sFlag + 10
         Else
            If flxGrid.TextMatrix(iRow, 11) = cmbPropCombo.text Then sFlag = sFlag + 10
         End If
      Else
         sFlag = sFlag + 10
      End If

     ' If Len(txtSupplierSearc.text) > 0 Then
         If cmbSuppCombo.text = "ALL" Then
            sFlag = sFlag + 1
         Else
            If flxGrid.TextMatrix(iRow, 5) = cmbSuppCombo.text Then sFlag = sFlag + 1
         End If
'      Else
'         sFlag = sFlag + 1
'      End If

      If sFlag = 111 Then
         flxGrid.RowHeight(iRow) = 240
      Else
         flxGrid.RowHeight(iRow) = 0
      End If
   Next iRow
End Sub

Private Sub txtProperty_Change()
    If Not bFormLoaded Then Exit Sub
    SortTheGrid flxPurchase, txtClientID, txtProperty, cboAccount
    flxPurchaseSplit.Clear
End Sub

Private Sub txtSearch1_Change()
         'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearch1.text) > 0 Then
        txtSearch2.text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
      flxSupplier(0).RowHeight(i) = 240
      
      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtSearch1.text))) <> UCase(txtSearch1.text) Then
            flxSupplier(0).RowHeight(i) = 0
      End If
      If flxSupplier(0).RowHeight(i) = 240 Then
            flxSupplier(0).row = i
      End If
   Next i
End Sub

Private Sub txtSearch1_KeyPress(KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 27 Then
          
          fraList.Visible = False
          tabPurExp.Enabled = True
         
         If sTextBox = "A/C" Then
              txtProperty.SetFocus
         Else
              cboAccount.SetFocus
         End If
    End If
End Sub

Private Sub txtSearch2_Change()
    'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearch2.text) > 0 Then
        txtSearch1.text = ""
   End If

   For i = flxSupplier(0).Rows - 1 To 1 Step -1
      flxSupplier(0).RowHeight(i) = 240
      
      If UCase(Left(flxSupplier(0).TextMatrix(i, 1), Len(txtSearch2.text))) <> UCase(txtSearch2.text) Then
            flxSupplier(0).RowHeight(i) = 0
      End If
      If flxSupplier(0).RowHeight(i) = 240 Then
            flxSupplier(0).row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientID_Change()
    'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
      flxClient.RowHeight(i) = 240
      If InStr(1, UCase(flxClient.TextMatrix(i, 0)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
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
    If KeyAscii = 13 Then
         txtSearchClientName.SetFocus
    End If
    If KeyAscii = 27 Then
          flxClient.Clear
          flxClient.Cols = 2
          flxClient.Rows = 2
          picClient.Visible = False
          tabPurExp.Enabled = True
         
       
           cmdClientSerc.SetFocus
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
      If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
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

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
         flxClient.SetFocus
    End If
End Sub
Private Sub cmdCreatePI_Click()
   Dim X As Byte
    
   With flxPurchase
      If .row <= 0 Then Exit Sub
      'added by anol 02 Dec 2015
      If iPIEdit = 0 Then
            ShowMsgInTaskBar "Please select a Purchase Order form the list", "Y", "N"
            Exit Sub
      End If


      Load frmPO2PI

      frmPO2PI.txtTransType.text = .TextMatrix(iPIEdit, 3)
      frmPO2PI.txtAc(0).text = .TextMatrix(iPIEdit, 5)
'      frmPO2PI.cmdACList(0).Enabled = IIf(.TextMatrix(iPIEdit, 9) <> .TextMatrix(iPIEdit, 12), False, True)
      frmPO2PI.txtDate.text = .TextMatrix(iPIEdit, 4)
      frmPO2PI.txtDueDate.text = Format(.TextMatrix(iPIEdit, 13), "dd/mm/yyyy")
      frmPO2PI.txtSupplierName.text = .TextMatrix(iPIEdit, 6)
      frmPO2PI.txtInv(0).text = .TextMatrix(iPIEdit, 7)
      frmPO2PI.cboClientPI.Value = .TextMatrix(iPIEdit, 14)
      frmPO2PI.szPropertyID = .TextMatrix(iPIEdit, 11)
      'issue 469
      'Below line is commented by anol 06 Jan 2015
      
      'frmPO2PI.lblPostingDate.ToolTipText = .TextMatrix(iPIEdit, 16)
      frmPO2PI.PO_Ref = .TextMatrix(iPIEdit, 0)
      frmPO2PI.txtProperty.text = .TextMatrix(iPIEdit, 18)

      LoadSplit4PI frmPO2PI.flxPI, frmPO2PI

      frmPO2PI.cmdSavePI.Enabled = True
   End With
   frmPO2PI.Show
   Me.Enabled = False
End Sub

Private Sub cmdEdit_Click(Index As Integer)
   If IsLoadedAndVisible("frmPO_Amend") Then
      ShowMsgInTaskBar "Purchase Order form is already open", "Y", "N"
      Exit Sub
   End If

'   If flxMaintenance.TextMatrix(flxMaintenance.row, 0) = "JOB" Then
'   If flxPurchase.TextMatrix(flxPurchase.row, 16) <> "S" Then
'      ShowMsgInTaskBar "A purchase order cannot be created, as  this job is assigned internally", "Y", "N"
'      Exit Sub
'   End If

'   If MsgBox("Do you wish to create a purchase order from this job?", vbQuestion + vbYesNo, "Purchase Order") = vbNo Then Exit Sub

   Load frmPO_Amend

   With frmPO_Amend
      .Caption = "Edit Purchase Order - " & flxPurchase.TextMatrix(flxPurchase.row, 2)
      .szPO = flxPurchase.TextMatrix(flxPurchase.row, 0)
'      .LoadDate
      .bEditMode = True
      .szCallerForm = "P"
      .sPI = flxPurchase.TextMatrix(iPIEdit, 0)

      .Show
      .txtInv(0).SetFocus
   End With
   
   If iPIEdit = 0 Then Exit Sub
'
'   Dim X As Byte
'
'   If Not IsPossible2Edit Then
'      ShowMsgInTaskBar "The transaction is fully/partially paid.", "Y", "N"
'      Exit Sub
'   End If

   cmdEdit(1).Enabled = False                'At the saving time system will know if PO is in EDIT mode

   With flxPurchase
      frmPO_Amend.txtTransType.text = .TextMatrix(iPIEdit, 3)
      frmPO_Amend.txtAc(0).text = .TextMatrix(iPIEdit, 5)
      frmPO_Amend.cmdACList(0).Enabled = IIf(.TextMatrix(iPIEdit, 9) <> .TextMatrix(iPIEdit, 12), False, True)
      frmPO_Amend.txtDate.text = .TextMatrix(iPIEdit, 4)
      frmPO_Amend.txtDueDate.text = Format(.TextMatrix(iPIEdit, 13), "dd/mm/yyyy")
      frmPO_Amend.txtSupplierName.text = .TextMatrix(iPIEdit, 6)
      frmPO_Amend.txtInv(0).text = .TextMatrix(iPIEdit, 7)
      frmPO_Amend.txtClientID.text = .TextMatrix(iPIEdit, 14)
      frmPO_Amend.szPropertyID = .TextMatrix(iPIEdit, 11)
      frmPO_Amend.txtProperty.text = .TextMatrix(iPIEdit, 18)
'      frmPO_Amend.lblPostingDate.ToolTipText = .TextMatrix(iPIEdit, 16)

      LoadSplit4Edit frmPO_Amend.flxPI, frmPO_Amend

'      fraLay(0).Left = 120
'      fraLay(0).Top = 360
      frmPO_Amend.cmdNew(0).Enabled = False
'      fraLay(1).Caption = "Transaction ID: " & .TextMatrix(iPIEdit, 2)
      frmPO_Amend.cmdUpdate(1).Enabled = True

'      cmdSavePI.Enabled = True
   End With
   
   Me.Enabled = False
End Sub

Private Sub LoadInvoicedSplits(GridFlxPI As MSHFlexGrid, adoConn As ADODB.Connection)
   Dim adoInvSp As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer
'Modified by anol 14 Dec 2015
 szSQL = "select POPICross from tblPurInvSRec where ParentID= '" & flxPurchase.TextMatrix(iPIEdit, 0) & "'"
 adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT DISTINCT tblPurInvSRec.*, Fund.FundName, Fund.FundCode " & _
           "FROM tblPurInvSRec, Fund " & _
           "WHERE tblPurInvSRec.MY_ID ='" & adoInvSp.Fields(0).Value & "' AND " & _
                 "tblPurInvSRec.DEPT_ID = Fund.FundID AND NOt Invoiced " & _
           "ORDER BY TRAN_ID;"
'Debug.Print szSQL
adoInvSp.Close
   adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With GridFlxPI
      .Rows = 2
      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 1) = frmPO2PI.txtAc(0).text
         .TextMatrix(iRow, 2) = frmPO2PI.txtDate.text
         .TextMatrix(iRow, 3) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 4) = IIf(frmPO2PI.txtTransType.text = "Invoice", "Invoice", "Credit")
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 6) = frmPO2PI.txtInv(0).text
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 9) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 12) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 13) = adoInvSp.Fields.Item("TAX_CODE").Value
         .TextMatrix(iRow, 14) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 15) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 20) = IIf(IsNull(adoInvSp.Fields.Item("ScheduleID").Value), "", adoInvSp.Fields.Item("ScheduleID").Value)
         .TextMatrix(iRow, 21) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 22) = adoInvSp.Fields.Item("RecoverablePt").Value
         .TextMatrix(iRow, 23) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 24) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 25) = adoInvSp.Fields.Item("FundName").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      .row = 0
   End With

   adoInvSp.Close
   Set adoInvSp = Nothing
End Sub

Private Sub LoadSplit4Edit(GridFlxPI As MSHFlexGrid, frmForm As Form)
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer

   adoConn.Open getConnectionString

   szSQL = "SELECT DISTINCT tblPurInvSRec.*, Fund.FundName, Fund.FundCode " & _
           "FROM tblPurInvSRec, Fund " & _
           "WHERE tblPurInvSRec.ParentID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "' AND " & _
                 "tblPurInvSRec.DEPT_ID = Fund.FundID AND DESCRIPTION <> 'DELETED' " & _
           "ORDER BY TRAN_ID;"
'Debug.Print szSQL
   adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With GridFlxPI
      .Rows = 2
      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 1) = frmForm.txtAc(0).text
         .TextMatrix(iRow, 2) = frmForm.txtDate.text
         .TextMatrix(iRow, 3) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 4) = IIf(frmForm.txtTransType.text = "Invoice", "Invoice", "Credit")
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 6) = frmForm.txtInv(0).text
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 9) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 12) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 13) = adoInvSp.Fields.Item("TAX_CODE").Value
         .TextMatrix(iRow, 14) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 15) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 20) = IIf(IsNull(adoInvSp.Fields.Item("ScheduleID").Value), "", adoInvSp.Fields.Item("ScheduleID").Value)
         .TextMatrix(iRow, 21) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 22) = adoInvSp.Fields.Item("RecoverablePt").Value
         .TextMatrix(iRow, 23) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 24) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 25) = adoInvSp.Fields.Item("FundName").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      .row = 0
'
'      LoadInvoicedSplits frmForm.flxInvoiced
   End With

   adoInvSp.Close
   Set adoInvSp = Nothing

   adoConn.Close
   Set adoConn = Nothing

   UpdateTotalPICN frmForm
End Sub

Private Sub LoadSplit4PI(GridFlxPI As MSHFlexGrid, frmForm As Form)
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String, iRow As Integer
   Dim amount As Double
   Dim vat As Double
   Dim Total As Double
   adoConn.Open getConnectionString

   szSQL = "SELECT DISTINCT tblPurInvSRec.*, Fund.FundName, Fund.FundCode " & _
           "FROM tblPurInvSRec, Fund " & _
           "WHERE tblPurInvSRec.ParentID = '" & flxPurchase.TextMatrix(iPIEdit, 0) & "' AND " & _
                 "tblPurInvSRec.DEPT_ID = Fund.FundID AND Not Invoiced AND DESCRIPTION <> 'DELETED' " & _
           "ORDER BY TRAN_ID;"
''Modified by anol 14 Dec 2015
' szSQL = "select POPICross from tblPurInvSRec where ParentID= '" & flxPurchase.TextMatrix(iPIEdit, 0) & "'"
' adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'            szSQL = "SELECT DISTINCT tblPurInvSRec.*, Fund.FundName, Fund.FundCode " & _
'           "FROM tblPurInvSRec, Fund " & _
'           "WHERE tblPurInvSRec.MY_ID ='" & adoInvSp.Fields(0).Value & "' AND " & _
'                 "tblPurInvSRec.DEPT_ID = Fund.FundID AND Not Invoiced AND DESCRIPTION <> 'DELETED' " & _
'           "ORDER BY TRAN_ID;"
'Debug.Print szSQL
'   adoInvSp.Close
   adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   With GridFlxPI
      .Rows = 2
      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 1) = frmPO2PI.txtAc(0).text
         .TextMatrix(iRow, 2) = frmPO2PI.txtDate.text
         .TextMatrix(iRow, 3) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 4) = IIf(frmPO2PI.txtTransType.text = "Invoice", "Invoice", "Credit")
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 6) = frmPO2PI.txtInv(0).text
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 8) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 9) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 12) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         'added by anol 08 Dec 2015
         amount = amount + adoInvSp.Fields.Item("NET_AMOUNT").Value
         .TextMatrix(iRow, 13) = adoInvSp.Fields.Item("TAX_CODE").Value
         .TextMatrix(iRow, 14) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         vat = vat + adoInvSp.Fields.Item("VAT").Value
         .TextMatrix(iRow, 15) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         Total = Total + adoInvSp.Fields.Item("TOTAL_AMOUNT").Value
         .TextMatrix(iRow, 20) = IIf(IsNull(adoInvSp.Fields.Item("ScheduleID").Value), "", adoInvSp.Fields.Item("ScheduleID").Value)
         .TextMatrix(iRow, 21) = IIf(IsNull(adoInvSp.Fields.Item("UNIT_ID").Value), "", adoInvSp.Fields.Item("UNIT_ID").Value)
         .TextMatrix(iRow, 22) = adoInvSp.Fields.Item("RecoverablePt").Value
         .TextMatrix(iRow, 23) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 24) = adoInvSp.Fields.Item("FundCode").Value
         .TextMatrix(iRow, 25) = adoInvSp.Fields.Item("FundName").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      .row = 0
      
      LoadInvoicedSplits frmForm.flxInvoiced, adoConn
   End With
    'added by anol 08 Dec 2015
    If frmForm.Name = "frmPO2PI" Then
           frmPO2PI.txtPICNNet.text = amount
           frmPO2PI.txtPICNVat.text = vat
           frmPO2PI.txtPICNTotal.text = Total
    End If
   adoInvSp.Close
   Set adoInvSp = Nothing
   adoConn.Close
   Set adoConn = Nothing

   UpdateTotalPICN frmForm
End Sub

Private Sub UpdateTotalPICN(frmForm As Form)
   Dim i As Integer

   With frmForm
      .txtPICNNet.text = "0.00"
      .txtPICNVat.text = "0.00"
      .txtPICNTotal.text = "0.00"
'
      For i = 1 To .flxPI.Rows - 1
         .txtPICNNet.text = Val(.txtPICNNet.text) + Val(.flxPI.TextMatrix(i, 12))
         .txtPICNVat.text = Val(.txtPICNVat.text) + Val(.flxPI.TextMatrix(i, 14))
         .txtPICNTotal.text = Val(.txtPICNTotal.text) + Val(.flxPI.TextMatrix(i, 15))
      Next i

      .txtPICNNet.text = Format(.txtPICNNet.text, "0.00")
      .txtPICNVat.text = Format(.txtPICNVat.text, "0.00")
      .txtPICNTotal.text = Format(.txtPICNTotal.text, "0.00")
   End With
End Sub

Private Sub SaveAttachment(szFile As String, szSupplier As String, szSupplierEmail As String)
   Dim i As Integer

   i = 0
   Set uSupplier(i).colAtt = New Collection
   uSupplier(i).colAtt.Add szFile
   uSupplier(i).szSuppEmail = szSupplierEmail
End Sub

Private Sub cmdEmail_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PO.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Report.ParameterFields(1).AddCurrentValue flxPurchase.TextMatrix(flxPurchase.row, 0)
   Report.ParameterFields(2).AddCurrentValue "PURCHASE ORDER"
   
   Report.ExportOptions.DiskFileName = DB_PATH & "\AllStuff\Temp\PO_" & flxPurchase.TextMatrix(flxPurchase.row, 0) & ".pdf"
   Report.ExportOptions.DestinationType = crEDTDiskFile
   Report.ExportOptions.FormatType = crEFTPortableDocFormat
   Report.ExportOptions.PDFExportAllPages = True
   Report.Export False

   Set Report = Nothing
   
   SaveAttachment DB_PATH & "\AllStuff\Temp\PO_" & flxPurchase.TextMatrix(flxPurchase.row, 0) & ".pdf", flxPurchase.TextMatrix(flxPurchase.row, 5), flxPurchase.TextMatrix(flxPurchase.row, 17)
   
   EmailDelay 20
   Set Report = Nothing
   
'  Sending Email with demand invoice as attachments
   bEmailResult = SendJobByE_Mail("A purchase order has been sent to you", _
                                     "Please find the purchase order details in the attachment")
   ShowMsgInTaskBar "The purchase order has been sent by email", "Y", "P"
End Sub

Private Function SendJobByE_Mail(szSub As String, szBody As String) As Boolean
   Dim i As Integer

   i = 0
      SendJobByE_Mail = SendEmail(szFromEmail, Trim(uSupplier(i).szSuppEmail), _
                                     szSub, _
                                     szBody, , , _
                                     uSupplier(i).colAtt, uSupplier(i).szSuppID, "SI")
End Function

Private Sub cmdNew_Click(Index As Integer)
   If IsLoadedAndVisible("frmPO_Amend") Then
      ShowMsgInTaskBar "Purchase Order form is already open", "Y", "N"
      Exit Sub
   End If

   Load frmPO_Amend
   frmPO_Amend.szCallerForm = "P"
   frmPO_Amend.Show
   frmPO_Amend.cmdACList(0).SetFocus
   frmPO_Amend.bEditMode = False
   Me.Enabled = False
End Sub

Private Sub cmdPostDemands_Click()
   Dim szTemp As String

   'frmPopUpMenu.Top = frmMMain.fraCmdButton.Height + Me.Top + frmPO.Top + fraEditDemand.Top + cmdPostDemands.Top + 1140
   'frmPopUpMenu.Left = frmMMain.tvwLandLord.Width + Me.Left + fraEditDemand.Left + frmPO.Left + cmdPostDemands.Left + 80
   frmPopUpMenu.CallingFrom "PostPO"

   szTemp = SelectedPurInvID()

   If szTemp <> "" Then
      frmPopUpMenu.optSelPI.Value = True
   Else
      frmPopUpMenu.optSelPI.Value = False
      frmPopUpMenu.optSelPI.Enabled = False
   End If

   frmPopUpMenu.Show 1
End Sub

Private Function SelectedPurInvID() As String
   Dim i As Integer

   SelectedPurInvID = ""
   For i = 1 To flxPurchase.Rows - 1
      If flxPurchase.TextMatrix(i, 1) = "X" Then
         SelectedPurInvID = SelectedPurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         SelectedPurInvID = SelectedPurInvID & ","
      End If
   Next i
   If SelectedPurInvID <> "" Then
      SelectedPurInvID = Left(SelectedPurInvID, Len(SelectedPurInvID) - 1)
   Else
      SelectedPurInvID = ""
   End If
End Function

Private Sub cmdPrintListHistory_Click()
   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tblPurInv SET Prn = 'N' WHERE TransactionType = 25;"

   szMY_ID = ""
   For i = 1 To flxPurchHistory.Rows - 2
      If flxPurchHistory.RowHeight(i) > 0 Then _
         szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "', "
   Next i
   szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "'"

   szMY_ID = "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ") AND TransactionType = 25;"

   adoConn.Execute szMY_ID
   adoConn.Close
   Set adoConn = Nothing

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PO_List.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue "Y"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdPrintPI_List_Click()
   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   adoConn.Open getConnectionString

   adoConn.Execute "UPDATE tblPurInv SET Prn = 'N' WHERE TransactionType = 25;"

   szMY_ID = ""
   For i = 1 To flxPurchase.Rows - 2
      If flxPurchase.RowHeight(i) > 0 Then _
         szMY_ID = szMY_ID + "'" + flxPurchase.TextMatrix(i, 0) + "', "
   Next i
   szMY_ID = szMY_ID + "'" + flxPurchase.TextMatrix(i, 0) + "'"

   szMY_ID = "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ") AND TransactionType = 25;"

   adoConn.Execute szMY_ID
   adoConn.Close
   Set adoConn = Nothing

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PO_List.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue "N"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Private Sub cmdRevHistory_Click()
   If MsgBox("Do you wish to reverse the purchase order from history?", vbQuestion + vbYesNo, "Purchase Order") = vbNo Then Exit Sub

   Dim szPI_ID As String
   Dim szSQL As String
   Dim iPosted As Integer               'Finally posted
   Dim iIP     As Integer               'To be posted
   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   szPI_ID = flxPurchHistory.TextMatrix(flxPurchHistory.row, 0)

   If Len(szPI_ID) = 0 Then
      ShowMsgInTaskBar "No purchase order to post to history.", "Y", "N"

      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   szSQL = "UPDATE tblPurInv " & _
           "SET History = FALSE " & _
           "WHERE MY_ID ='" & szPI_ID & "';"

   adoConn.Execute szSQL

   LoadFlxPurchase adoConn
   LoadFlxPurchHistory adoConn

   MousePointer = vbDefault

   adoConn.Close
   Set adoConn = Nothing
   ShowMsgInTaskBar "System has reveresed the purchase orders from history.", "Y", "P"
End Sub

Private Sub flxPurchase_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection

   If flxPurchase.TextMatrix(flxPurchase.row, 0) = "" Then Exit Sub
   If flxPurchase.RowHeight(flxPurchase.row) = 0 Then
      iPIEdit = 0
      Exit Sub
   End If

   adoConn.Open getConnectionString

'   HighLightRowFlxGrid flxPurchase, flxPurchase.row
   SelectFlxGridRow 1, flxPurchase, flxPurchase.row

   iPIEdit = flxPurchase.row

   ConfigFlxPurchaseSplit

   With flxPurchaseSplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyID, U.UnitNumber, " & _
                  "P.PropertyName, U.UnitName " & _
              "FROM (tblPurInvSRec AS S " & _
                  "LEFT JOIN  Property AS P ON S.TRANS = P.PropertyID) " & _
                  "LEFT JOIN Units AS U ON S.UNIT_ID = U.UnitNumber " & _
              "WHERE S.ParentID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "' AND " & _
                  "DESCRIPTION <> 'DELETED' " & _
              "ORDER BY TRAN_ID;"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = IIf(IsNull(adoInvSp.Fields.Item("UnitNumber").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyID").Value), "", _
                                 adoInvSp.Fields.Item("PropertyID").Value), adoInvSp.Fields.Item("UnitNumber").Value)
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("UnitName").Value), _
                                 IIf(IsNull(adoInvSp.Fields.Item("PropertyName").Value), "", _
                                 adoInvSp.Fields.Item("PropertyName").Value), adoInvSp.Fields.Item("UnitName").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = IIf(IsNull(adoInvSp.Fields.Item("DEPT_ID").Value), "", adoInvSp.Fields.Item("DEPT_ID").Value)
         .TextMatrix(iRow, 6) = IIf(IsNull(adoInvSp.Fields.Item("JOB_ID").Value), "", adoInvSp.Fields.Item("JOB_ID").Value)
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = Format(adoInvSp.Fields.Item("NET_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 9) = Format(adoInvSp.Fields.Item("VAT").Value, "0.00")
         .TextMatrix(iRow, 10) = Format(adoInvSp.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iRow, 11) = adoInvSp.Fields.Item("RecoverablePt").Value & "%"

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoConn.Close
   Set adoInvSp = Nothing
   Set adoConn = Nothing
End Sub

Private Sub flxPurchHistory_Click()
   Dim szSQL As String, iRow As Integer
   Dim adoInvSp As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection

   If flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) = "" Then Exit Sub
   If flxPurchHistory.RowHeight(flxPurchHistory.row) = 0 Then Exit Sub

   adoConn.Open getConnectionString

   HighLightRowFlxGrid flxPurchHistory, flxPurchHistory.row

   ConfigFlxSplit flxPurchHistorySplit, 29

   With flxPurchHistorySplit
'         Adding the split of the header
      szSQL = "SELECT DISTINCT S.*, P.PropertyName AS XX " & _
              "FROM tblPurInvSRec AS S LEFT JOIN Property AS P ON S.TRANS = P.PropertyID " & _
              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' " & _
              "ORDER BY TRAN_ID;"
'
'      szSQL = szSQL + " UNION "
'
'      szSQL = szSQL + _
'              "SELECT DISTINCT S.*, U.UnitName AS XX " & _
'              "FROM tblPurInvSRec AS S, Units AS U " & _
'              "WHERE S.ParentID = '" & flxPurchHistory.TextMatrix(flxPurchHistory.row, 0) & "' AND " & _
'                  "S.TRANS = 'Unit' AND S.UNIT_ID = U.UnitNumber " & _
'              "ORDER BY TRAN_ID"
'Debug.Print szSQL
      adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

'   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
'               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"

      While Not adoInvSp.EOF
         iRow = iRow + 1
         .TextMatrix(iRow, 0) = adoInvSp.Fields.Item("MY_ID").Value
         .TextMatrix(iRow, 1) = adoInvSp.Fields.Item("TRAN_ID").Value
         .TextMatrix(iRow, 2) = adoInvSp.Fields.Item("TRANS").Value
         .TextMatrix(iRow, 3) = IIf(IsNull(adoInvSp.Fields.Item("XX").Value), "", adoInvSp.Fields.Item("XX").Value)
         .TextMatrix(iRow, 4) = adoInvSp.Fields.Item("NOMINAL_CODE").Value
         .TextMatrix(iRow, 5) = adoInvSp.Fields.Item("DEPT_ID").Value
         .TextMatrix(iRow, 6) = adoInvSp.Fields.Item("JOB_ID").Value
         .TextMatrix(iRow, 7) = adoInvSp.Fields.Item("DESCRIPTION").Value
         .TextMatrix(iRow, 8) = adoInvSp.Fields.Item("NET_AMOUNT").Value
         .TextMatrix(iRow, 9) = adoInvSp.Fields.Item("VAT").Value
         .TextMatrix(iRow, 10) = adoInvSp.Fields.Item("TOTAL_AMOUNT").Value

         adoInvSp.MoveNext
         If Not adoInvSp.EOF Then .AddItem ""
      Wend
      adoInvSp.Close
   End With

   adoConn.Close
   Set adoInvSp = Nothing
   Set adoConn = Nothing
End Sub

Private Sub Form_Activate()
   bFormLoaded = True
   bHistoryLoaded = False
End Sub

Private Sub Form_Load()
   bFormLoaded = False
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Height = 8685
   Me.Width = 12960
   Me.BackColor = MODULEBACKCOLOR

   tabPurExp.BackColor = MODULEBACKCOLOR
   bTotalPayTyped = False
   tabPurExp.Tab = 0
   iSelected = 0
   iPIEdit = 0
   cGridSPTotal = 0
   bChangesMade = False

   nTaxCode = 0
   bEditMode = False
   bPayAll = False
   txtClientID.text = "ALL"
   txtProperty.text = "ALL"
   cboAccount.text = "ALL"
   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   MousePointer = vbHourglass
   
   'PrepareList adoConn, cmbClient, cmbProperty

   LoadFlxPurchase adoConn
     Call WheelHook(Me.hWnd)
   MousePointer = vbDefault
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

'        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
'                        If Not ctl Is picClient Then
'                            PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
'                        Else
                            bHandled = False
'                        End If

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
Public Sub LoadFlxPurchase(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer, bFirstSp As Boolean
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset

   ConfigFlxPurchase
   ConfigFlxPurchaseSplit

   szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType, " & _
               "PI.TRAN_DATE, PI.SUPP_AC, Supplier.SupplierName, PI.PostingDate, " & _
               "PI.TOTAL_AMOUNT, PI.INV_NO, Pt.OSAmount, PI.PropertyID, PI.DueDate, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID AS ClientID, Pt.OSAmount, " & _
               "Supplier.SupplierOfficeEmail, Q.INVOICED, P.PropertyName " & _
           "FROM ((((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "LEFT JOIN tlbPayment AS Pt ON PI.MY_ID = Pt.PI) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID) " & _
               "LEFT JOIN Property AS P ON PI.PropertyID = P.PropertyID) " & _
               "LEFT JOIN (SELECT PO, SUM(TOTAL_AMOUNT) AS INVOICED " & _
                  "FROM tblPurInv " & _
                  "WHERE PO <> '' " & _
                  "GROUP BY PO) AS Q ON PI.MY_ID = Q.PO " & _
           "Where PI.History = False AND (PI.TransactionType = 25 OR PI.TransactionType = 26) " & _
           "ORDER BY 3, 2;"
'Debug.Print szSQL
   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iKount = 1
   With flxPurchase
      While Not adoInv.EOF
'         Adding the header of the invoice
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("PF").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 25, "P Order", "P Quote")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 11) = IIf(IsNull(adoInv.Fields.Item("PropertyID").Value), "", adoInv.Fields.Item("PropertyID").Value)
         .TextMatrix(iKount, 12) = Format(adoInv.Fields.Item("OSAmount").Value, "0.00")
         .TextMatrix(iKount, 13) = adoInv.Fields.Item("DueDate").Value
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("ClientID").Value), "", adoInv.Fields.Item("ClientID").Value)
         .TextMatrix(iKount, 15) = Format(adoInv.Fields.Item("INVOICED").Value, "0.00")
         .TextMatrix(iKount, 16) = IIf(IsNull(adoInv.Fields.Item("PostingDate").Value), "", adoInv.Fields.Item("PostingDate").Value)
         .TextMatrix(iKount, 17) = IIf(IsNull(adoInv.Fields.Item("SupplierOfficeEmail").Value), "", adoInv.Fields.Item("SupplierOfficeEmail").Value)
         .TextMatrix(iKount, 18) = IIf(IsNull(adoInv.Fields.Item("PropertyName").Value), "", adoInv.Fields.Item("PropertyName").Value)
'######################################################################################################################
'         Adding description of the header from the first split
         szSQL = "SELECT DISTINCT * " & _
                 "FROM tblPurInvSRec " & _
                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
                 "ORDER BY TRAN_ID;"
'Debug.Print szSQL
         adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         bFirstSp = True
         If Not adoInvSp.EOF Then _
            .TextMatrix(iKount, 8) = IIf(IsNull(adoInvSp.Fields.Item("DESCRIPTION").Value), "", adoInvSp.Fields.Item("DESCRIPTION").Value)

         adoInvSp.Close

         adoInv.MoveNext
         iKount = iKount + 1
         If Not adoInv.EOF Then .AddItem ""
      Wend
   End With

   adoInv.Close
   Set adoInv = Nothing
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i
   cboC.Column() = Data()
   cboC.ListIndex = 0
   adoRST.Close
'*************************************** PROPERTY ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then GoTo NoRes

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i

   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRST.Close
   Set adoRST = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub ConfigFlxPurchase()
   Dim szHeader As String, iCol As Integer

   flxPurchase.Clear
   flxPurchase.Cols = 19
   flxPurchase.Rows = 2
   flxPurchase.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID|>INVOICED|PostingDate|Email|PropertyName"

   flxPurchase.FormatString = szHeader$
   flxPurchase.ColWidth(0) = 0
   flxPurchase.ColWidth(1) = Label20(10).Left - flxPurchase.Left
   For iCol = 2 To flxPurchase.Cols - 10
      flxPurchase.ColWidth(iCol) = Label20(iCol + 9).Left - Label20(iCol + 8).Left
   Next iCol
   flxPurchase.ColWidth(iCol) = 0                        'Client
   flxPurchase.ColWidth(iCol + 1) = 0                    'Property
   flxPurchase.ColWidth(iCol + 2) = 0                    'OS Amt
   flxPurchase.ColWidth(iCol + 3) = 0                    'Due Date
   flxPurchase.ColWidth(iCol + 4) = 0                    'Client ID
   flxPurchase.ColWidth(iCol + 5) = flxPurchase.Width + flxPurchase.Left - Label20(18).Left - 340  'INVOICED
   flxPurchase.ColWidth(iCol + 6) = 0                    'Posting Date
   flxPurchase.ColWidth(iCol + 7) = 0                    'Posting Date
   flxPurchase.ColWidth(iCol + 8) = 0                    'Property Name
End Sub

Private Sub ConfigFlxPurchaseSplit()
   Dim szHeader As String, iCol As Integer
   Dim iLabel As Integer

   iLabel = 19

   flxPurchaseSplit.Clear
   flxPurchaseSplit.Cols = 12
   flxPurchaseSplit.Rows = 2
   flxPurchaseSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount|>Recoverable"
   flxPurchaseSplit.FormatString = szHeader$

   flxPurchaseSplit.ColWidth(0) = 0
   flxPurchaseSplit.ColWidth(1) = lblPurchaseSplit(1 + iLabel).Left - flxPurchaseSplit.Left

   For iCol = 2 To flxPurchaseSplit.Cols - 2
      flxPurchaseSplit.ColWidth(iCol) = lblPurchaseSplit(iCol + iLabel).Left - lblPurchaseSplit(iCol - 1 + iLabel).Left
   Next iCol
   flxPurchaseSplit.ColWidth(iCol) = flxPurchaseSplit.Width + flxPurchaseSplit.Left - lblPurchaseSplit(iCol - 1 + iLabel).Left - 340
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub fraEditDemand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub tabPurExp_Click(PreviousTab As Integer)
   If tabPurExp.Tab = 1 And Not bHistoryLoaded Then
      Dim adoConn As New ADODB.Connection
      Dim iRow As Integer

      'Set the ADO Connections to the dataset
      adoConn.Open getConnectionString

      LoadFlxPurchHistory adoConn

      adoConn.Close
      Set adoConn = Nothing
      bHistoryLoaded = True
   End If
End Sub

Private Sub tabPurExp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Public Sub PostInvoice()
   Dim szPI_ID As String
   Dim szSQL As String
   Dim iPosted As Integer               'Finally posted
   Dim iIP     As Integer               'To be posted
   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString

   If frmPopUpMenu.optSelPO.Value Then szPI_ID = SelectedPurInvID()
   If frmPopUpMenu.optPODtRange.Value Then szPI_ID = _
            DateRangePurInvID(CDate(frmPopUpMenu.txtDtRangeFromPO.text), CDate(frmPopUpMenu.txtDtRangeToPO.text))
   If frmPopUpMenu.optPONoRange.Value Then szPI_ID = _
            SlNoRangePurInv(frmPopUpMenu.txtPORangeFrom.text, frmPopUpMenu.txtPORangeTo.text)

   If Len(szPI_ID) = 0 Then
      ShowMsgInTaskBar "No purchase order to post to history.", "Y", "N"

      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   szSQL = "UPDATE tblPurInv " & _
           "SET History = TRUE " & _
           "WHERE MY_ID IN (" & szPI_ID & ");"

   adoConn.Execute szSQL

   LoadFlxPurchase adoConn
   LoadFlxPurchHistory adoConn

   MousePointer = vbDefault

   adoConn.Close
   Set adoConn = Nothing
   ShowMsgInTaskBar "System has posted " & iIP - iPosted & " purchase orders to history.", "Y", "P"
End Sub

Private Function DateRangePurInvID(dtFrom As Date, dtTo As Date) As String
   Dim i As Integer

   DateRangePurInvID = ""
   For i = 1 To flxPurchase.Rows - 1
      If CDate(flxPurchase.TextMatrix(i, 4)) >= dtFrom And CDate(flxPurchase.TextMatrix(i, 4)) <= dtTo Then
         DateRangePurInvID = DateRangePurInvID & "'" & flxPurchase.TextMatrix(i, 0) & "'"
         DateRangePurInvID = DateRangePurInvID & ","
      End If
   Next i
   If DateRangePurInvID <> "" Then
      DateRangePurInvID = Left(DateRangePurInvID, Len(DateRangePurInvID) - 1)
   Else
      DateRangePurInvID = ""
   End If
End Function

Private Function SlNoRangePurInv(szNoFrom As String, szNoTo As String) As String
   Dim i    As Integer
   Dim szIC As String
   Dim lF   As Long
   Dim lT   As Long

   szIC = UCase(Left(szNoFrom, 2))
   lF = StrDigitVal(szNoFrom)
   lT = StrDigitVal(szNoTo)

   SlNoRangePurInv = ""
   For i = 1 To flxPurchase.Rows - 1
      If Left(flxPurchase.TextMatrix(i, 2), 2) = szIC Then
         If StrDigitVal(flxPurchase.TextMatrix(i, 2)) >= lF And StrDigitVal(flxPurchase.TextMatrix(i, 2)) <= lT Then
            SlNoRangePurInv = SlNoRangePurInv & "'" & flxPurchase.TextMatrix(i, 0) & "'"
            SlNoRangePurInv = SlNoRangePurInv & ","
         End If
      End If
   Next i

   If SlNoRangePurInv <> "" Then
      SlNoRangePurInv = Left(SlNoRangePurInv, Len(SlNoRangePurInv) - 1)
   Else
      SlNoRangePurInv = ""
   End If
End Function

Private Sub LoadFlxPurchHistory(adoConn As ADODB.Connection)
   Dim szSQL As String, iKount As Integer, iChild As Integer
   Dim adoInv As New ADODB.Recordset, adoInvSp As New ADODB.Recordset

   ConfigFlxPurchHeader flxPurchHistory, 0
   ConfigFlxSplit flxPurchHistorySplit, 29

   szSQL = "SELECT DISTINCT PI.MY_ID, PI.SlNumber, PI.TransactionType, PI.TRAN_DATE, " & _
               "PI.SUPP_AC, Supplier.SupplierName, PI.TOTAL_AMOUNT, PI.INV_NO, " & _
               "MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS PF, PI.CL_ID " & _
           "FROM ((tblPurInv AS PI INNER JOIN Supplier ON PI.SUPP_AC = Supplier.SupplierID) " & _
               "INNER JOIN tblPurInvSRec AS S ON PI.MY_ID = S.ParentID) " & _
               "INNER JOIN tlbTransactionTypes AS T ON PI.TransactionType = T.TYPE_ID " & _
           "WHERE History = YES AND (PI.TransactionType = 25 OR PI.TransactionType = 26) " & _
           "ORDER BY PI.TransactionType, PI.SlNumber;"
'   Debug.Print szSQL

   adoInv.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iKount = 1
   With flxPurchHistory
      While Not adoInv.EOF
         .TextMatrix(iKount, 0) = adoInv.Fields.Item("MY_ID").Value
         .TextMatrix(iKount, 2) = adoInv.Fields.Item("pf").Value & IIf(IsNull(adoInv.Fields.Item("SlNumber").Value), "", adoInv.Fields.Item("SlNumber").Value)
         .TextMatrix(iKount, 3) = IIf(adoInv.Fields.Item("TransactionType").Value = 25, "P Order", "P Quote")
         .TextMatrix(iKount, 4) = IIf(IsNull(adoInv.Fields.Item("TRAN_DATE").Value), "", adoInv.Fields.Item("TRAN_DATE").Value)
         .TextMatrix(iKount, 5) = IIf(IsNull(adoInv.Fields.Item("SUPP_AC").Value), "", adoInv.Fields.Item("SUPP_AC").Value)
         .TextMatrix(iKount, 6) = IIf(IsNull(adoInv.Fields.Item("SupplierName").Value), "", adoInv.Fields.Item("SupplierName").Value)
         .TextMatrix(iKount, 7) = IIf(IsNull(adoInv.Fields.Item("INV_NO").Value), "", adoInv.Fields.Item("INV_NO").Value)
         .TextMatrix(iKount, 9) = Format(adoInv.Fields.Item("TOTAL_AMOUNT").Value, "0.00")
         .TextMatrix(iKount, 14) = IIf(IsNull(adoInv.Fields.Item("CL_ID").Value), "", adoInv.Fields.Item("CL_ID").Value)
'######################################################################################################################
'         Adding the split of the header
         szSQL = "SELECT DISTINCT * " & _
                 "FROM tblPurInvSRec " & _
                 "WHERE tblPurInvSRec.ParentID = '" & .TextMatrix(iKount, 0) & "' " & _
                 "ORDER BY TRAN_ID;"
'Debug.Print szSQL
         adoInvSp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

         If Not adoInvSp.EOF Then _
            .TextMatrix(iKount, 8) = IIf(IsNull(adoInvSp.Fields.Item("DESCRIPTION").Value), "", adoInvSp.Fields.Item("DESCRIPTION").Value)
         adoInvSp.Close

         adoInv.MoveNext
         iKount = iKount + 1
         If Not adoInv.EOF Then .AddItem ""
      Wend
   End With

   adoInv.Close
   Set adoInv = Nothing
End Sub

Private Sub ConfigFlxPurchHeader(ctrHeader As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrHeader.Clear
   ctrHeader.Cols = 15
   ctrHeader.Rows = 2
   ctrHeader.RowHeight(0) = 0

   szHeader$ = "TableID|>+-|<Transaction ID|<Transaction Type|<Transaction Date" & _
               "|<Suppplier ID|<Supplier Name|<Ref|<Desc|>Amount|<Client|<Property" & _
               "|>OS Amt|DueDate|ClientID"

   ctrHeader.FormatString = szHeader$
   ctrHeader.ColWidth(0) = 0
   ctrHeader.ColWidth(1) = Label20(1 + iLabel).Left - ctrHeader.Left
   For iCol = 2 To ctrHeader.Cols - 7
      ctrHeader.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrHeader.ColWidth(iCol) = ctrHeader.Width + ctrHeader.Left - Label20(iCol - 1 + iLabel).Left - 340
   ctrHeader.ColWidth(iCol + 1) = 0
   ctrHeader.ColWidth(iCol + 2) = 0
   ctrHeader.ColWidth(iCol + 3) = 0                   'OS Amt
   ctrHeader.ColWidth(iCol + 4) = 0                   'Due Date
   ctrHeader.ColWidth(iCol + 5) = 0                   'Client ID
End Sub

Private Sub ConfigFlxSplit(ctrSplit As MSHFlexGrid, iLabel As Integer)
   Dim szHeader As String, iCol As Integer

   ctrSplit.Clear
   ctrSplit.Cols = 11
   ctrSplit.Rows = 2
   ctrSplit.RowHeight(0) = 0

   szHeader$ = "TableID|<SL No|<Prop/Unit|<Prop/Unit Name|<N/C" & _
               "|<Fund|<Job No|<Desc|>Net|>VAT|>Amount"
   ctrSplit.FormatString = szHeader$

   ctrSplit.ColWidth(0) = 0
   ctrSplit.ColWidth(1) = Label20(1 + iLabel).Left - ctrSplit.Left

   For iCol = 2 To ctrSplit.Cols - 2
      ctrSplit.ColWidth(iCol) = Label20(iCol + iLabel).Left - Label20(iCol - 1 + iLabel).Left
   Next iCol
   ctrSplit.ColWidth(iCol) = ctrSplit.Width + ctrSplit.Left - Label20(iCol - 1 + iLabel).Left - 340
End Sub
