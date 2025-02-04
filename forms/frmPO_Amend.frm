VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BB5807FE-DBD2-11D3-87C1-4C980CC10374}#1.0#0"; "MyHover.ocx"
Begin VB.Form frmPO_Amend 
   BackColor       =   &H00DFDFDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14250
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPO_Amend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   14250
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   7560
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   57
      Top             =   8100
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
         TabIndex        =   62
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3345
         Left            =   45
         TabIndex        =   60
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   65
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   64
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   63
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
         Left            =   1875
         TabIndex        =   61
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         Left            =   0
         Top             =   120
         Width           =   5355
      End
   End
   Begin VB.Frame fraCmds 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   90
      TabIndex        =   104
      Top             =   7470
      Width           =   12525
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   400
         Index           =   0
         Left            =   5761
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C&lose"
         Height          =   400
         Index           =   0
         Left            =   10680
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1450
      End
      Begin VB.CommandButton cmdSavePI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   400
         Left            =   4200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1450
      End
      Begin MyHoverButton.Button cmdNew 
         Height          =   405
         Index           =   0
         Left            =   360
         TabIndex        =   105
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         HoverBackColor  =   15066597
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPO_Amend.frx":030A
         HoverPicture    =   "frmPO_Amend.frx":0326
         DisabledPicture =   "frmPO_Amend.frx":0342
         DownPicture     =   "frmPO_Amend.frx":035E
         MouseIcon       =   "frmPO_Amend.frx":037A
         Caption         =   "Add New____"
         HoverCaption    =   "Add New"
         DownCaption     =   ""
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DFDFDF&
      Height          =   4515
      Left            =   90
      TabIndex        =   83
      Top             =   2925
      Width           =   12525
      Begin VB.TextBox txtPICNTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "0.00"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtPICNNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "0.00"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtPICNVat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10485
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0.00"
         Top             =   4080
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit the Line"
         Height          =   355
         Left            =   45
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   4035
         Width           =   1440
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete the Line"
         Height          =   355
         Left            =   2085
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   4035
         Width           =   1450
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPI 
         Height          =   3480
         Left            =   90
         TabIndex        =   89
         Top             =   495
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   6138
         _Version        =   393216
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   14
         Left            =   11175
         TabIndex        =   102
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         Height          =   195
         Index           =   13
         Left            =   10320
         TabIndex        =   101
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T/C"
         Height          =   195
         Index           =   12
         Left            =   9735
         TabIndex        =   100
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
         Height          =   195
         Index           =   11
         Left            =   8925
         TabIndex        =   99
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         Height          =   195
         Index           =   10
         Left            =   3225
         TabIndex        =   98
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         Height          =   195
         Index           =   9
         Left            =   6855
         TabIndex        =   97
         Top             =   225
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   195
         Index           =   8
         Left            =   6375
         TabIndex        =   96
         Top             =   225
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job No"
         Height          =   195
         Index           =   7
         Left            =   1530
         TabIndex        =   95
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         Height          =   195
         Index           =   6
         Left            =   5715
         TabIndex        =   94
         Top             =   225
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
         Height          =   195
         Index           =   5
         Left            =   570
         TabIndex        =   93
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         Height          =   195
         Index           =   4
         Left            =   4620
         TabIndex        =   92
         Top             =   225
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   91
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   18
         Left            =   8205
         TabIndex        =   90
         Top             =   4080
         Width           =   390
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
         Height          =   195
         Index           =   49
         Left            =   180
         TabIndex        =   103
         Top             =   225
         Width           =   12300
      End
   End
   Begin VB.Frame fraControls 
      BackColor       =   &H00DFDFDF&
      Height          =   1590
      Left            =   90
      TabIndex        =   66
      Top             =   1350
      Width           =   12525
      Begin VB.TextBox txtDetails_ 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   2910
         TabIndex        =   68
         Top             =   1035
         Width           =   1760
      End
      Begin VB.CommandButton cmdUnitList 
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
         Left            =   4410
         TabIndex        =   9
         Top             =   150
         Width           =   255
      End
      Begin VB.CommandButton cmdJobNo 
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
         Index           =   0
         Left            =   8955
         TabIndex        =   12
         Top             =   165
         Width           =   255
      End
      Begin VB.TextBox txtUnit 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   135
         Width           =   3000
      End
      Begin VB.CommandButton cmdSchedules 
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
         Index           =   0
         Left            =   8955
         TabIndex        =   13
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdDeptList 
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
         Height          =   290
         Left            =   2205
         TabIndex        =   11
         Top             =   735
         Width           =   255
      End
      Begin VB.CommandButton cmdTaxList 
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
         Height          =   290
         Index           =   0
         Left            =   10935
         TabIndex        =   16
         Top             =   435
         Width           =   255
      End
      Begin VB.TextBox txtVat_ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   11190
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox txtNet_ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   10530
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   135
         Width           =   1875
      End
      Begin VB.TextBox txtDetails_ 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   5970
         TabIndex        =   14
         Top             =   735
         Width           =   3255
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         Height          =   375
         Index           =   1
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1110
         Width           =   1120
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   375
         Index           =   2
         Left            =   10020
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1110
         Width           =   1215
      End
      Begin VB.TextBox txtNC 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   435
         Width           =   795
      End
      Begin VB.CommandButton cmdNCList 
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
         Height          =   290
         Left            =   2205
         TabIndex        =   10
         Top             =   435
         Width           =   255
      End
      Begin VB.CheckBox chkRecover 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDFDF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1140
         TabIndex        =   67
         Top             =   1035
         Width           =   255
      End
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   735
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   735
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv Ref:"
         Height          =   195
         Index           =   1
         Left            =   2370
         TabIndex        =   82
         Top             =   1080
         Width           =   510
      End
      Begin MSForms.TextBox txtPFName 
         Height          =   285
         Left            =   2460
         TabIndex        =   30
         Top             =   735
         Width           =   2205
         VariousPropertyBits=   679495709
         BorderStyle     =   1
         Size            =   "3889;503"
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVatCode 
         Height          =   255
         Index           =   0
         Left            =   10410
         TabIndex        =   81
         Top             =   480
         Width           =   375
         VariousPropertyBits=   8388627
         Size            =   "661;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtTotal 
         Height          =   285
         Left            =   10530
         TabIndex        =   18
         Top             =   735
         Width           =   1875
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "3307;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   3
         Left            =   9930
         TabIndex        =   80
         Top             =   735
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule ID:"
         Height          =   195
         Index           =   2
         Left            =   5025
         TabIndex        =   79
         Top             =   435
         Width           =   885
      End
      Begin MSForms.TextBox txtJobNo 
         Height          =   285
         Left            =   5970
         TabIndex        =   31
         Top             =   135
         Width           =   3255
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5741;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job No:"
         Height          =   195
         Index           =   1
         Left            =   5025
         TabIndex        =   78
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name:"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   77
         Top             =   135
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net:"
         Height          =   195
         Index           =   0
         Left            =   9930
         TabIndex        =   76
         Top             =   135
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   0
         Left            =   5025
         TabIndex        =   75
         Top             =   735
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   74
         Top             =   735
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/C:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   73
         Top             =   435
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         Height          =   195
         Index           =   0
         Left            =   9930
         TabIndex        =   72
         Top             =   435
         Width           =   300
      End
      Begin MSForms.TextBox txtNCName 
         Height          =   285
         Left            =   2460
         TabIndex        =   27
         Top             =   435
         Width           =   2205
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "3889;503"
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSchedules 
         Height          =   285
         Left            =   5970
         TabIndex        =   32
         Top             =   435
         Width           =   3255
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5741;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recoverable:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   71
         Top             =   1035
         Width           =   915
      End
      Begin MSForms.TextBox txtRecoverable 
         Height          =   255
         Index           =   0
         Left            =   1410
         TabIndex        =   70
         Top             =   1035
         Width           =   540
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "952;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   2070
         TabIndex        =   69
         Top             =   1080
         Width           =   135
      End
   End
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   2640
      ScaleHeight     =   2895
      ScaleWidth      =   4815
      TabIndex        =   46
      Top             =   8100
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
         TabIndex        =   47
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2175
         Index           =   0
         Left            =   15
         TabIndex        =   48
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
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1350
         TabIndex        =   55
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
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   54
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
      Begin MSForms.Label lblSearch2 
         Height          =   195
         Left            =   3720
         TabIndex        =   53
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
      Begin MSForms.Label lblSearch1 
         Height          =   195
         Left            =   1560
         TabIndex        =   52
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
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   51
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   50
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   49
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
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   4500
      End
   End
   Begin VB.Frame fraLay 
      BackColor       =   &H00DFDFDF&
      Height          =   1335
      Left            =   75
      TabIndex        =   34
      Top             =   0
      Width           =   12555
      Begin VB.CommandButton cmdaddnewline 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add a &New Line"
         Height          =   345
         Left            =   10710
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   945
         Width           =   1560
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
         Left            =   12015
         TabIndex        =   6
         Top             =   225
         Width           =   255
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
         Left            =   12000
         TabIndex        =   7
         Top             =   620
         Width           =   255
      End
      Begin VB.CommandButton cmdACList 
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
         Height          =   285
         Index           =   0
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtInv 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3780
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   2985
      End
      Begin VB.TextBox txtAc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txtTransType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Purchase Order"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDueDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7620
         TabIndex        =   33
         Top             =   600
         Width           =   1080
      End
      Begin VB.TextBox txtTabSwitch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   13800
         TabIndex        =   35
         Text            =   "TabSwitch"
         Top             =   600
         Width           =   240
      End
      Begin VB.TextBox txtProperty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   2860
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7620
         TabIndex        =   4
         Top             =   240
         Width           =   1080
      End
      Begin MSForms.TextBox txtClientID 
         Height          =   255
         Left            =   9405
         TabIndex        =   5
         Top             =   225
         Width           =   2880
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5080;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox cmbSC 
         Height          =   285
         Left            =   1500
         TabIndex        =   45
         Top             =   600
         Width           =   1455
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "2566;503"
         Value           =   "Supplier"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txtSupplierName 
         Height          =   285
         Left            =   4840
         TabIndex        =   2
         Top             =   240
         Width           =   1920
         VariousPropertyBits=   679495709
         BorderStyle     =   1
         Size            =   "3387;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Category:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   43
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Type:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   41
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C:"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   40
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   16
         Left            =   8760
         TabIndex        =   39
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:"
         Height          =   195
         Index           =   19
         Left            =   6840
         TabIndex        =   38
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   9
         Left            =   8760
         TabIndex        =   37
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   5310
      TabIndex        =   56
      Top             =   6750
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmPO_Amend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bEditMode     As Boolean        'Is the PO in Edit mode?
Public szClientID    As String
Public szPropertyID  As String
Public szCallerForm  As String
Public szPO          As String
Private iCurEditRow  As Integer
Private sTextBox     As String
Private nTaxCode     As Double             'Tax code for Invoice
Private Const iXflxPI = 18
Private iSelected    As Integer
Public sPI           As String
Dim szaSupplierBal()    As String      'Supplier   balance
Dim iDayTerms           As String
Private sVCFound  As Single             'Vat code found either in Supplier (1) or Global Data (2)

Dim sAddChoice As String

Private Enum ComponentMode
   DefaultMode = 0
   newLine = 1
   EditLine = -1
   GridLostFocus = -2
   GridRowOnSelection = 2
   SavedMode = 3
   RefundMode = -3
   ExpensesMode = 4
End Enum

Private Sub ConfigFlxPI()
   With flxPI
      .Clear
      .Cols = 27
      .Rows = 2
      .RowHeight(0) = 0
      .ColWidth(0) = Label7(5).Left - Label7(3).Left  '"TransactionID" SL NO
      .ColWidth(1) = 0 'Label7(5).Left - Label7(4).Left '"A/C" SupplierId
      .ColWidth(2) = 0 'Label7(6).Left - Label7(5).Left '"Date"
      .ColWidth(3) = 0 ' Label7(7).Left - Label7(6).Left '"Type" Property
      .ColWidth(4) = 0 'Label7(8).Left - Label7(7).Left '"Trans"
      .ColWidth(5) = 0 'Label7(5).Left - Label7(4).Left '"Unit ID + Name"
      .ColWidth(6) = 0 'Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
      .ColWidth(7) = Label7(7).Left - Label7(5).Left + 20                    '"N/C"
      .ColWidth(8) = 0                      '"Dept"
      .ColWidth(9) = Label7(10).Left - Label7(7).Left + 20                    '"Job No"
      .ColWidth(10) = 0                     '"Cost Code"
      .ColWidth(11) = Label7(11).Left - Label7(10).Left + 20 '"Details"
      .ColWidth(12) = Label7(12).Left - Label7(11).Left + 20 '"Net"
      .ColWidth(13) = Label7(13).Left - Label7(12).Left + 20 '"T/C"
      .ColWidth(14) = Label7(14).Left - Label7(13).Left + 20 '"VAT"
      .ColWidth(15) = Label7(14).Left - Label7(13).Left + 300 '"Total"
      .ColWidth(16) = 0                     '"Sage"
      .ColWidth(17) = 0           'Stores PI Id hidenly
      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
      .ColWidth(19) = 0           'keep value 0 or 1 for edit
      .ColWidth(20) = 0 'Label7(13).Left - Label7(12).Left           'Stores ScheduleId
      .ColWidth(21) = 0           'Stores Unit ID
      .ColWidth(22) = 0           '% Recoverable
      .ColWidth(23) = 0           'ID
      .ColWidth(24) = 0 '.Width - Label7(1).Left - 120           'FundCode
      .ColWidth(25) = 0           'FundName
      .ColWidth(26) = 0           'not in use, and dont use it in this module
      
      
'      .ColWidth(0) = Label7(4).Left - .Left '"TransactionID"
'      .ColWidth(1) = Label7(5).Left - Label7(4).Left '"A/C"
'      .ColWidth(2) = Label7(6).Left - Label7(5).Left '"Date"
'      .ColWidth(3) = Label7(7).Left - Label7(6).Left '"Type"
'      .ColWidth(4) = Label7(8).Left - Label7(7).Left '"Trans"
'      .ColWidth(5) = Label7(9).Left - Label7(8).Left '"Unit ID + Name"
'      .ColWidth(6) = Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
'      .ColWidth(7) = 0                      '"N/C"
'      .ColWidth(8) = 0                      '"Dept"
'      .ColWidth(9) = 0                      '"Job No"
'      .ColWidth(10) = 0                     '"Cost Code"
'      .ColWidth(11) = Label7(11).Left - Label7(10).Left '"Details"
'      .ColWidth(12) = Label7(12).Left - Label7(11).Left '"Net"
'      .ColWidth(13) = Label7(13).Left - Label7(12).Left '"T/C"
'      .ColWidth(14) = Label7(14).Left - Label7(13).Left '"VAT"
'      .ColWidth(15) = .Width - Label7(14).Left - 120 '"Total"
'      .ColWidth(16) = 0                     '"Sage"
'      .ColWidth(17) = 0           'Stores PI Id hidenly
'      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
'      .ColWidth(19) = 0           'keep value 0 or 1 for edit
'      .ColWidth(20) = 0           'Stores ScheduleId
'      .ColWidth(21) = 0           'Stores Unit ID
'      .ColWidth(22) = 0           '% Recoverable
'      .ColWidth(23) = 0           'ID
'      .ColWidth(24) = 0           'FundCode
'      .ColWidth(25) = 0           'FundName
'      .ColWidth(26) = 0           'not in use, and dont use it in this module
      .row = 0
   End With

   txtPICNNet.Left = Label7(11).Left
   txtPICNNet.Width = flxPI.ColWidth(12)
   txtPICNVat.Left = Label7(13).Left
   txtPICNVat.Width = flxPI.ColWidth(14)
   txtPICNTotal.Left = Label7(14).Left
   txtPICNTotal.Width = flxPI.ColWidth(15)
End Sub

Private Sub cmdACList_Click(Index As Integer)
   LoadSupplierAccount

   txtSearch1.Visible = True
   txtSearch2.Visible = True
   txtSearch1.text = ""
   txtSearch2.text = ""
   fraList.Width = 4500
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtAc(0).Left + 100
   fraList.Top = txtAc(0).Top
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "A/C"
   txtSearch1.SetFocus
   fraLay.Enabled = False
   fraControls.Enabled = False
   Frame1.Enabled = False
   fraCmds.Enabled = False
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

      If cmbSC.text = "Supplier" Then
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
            UpdateBalance
      Else
         If cmbSC.text = "Client" Then
            szSQL = "SELECT ClientID, ClientName, spare2 " & _
                    "FROM Client " & _
                    "ORDER BY ClientName;"

            rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            .Clear
            .Rows = 2
            iRow = 1

            While Not rstRst.EOF
               .TextMatrix(iRow, 0) = rstRst!ClientID
               .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!ClientName), "", rstRst!ClientName)
               .TextMatrix(iRow, 2) = IIf(IsNull(rstRst!spare2), "", rstRst!spare2)
               rstRst.MoveNext
               If Not rstRst.EOF Then .AddItem ""
               iRow = iRow + 1
            Wend
         Else
'
'            If cmbSC.text = "Landlord" Then
'               szSQL = "SELECT LandlordID, LandlordName " & _
'                       "FROM Landlord " & _
'                       "ORDER BY LandlordName;"
'
'               .Clear
'               .Rows = 2
'               rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'               iRow = 1
'
'               While Not rstRst.EOF
'                  .TextMatrix(iRow, 0) = rstRst!landLordID
'                  .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!LandlordName), "", rstRst!LandlordName)
'                  rstRst.MoveNext
'                  If Not rstRst.EOF Then .AddItem ""
'                  iRow = iRow + 1
'               Wend
'            Else
            If cmbSC.text = "Managing Agent" Then
               szSQL = "SELECT AgentID, AgentName " & _
                       "FROM Agent " & _
                       "ORDER BY AgentName;"
               .Clear
               .Rows = 2
               rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
               iRow = 1

               While Not rstRst.EOF
                  .TextMatrix(iRow, 0) = rstRst!AgentID
                  .TextMatrix(iRow, 1) = IIf(IsNull(rstRst!AgentName), "", rstRst!AgentName)
                  rstRst.MoveNext
                  If Not rstRst.EOF Then .AddItem ""
                  iRow = iRow + 1
               Wend
            End If
'            End If
         End If
      End If
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

Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplier(0).Rows - 1
      For j = 0 To UBound(szaSupplierBal, 2) - 1
         If flxSupplier(0).TextMatrix(i, 0) = szaSupplierBal(0, j) Then
            flxSupplier(0).TextMatrix(i, 4) = Format(szaSupplierBal(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBal, 2) Then flxSupplier(0).TextMatrix(i, 4) = ""
   Next i
End Sub

Private Sub cmdaddnewline_Click()
    fraControls.Enabled = True
    cmdaddnewline.Enabled = False
    cmdNCList.SetFocus
End Sub

Private Sub cmdCancel_Click(Index As Integer)
   If bEditMode Then
      If MsgBox("Do you want to cancel Edit?", vbQuestion + vbYesNo, "Edit Record") = vbNo Then Exit Sub
      bEditMode = False
      iSelected = 0
      ConfigFlxPI
   Else
      If MsgBox("Do you want to cancel?" & Chr(13) & "If you wish to save the data you already entered click No", vbQuestion + vbYesNo, "Add Record") = vbNo Then Exit Sub

      ConfigFlxPI
   End If
   PIComponents DefaultMode

   HandleCommandButton "Cancel"
'   cmdUnitList.Enabled = False
   flxPI.Enabled = True
   flxPI.col = 0
   flxPI.CellBackColor = vbWhite

'   fraLay(1).Enabled = True
End Sub

Private Sub cmdClientSerc_Click()
    
    picClient.Left = 7290
    picClient.Top = 180
    LoadflxClient
    picClient.Visible = True
    fraLay.Enabled = False
    fraControls.Enabled = False
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
   
'   txtSearchClientID.Width = 1200
'   txtSearchClientID.Left = 40
'   txtSearchClientName.Width = 900
'   txtSearchClientName.Left = txtSearchClientID.Left + flxClient.ColWidth(0)

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
            rRow = 1
            While Not rstRec.EOF
               flxClient.row = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
   
  

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub

Private Sub cmdClose_Click(Index As Integer)
   If Not cmdNew(0).Enabled Or Not cmdEdit.Enabled And cmdSavePI.Visible Then
        'below line added by anol 08 Dec 2015
        If Left(Me.Caption, 4) <> "View" Then
            If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
               If cmdSavePI.Enabled Then cmdSavePI.SetFocus
               Exit Sub
            End If
        End If
   End If

   Unload Me
End Sub

Private Sub cmdDelete_Click()
   If flxPI.TextMatrix(1, 1) = "" Then Exit Sub

   If iSelected = 0 Then
      ShowMsgInTaskBar "Please select a record from the grid", , "N"
      Exit Sub
   End If

   If flxPI.row = 0 Then Exit Sub

   If MsgBox("Do you want to delete: " & flxPI.TextMatrix(iCurEditRow, 0) & "?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub

   Dim iRow    As Integer
   Dim iCol    As Integer
   Dim iGrids  As Integer
'
'   If flxPI.Rows = 2 And flxPI.row = 1 Then
'      ConfigFlxPI
'   End If
'
'   If flxPI.Rows > 2 Then
'      For iRow = iCurEditRow To flxPI.Rows - 2
'         For iCol = 1 To flxPI.Cols - 1
'            flxPI.TextMatrix(iRow, iCol) = flxPI.TextMatrix(iRow + 1, iCol)
'         Next iCol
'      Next iRow
'
'      flxPI.RemoveItem flxPI.Rows - 1
'   End If
   iCol = 1
   flxPI.RowHeight(flxPI.row) = 0
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.RowHeight(iRow) > 0 Then
         flxPI.TextMatrix(iRow, 0) = iCol
         iCol = iCol + 1
      End If
   Next iRow
   UpdateTotalPICN
End Sub

Private Sub cmdDeptList_Click()
   MousePointer = vbHourglass
   LoadDept
   
'   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   
   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtDept(0).Left + 100
   fraList.Top = txtDept(0).Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "Dept"
   MousePointer = vbDefault
   flxSupplier(0).SetFocus
End Sub

Private Sub LoadDept()
   flxSupplier(0).Rows = 3
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColWidth(2) = 0
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

   ' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   szSQL = "SELECT FundID, FundName, FundCode FROM Fund;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      flxSupplier(0).Clear
      
                 '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(0).Caption = "Fund Code"
      lblSearch1.Caption = "Fund Name"
      lblSearch2.Visible = False
      
      flxSupplier(0).RowHeight(0) = 0
      flxSupplier(0).Rows = 2

      rRow = 1
      While Not adoRst.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = adoRst.Fields.Item("FundCode").Value
         flxSupplier(0).TextMatrix(rRow, 1) = adoRst.Fields.Item("FundName").Value
         flxSupplier(0).TextMatrix(rRow, 2) = adoRst.Fields.Item("FundID").Value
         rRow = rRow + 1
         adoRst.MoveNext
         If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      Wend
   End If

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdEdit_Click()
   If flxPI.RowHeight(flxPI.row) = 0 Then Exit Sub
   If flxPI.TextMatrix(flxPI.row, 1) = "" Then Exit Sub

   If iSelected = 0 Then
      ShowMsgInTaskBar "Select a record in the grid to edit.", "Y", "N"
      Exit Sub
   End If
   cmdEdit.Enabled = False
   PIComponents EditLine

   With flxPI
      txtUnit(0).text = .TextMatrix(.row, 21)
      txtNC(0).text = .TextMatrix(.row, 7)
      txtDept(1).text = .TextMatrix(.row, 8)
      txtDept(0).text = .TextMatrix(.row, 24)
      txtPFName.text = .TextMatrix(.row, 25)
      txtJobNo.text = .TextMatrix(.row, 9)
      txtDetails_(0).text = .TextMatrix(.row, 11)
      txtNet_(0).text = .TextMatrix(.row, 12)
      lblVatCode(0).Caption = .TextMatrix(.row, 13)
      txtVat_(0).text = .TextMatrix(.row, 14)
      txtSchedules.text = .TextMatrix(.row, 20)
      txtRecoverable(0).text = .TextMatrix(.row, 22)
      chkRecover.Value = IIf(Val(txtRecoverable(0).text) > 0, 1, 0)
      txtTotal.text = .TextMatrix(.row, 15)

      sAddChoice = IIf(.TextMatrix(.row, 4) = "Invoice", "IN", "CN")
      bEditMode = True
      .TextMatrix(.row, 19) = "1"

      HandleCommandButton "Edit"
      .Enabled = False

      .row = 0
   End With
End Sub

Private Sub cmdGridUnitLookup_Click(Index As Integer)
'   If Index = 2 Then
'      tabPurExp.Enabled = True
'      picAccList.Visible = False
'      Exit Sub
'   End If
'   If Index = 1 Then
'      tabPurExp.Enabled = True
'      picAccounts.Visible = False
'      Exit Sub
'   End If
   fraLay.Enabled = True
   fraControls.Enabled = True
   Frame1.Enabled = True
   fraCmds.Enabled = True
   
   fraList.Visible = False

'   tabPurExp.Enabled = True
'   tabPayment.Enabled = True
'   fraList.Height = 2565
End Sub

Private Sub cmdJobNo_Click(Index As Integer)
   MousePointer = vbHourglass
   LoadJobSheet

   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtJobNo.Left + 100
   fraList.Top = txtJobNo.Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "job"
   MousePointer = vbDefault
   flxSupplier(0).SetFocus
End Sub

Private Sub LoadJobSheet()
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1400
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   Dim rRow As Integer, szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "Job No."
   lblSearch1.Caption = "Job Name"
   lblSearch2.Visible = False

   flxSupplier(0).Clear
   flxSupplier(0).Cols = 2
   flxSupplier(0).Rows = 2

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

'   szSQL = "SELECT ID, Job_DiaryName " & _
'           "FROM   PropertyMaintHistory " & _
'           "WHERE  AssignedTo = '" & txtAc(0).text & "' " & _
'           "ORDER BY ID;"

'10/12/2013
'Salia: any invoice will be connected with any job. any job (supplier/internal) can be attached.
   szSQL = "SELECT ID, Job_DiaryName " & _
           "FROM   PropertyMaintHistory " & _
           "WHERE  RecordType = 'J' AND PropertyID = '" & szPropertyID & "' " & _
           "ORDER BY ID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!ID
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!Job_DiaryName), "", rstRec!Job_DiaryName)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdNCList_Click()
   fraList.Height = 2925
   LoadNominalCode

   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtNC(0).Left + 100
   fraList.Top = txtNC(0).Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "NC"
   flxSupplier(0).SetFocus
End Sub




Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    fraLay.Enabled = True
    fraControls.Enabled = True
    cmdClientSerc.SetFocus
End Sub

Private Sub cmdSavePI_Click()
   If frmMMain.rtxtMessageDisplay.text = "Please update the invoice line." Then Exit Sub
   
   If Val(txtNet_(0).text) > 0 Then cmdUpdate_Click 1

   Dim adoConn As New ADODB.Connection
   Dim adoPIHeader As New ADODB.Recordset, adoPISplit As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, uID As String

   adoConn.Open getConnectionString

'  ***************************************************************************************************
'           SAVING HEADER PART OF THE PURCHASE ORDER                                                 '
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInv"
   If bEditMode Then szSQL = szSQL + " WHERE MY_ID = '" & sPI & "';"

   With adoPIHeader                                               'Add New Mode
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
      If Not bEditMode Then                'Add New Mode
         .AddNew
         uID = UniqueID()
        .Fields.Item("MY_ID").Value = uID
        .Fields.Item("CreatedBy").Value = User
        .Fields.Item("CreatedDate").Value = Now
        .Fields.Item("SlNumber").Value = SlNumber("PO", "tblPurInv", adoConn)
      Else
         uID = .Fields.Item("MY_ID").Value
      End If
'
'      If Not bEditMode Then
''         lSlNumber = SlNumber("PO", "tblPurInv", adoConn)
'      End If
      .Fields.Item("SUPP_AC").Value = txtAc(0).text
      .Fields.Item("TRAN_DATE").Value = Format(txtDate.text, "DD/MMMM/YYYY")
      .Fields.Item("TransactionType").Value = 25                                    'PURCHASE ORDER
      .Fields.Item("INV_NO").Value = txtInv(0).text
      .Fields.Item("TOTAL_AMOUNT").Value = CCur(txtPICNTotal.text)
      .Fields.Item("TTP").Value = CByte(TransactionTakePlace("TTP", "PURCHASE ORDER", adoConn))
      .Fields.Item("History").Value = False
      .Fields.Item("TrfPayment").Value = True
      .Fields.Item("PropertyID").Value = szPropertyID 'txtProperty.text
      .Fields.Item("CL_ID").Value = txtClientID.text
      If Len(txtDueDate.text) = 10 Then _
         .Fields.Item("DueDate").Value = Format(txtDueDate.text, "dd mmmm yyyy")

      .Update
      .Close
   End With

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE ORDER
'  ***************************************************************************************************
'   If Not bEditMode Then                                             'Edit Mode
'      adoConn.Execute "DELETE S.* " & _
'                      "FROM tlbPayment AS P, tlbPaymentSplit AS S " & _
'                      "WHERE PI = '" & sPI & "' AND " & _
'                            "P.TransactionID = S.PayHeader;"
'   End If

'Add New Records. At least there is only one split line
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(iRow, 0) <> "" Then
         If flxPI.RowHeight(iRow) > 0 Then
            szSQL = "SELECT * FROM tblPurInvSRec WHERE MY_ID = '" & flxPI.TextMatrix(iRow, 23) & "';"
            
            With adoPISplit
               .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
               If .EOF Then
                  .AddNew
                  .Fields.Item("MY_ID").Value = UniqueID()
                  .Fields.Item("ParentID").Value = uID
               End If
               .Fields.Item("TRAN_ID").Value = flxPI.TextMatrix(iRow, 0)
               .Fields.Item("TRANS").Value = szPropertyID
               .Fields.Item("UNIT_ID").Value = flxPI.TextMatrix(iRow, 21)
               .Fields.Item("NOMINAL_CODE").Value = flxPI.TextMatrix(iRow, 7)
               .Fields.Item("DEPT_ID").Value = flxPI.TextMatrix(iRow, 8)
               .Fields.Item("JOB_ID").Value = Mid(flxPI.TextMatrix(iRow, 9), 6)            'Job No
               .Fields.Item("COST_CODE").Value = flxPI.TextMatrix(iRow, 10)
               .Fields.Item("description").Value = flxPI.TextMatrix(iRow, 11)
               .Fields.Item("NET_AMOUNT").Value = CCur(flxPI.TextMatrix(iRow, 12))
               .Fields.Item("TAX_CODE").Value = flxPI.TextMatrix(iRow, 13)
               .Fields.Item("VAT").Value = CCur(flxPI.TextMatrix(iRow, 14))
               .Fields.Item("ScheduleID").Value = IIf(flxPI.TextMatrix(iRow, 20) = "", Null, _
                                                      flxPI.TextMatrix(iRow, 20))
               .Fields.Item("TOTAL_AMOUNT").Value = CCur(flxPI.TextMatrix(iRow, 14)) + _
                                                    CCur(flxPI.TextMatrix(iRow, 12))
               .Fields.Item("RecoverablePt").Value = flxPI.TextMatrix(iRow, 22)
   
               .Update
               .Close
            End With
         Else              'user had deleted the line
            adoConn.Execute "UPDATE tblPurInvSRec " & _
                            "SET description = 'DELETED' " & _
                            "WHERE MY_ID = '" & flxPI.TextMatrix(iRow, 23) & "';"
         End If
      End If
   Next iRow
'
''  User has deleted all split line.
'   If flxPI.TextMatrix(1, 0) = "" Then
'      With adoPISplit
'         .AddNew
'         .Fields.Item("MY_ID").Value = UniqueID()
'         .Fields.Item("ParentID").Value = uID
'         .Fields.Item("TRAN_ID").Value = 1
'         .Fields.Item("TRANS").Value = szPropertyID 'txtProperty.text
'         .Fields.Item("NOMINAL_CODE").Value = "0000"
'         .Fields.Item("description").Value = "DELETED ALL SPLITS"
'         .Fields.Item("NET_AMOUNT").Value = 0
'         .Fields.Item("TAX_CODE").Value = "T9"
'         .Fields.Item("VAT").Value = 0
'         .Fields.Item("TOTAL_AMOUNT").Value = 0
'         .Fields.Item("RecoverablePt").Value = 0
'         .Update
'      End With
'   End If
'   adoPISplit.Close

'  If the PO form is open then refresh the form
   RefreshPO adoConn
   'LoadFlxPurchase adoConn
   adoConn.Close

   Set adoPISplit = Nothing
   Set adoPIHeader = Nothing
   Set adoConn = Nothing

   PIComponents DefaultMode
   If bEditMode = False Then
'      fraLay(0).Top = Me.Height + 300
      frmPO.cmdEdit(1).Enabled = True
'   Else
'      HandleCommandButton "Save"
'      cmdNew(0).SetFocus
   End If
 
   ShowMsgInTaskBar "Data has been saved successfully."
   flxPI.Clear
   flxPI.Rows = 2
   Me.Hide
   Unload Me
'   lblSearch0(5).Caption = "NotLoaded"
End Sub

Private Sub RefreshPO(adoConn As ADODB.Connection)
   If IsLoadedAndVisible("frmPO") Then
      frmPO.LoadFlxPurchase adoConn
   End If
End Sub

Private Sub cmdSchedules_Click(Index As Integer)
   LoadSchedules

'   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtSchedules.Left + 100
   fraList.Top = txtSchedules.Top + 350
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "Schedules"
   flxSupplier(0).SetFocus
End Sub

Private Sub LoadSchedules()
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify
   
   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)
   
   txtSearch1.Width = 1400
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)
   
   Dim rRow As Integer
   Dim adoConn As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset
   
   '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(0).Caption = "Schedule ID"
      lblSearch1.Caption = "Schedule Name"
      lblSearch2.Visible = False
      
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   szSQL = "SELECT ScheduleID, ScheduleName " & _
           "FROM Schedule " & _
           "ORDER BY ScheduleID;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then


      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0
   
      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!ScheduleID
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!ScheduleName), "", rstRec!ScheduleName)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   adoConn.Close
End Sub

Private Sub cmdTaxList_Click(Index As Integer)
   LoadVAT

   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""
   txtSearch2.Width = 1000
   fraList.Width = 2400
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtVat_(0).Left - 400
   fraList.Top = txtVat_(0).Top + txtVat_(0).Height
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "VAT"
   flxSupplier(0).SetFocus
End Sub

Private Sub LoadVAT()
   flxSupplier(0).ColWidth(0) = 1000
   flxSupplier(0).ColWidth(1) = 1000
   flxSupplier(0).TextMatrix(0, 0) = "CODE"
   flxSupplier(0).TextMatrix(0, 1) = "RATE"

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 900
   lblSearch0(0).Left = 50
   lblSearch1.Width = 1900
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 900
   txtSearch1.Left = 40

   txtSearch2.Width = 1900
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "CODE"
   lblSearch1.Caption = "RATE"
   lblSearch2.Visible = False

   flxSupplier(0).RowHeight(0) = 0

   Dim rRow As Integer
   Dim Conn2 As New ADODB.Connection

   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset

'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString

   szSQL = "SELECT VAT_CODE, VAT_RATE " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(1) = vbRightJustify

      flxSupplier(0).TextMatrix(0, 0) = "VAT Code"
      flxSupplier(0).TextMatrix(0, 1) = "VAT Rate"

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!VAT_CODE
         flxSupplier(0).TextMatrix(rRow, 1) = rstRec!VAT_RATE
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
   Conn2.Close
   
   Set rstRec = Nothing
   Set Conn2 = Nothing
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

'   On Error Resume Next

   szSQL = "SELECT PropertyID, PropertyName " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyID;"
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   rRow = 1
   While Not rstRec.EOF
      flxSupplier(0).TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
      flxSupplier(0).TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
      rstRec.MoveNext
      If Not rstRec.EOF Then flxSupplier(0).AddItem ""
      rRow = rRow + 1
   Wend

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdTypeList_Click()
   If txtAc(0).text = "" Then
      cmdACList(0).SetFocus
      ShowMsgInTaskBar "Please select the " & cmbSC.text & ".", "Y", "N"
      Exit Sub
   End If

   LoadPropertyList

'   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
'   fraList.Left = txtProperty.Left + fraLay(0).Left + 100
   fraList.Left = txtProperty.Left + txtProperty.Width - fraList.Width '+ fraLay(0).Left
   fraList.Top = txtProperty.Top '+ fraLay(0).Top '+ tabPurExp.Top '+ 380
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "PROPERTY"
   txtSearch1.SetFocus
End Sub

Private Sub LoadUnitList(szSQL As String, Conn2 As ADODB.Connection)
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

   '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1400
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   Dim rRow As Integer
   Dim rstRec As New ADODB.Recordset

   If szSQL = "" Then
      szSQL = "SELECT UnitNumber, UnitName " & _
              "FROM Units, Property AS P " & _
              "WHERE Units.PropertyID = P.PropertyID AND " & _
                  "P.PropertyID = '" & szPropertyID & "' AND " & _
                  "P.ClientID = '" & txtClientID.text & "' " & _
              "ORDER BY UnitNumber"
   End If
   rstRec.Open szSQL, Conn2, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxSupplier(0).Clear
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2

      rstRec.MoveFirst
      flxSupplier(0).ColAlignment(0) = vbRightJustify

      flxSupplier(0).RowHeight(0) = 0
      '~~~Added By Senthuran~~~ Code to configuer Label Caption
      lblSearch0(0).Caption = "Unit ID"
      lblSearch1.Caption = "Unit Name"
      lblSearch2.Visible = False

      rRow = 1
      While Not rstRec.EOF
         flxSupplier(0).TextMatrix(rRow, 0) = rstRec!UnitNumber
         flxSupplier(0).TextMatrix(rRow, 1) = IIf(IsNull(rstRec!UnitName), "", rstRec!UnitName)
         rstRec.MoveNext
         If Not rstRec.EOF Then flxSupplier(0).AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close
End Sub

Private Sub cmdUnitList_Click()
   If txtProperty.text = "" Then
      cmdTypeList.SetFocus
      Exit Sub
   End If

   fraList.Height = 2925

   Dim Conn2 As New ADODB.Connection
'   Reset screen to show all the units in cboUnits.
'   Set the RDO Connections to the dataset
   Conn2.Open getConnectionString

   LoadUnitList "", Conn2
   Conn2.Close
   Set Conn2 = Nothing

'   tabPayment.Enabled = False
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   fraList.Width = 4815
   cmdGridUnitLookup(0).Left = fraList.Width - cmdGridUnitLookup(0).Width
   Shape4(0).Width = fraList.Width - cmdGridUnitLookup(0).Width - 50
   flxSupplier(0).Width = fraList.Width - 50
   fraList.Left = txtUnit(0).Left + 100
   fraList.Top = txtUnit(0).Top + 380
   fraList.Visible = True
   fraList.ZOrder 0
   sTextBox = "UNIT"
   flxSupplier(0).SetFocus
End Sub

Private Sub cmdUpdate_Click(Index As Integer)
   If Index = 1 Then                                  'OK
      If txtDate.text = "" Then
         ShowMsgInTaskBar "You must enter the date from the list.", "Y", "N"
         txtDate.SetFocus
         Exit Sub
      End If
      If txtDueDate.text = "" Then
         ShowMsgInTaskBar "You must enter the due date from the list.", "Y", "N"
         txtDueDate.SetFocus
         Exit Sub
      End If
      If txtNC(0).text = "" Then
         ShowMsgInTaskBar "You must select Nominal Code from the list.", "Y", "N"
         cmdNCList.SetFocus
         Exit Sub
      End If

      If txtDept(0).text = "" Then
         ShowMsgInTaskBar "You must select a fund from the list.", "Y", "N"
         cmdDeptList().SetFocus
         Exit Sub
      End If
      If Val(txtNet_(0).text) <= 0 Then
         ShowMsgInTaskBar "You must enter the amount.", "Y", "N"
         txtNet_(0).SetFocus
         Exit Sub
      End If
      If chkRecover.Value = 1 And Val(txtRecoverable(0).text) = 0 Then
         ShowMsgInTaskBar "You must enter the amount.", "Y", "N"
         txtNet_(0).SetFocus
         Exit Sub
      End If

      With flxPI
         If cmdEdit.Enabled Then                                 ' ****************  ADD NEW PI  ************************
            If Not (.Rows = 2 And .TextMatrix(1, 1) = "") Then
               .AddItem ""
            End If
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = txtAc(0).text
            .TextMatrix(.Rows - 1, 2) = txtDate.text
            .TextMatrix(.Rows - 1, 3) = txtProperty.text
            .TextMatrix(.Rows - 1, 4) = IIf(sAddChoice = "IN" Or sAddChoice = "AI", "Invoice", "Credit")
            .TextMatrix(.Rows - 1, 5) = txtUnit(0).text
            .TextMatrix(.Rows - 1, 6) = txtInv(0).text
            .TextMatrix(.Rows - 1, 7) = txtNC(0).text
            .TextMatrix(.Rows - 1, 8) = txtDept(1).text
            .TextMatrix(.Rows - 1, 9) = txtJobNo.text
            .TextMatrix(.Rows - 1, 11) = txtDetails_(0).text
            .TextMatrix(.Rows - 1, 12) = txtNet_(0).text
            .TextMatrix(.Rows - 1, 13) = lblVatCode(0).Caption
            .TextMatrix(.Rows - 1, 14) = txtVat_(0).text
            .TextMatrix(.Rows - 1, 20) = txtSchedules.text
            .TextMatrix(.Rows - 1, 21) = txtUnit(0).text
            .TextMatrix(.Rows - 1, 22) = IIf(txtRecoverable(0).text = "", 0, txtRecoverable(0).text)
            .TextMatrix(.Rows - 1, 15) = Format(txtTotal.text, "0.00")
            .TextMatrix(.Rows - 1, 23) = UniqueID()
            .TextMatrix(.Rows - 1, 24) = txtDept(0).text
            .TextMatrix(.Rows - 1, 25) = txtPFName.text
         Else                                                  ' ****************  Update PI  ************************
            .TextMatrix(iCurEditRow, 5) = txtUnit(0).text
            .TextMatrix(iCurEditRow, 6) = txtInv(0).text
            .TextMatrix(iCurEditRow, 7) = txtNC(0).text
            .TextMatrix(iCurEditRow, 8) = txtDept(1).text
            .TextMatrix(iCurEditRow, 11) = txtDetails_(0).text
            .TextMatrix(iCurEditRow, 12) = txtNet_(0).text
            .TextMatrix(iCurEditRow, 13) = lblVatCode(0).Caption
            .TextMatrix(iCurEditRow, 14) = txtVat_(0).text
            .TextMatrix(iCurEditRow, 19) = ""
            .TextMatrix(iCurEditRow, 20) = txtSchedules.text
            .TextMatrix(iCurEditRow, 21) = txtUnit(0).text
            .TextMatrix(iCurEditRow, 22) = IIf(txtRecoverable(0).text = "", 0, txtRecoverable(0).text)
            .TextMatrix(iCurEditRow, 15) = Format(txtTotal.text, "0.00")
'            .TextMatrix(iCurEditRow, 23) = UniqueID()
            .TextMatrix(iCurEditRow, 24) = txtDept(0).text
            .TextMatrix(iCurEditRow, 25) = txtPFName.text
            HandleCommandButton "Update Record"
         End If
         PIComponents newLine
      End With

      UpdateTotalPICN

      cmdEdit.Enabled = True
      'added by anol 16 Dec 2015
         bEditMode = False
         cmdSavePI.Enabled = True
         cmdCancel(0).Enabled = True
'      If txtProperty.text = "" Then
'         cmdNCList.SetFocus
'      Else
'         cmdUnitList.SetFocus
'      End If
    fraControls.Enabled = False
    cmdaddnewline.Enabled = True
    cmdaddnewline.SetFocus
    'end of addition
   End If

   If Index = 2 Then                   'Clear
      PIComponents EditLine
      flxPI.Enabled = True
        
'      If txtProperty.text = "" Then
'         cmdNCList.SetFocus
'      Else
'         cmdUnitList.SetFocus
'      End If
'
'      If txtProperty.text = "" Then
'         cmdNCList.SetFocus
'      Else
'         cmdUnitList.SetFocus
'      End If
      fraControls.Enabled = False
      cmdaddnewline.Enabled = True
      cmdaddnewline.SetFocus
     
   End If
End Sub

Private Sub flxClient_Click()
    fraLay.Enabled = True
    fraControls.Enabled = True
    txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
    picClient.Visible = False
    txtProperty.text = ""
    cmdTypeList.SetFocus
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fraLay.Enabled = True
        fraControls.Enabled = True
        txtClientID.text = flxClient.TextMatrix(flxClient.row, 0)
        picClient.Visible = False
        cmdTypeList.SetFocus
    End If
End Sub

Private Sub flxPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   iSelected = 1
   SelectOnly1RowFlxGrid flxPI, flxPI.row, iXflxPI
End Sub

Private Sub flxPI_RowColChange()
   iCurEditRow = flxPI.row
End Sub

Private Sub flxSupplier_Click(Index As Integer)
'   If Index = 2 Then
'      cboAccount.Value = flxSupplier(2).TextMatrix(flxSupplier(2).row, 1)
'      SortTheGrid flxPurchase, cmbClient, cmbProperty, cboAccount
'      flxPurchaseSplit.Clear
'      cmdGridUnitLookup_Click (2)
'      Exit Sub
'   End If
'   If Index = 1 Then
'      cmbSPSupplier.Value = flxSupplier(1).TextMatrix(flxSupplier(1).row, 1)
'      cmdGridUnitLookup_Click (1)
'      Exit Sub
'   End If

'   tabPurExp.Enabled = True
        fraLay.Enabled = True
        fraControls.Enabled = True
        Frame1.Enabled = True
        fraCmds.Enabled = True
   
    
   
   If sTextBox = "A/C" Then
'      bTotalPayTyped = False
      txtAc(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      If cmbSC.text = "Client" Then
         txtClientID.text = txtAc(0).text
         txtClientID.Locked = True
      Else
         txtClientID.Locked = False
      End If

      txtSupplierName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      If txtNC(0).text = "" Then _
         txtNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
       'below line is added by anol 07 jan 2015
       'issue 469
      If txtNC(0).text = "0" Then
         txtNC(0).text = ""
         txtNCName.text = ""
      End If
      txtInv(0).SetFocus
      txtAc(0).SelStart = Len(txtAc(0).text)
      iDayTerms = Val(flxSupplier(0).TextMatrix(flxSupplier(0).row, 5))
      txtDueDate.text = DateAdd("d", iDayTerms, Date)

      Dim szaTemp() As String

      If InStr(flxSupplier(0).TextMatrix(flxSupplier(0).row, 3), "##") > 0 Then
         szaTemp = Split(flxSupplier(0).TextMatrix(flxSupplier(0).row, 3), "##")
         lblVatCode(0).Caption = szaTemp(0)
         If szaTemp(1) = "" Then
            nTaxCode = -1
         Else
            nTaxCode = CDbl(szaTemp(1))
         End If
         sVCFound = 1
      Else
         lblVatCode(0).Caption = ""
         txtProperty.text = ""
         sVCFound = 2
      End If
   End If
   If sTextBox = "PROPERTY" Then
      szPropertyID = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtProperty.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      cmdaddnewline.SetFocus
   End If
   If sTextBox = "UNIT" Then
      txtUnit(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      cmdNCList.SetFocus
      txtUnit(0).SelStart = Len(txtUnit(0).text)
      cmdJobNo(0).Enabled = True
   End If
   If sTextBox = "NC" Then
      txtNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      cmdDeptList().SetFocus
'      txtNC(0).SelStart = Len(txtNC(0).text)
   End If
   If sTextBox = "Dept" Then
      txtDept(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtPFName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      txtDept(1).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
      cmdJobNo(0).SetFocus
'      txtDept(0).SelStart = Len(txtDept(0).text)
   End If
   If sTextBox = "VAT" Then
      lblVatCode(0).Caption = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      nTaxCode = CSng(flxSupplier(0).TextMatrix(flxSupplier(0).row, 1))
      txtNet__LostFocus (0)
      cmdUpdate(1).SetFocus
   End If
   If sTextBox = "Schedules" Then
      txtSchedules.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      txtDetails_(0).SetFocus
      txtSchedules.SelStart = Len(txtSchedules.text)
   End If
   If sTextBox = "job" Then
      txtJobNo.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      cmdSchedules(0).SetFocus
      txtJobNo.SelStart = Len(txtSchedules.text)
   End If
   If sTextBox = "Bank" Then nTaxCode = TaxRate(1)
   If sTextBox = "VATBank" Then nTaxCode = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)

'   tabPayment.Enabled = True
'   Me.Enabled = True
   fraList.Visible = False
End Sub

Private Sub flxSupplier_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      flxSupplier(0).Clear
      flxSupplier(0).Clear
      
      flxSupplier(0).Cols = 2
      flxSupplier(0).Rows = 2
      fraList.Visible = False
'      tabPurExp.Enabled = True
'      If sTextBox = "A/C" Then cmdACList(0).SetFocus
      If sTextBox = "UNIT" Or sTextBox = "PROP" Then cmdNCList.SetFocus
      If sTextBox = "NC" Then txtNC(0).SetFocus
      If sTextBox = "Dept" Then txtDept(0).SetFocus
      If sTextBox = "VAT" Then cmdTaxList(0).SetFocus
      Exit Sub
   End If
   If KeyAscii = 13 Then
      flxSupplier_Click (0)
   End If
End Sub

Private Sub Form_Load()
   Me.Height = 8580
   Me.Width = 12645
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   ConfigFlxPI
   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString
   SupplierAccountBalance adoConn
   'LoadCboClientPI adoConn

   adoConn.Close
   Set adoConn = Nothing
   If Len(txtDate.text) < 10 Then
      txtDate.text = Format(Date, "dd/mm/yyyy")
   End If
   fraControls.Enabled = False
    Call WheelHook(Me.hWnd)
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
Private Sub txtDetails__KeyPress(Index As Integer, KeyAscii As Integer)
        If Index = 0 And KeyAscii = 13 Then
            txtNet_(0).SetFocus
        End If
End Sub

Private Sub txtInv_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 And Index = 0 Then
        txtDate.SetFocus
    End If
    
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
    If KeyAscii = 13 Then
        txtSearch2.SetFocus
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

Private Sub txtSearch2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        flxSupplier(0).SetFocus
    End If
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
          fraLay.Enabled = True
          fraControls.Enabled = True
          'Resolved by BOSL
          'Below line are modified by anol 29 Mar 2015
          'issue 553 : PRESTIGE GUI IMPROVEMENT
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
'Private Sub LoadCboClientPI(adoConn As ADODB.Connection)
'   Dim szSQL   As String
'   Dim adoRst  As New ADODB.Recordset
'
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CT " & _
'           "FROM   CLIENT " & _
'           "ORDER BY CLIENTNAME;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   Dim Data() As String
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboClientPI.Column() = Data()
'   cboClientPI.ListIndex = 0
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Function GetSupplierBalance(szSuppID As String) As Currency
   Dim j As Integer

   For j = 0 To UBound(szaSupplierBal, 2) - 1
      If szSuppID = szaSupplierBal(0, j) Then
         GetSupplierBalance = Format(szaSupplierBal(1, j), "0.00")
         Exit For
      End If
   Next j
   If j = UBound(szaSupplierBal, 2) Then GetSupplierBalance = 0
End Function
'  Build up SUPPLIERs' Account BALANCE
Private Sub SupplierAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Dr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type = 6 Or P.Type = 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.SupplierID " & _
               "WHERE S.TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaSupplierBal(1, adoPayDr.RecordCount) As String

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT X.SupplierID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
           "FROM ( " & _
               "SELECT S.SupplierID, P.Cr " & _
               "FROM Supplier AS S LEFT JOIN ( " & _
                  "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
                  "FROM tlbPayment AS P " & _
                  "Where P.Type <> 6 And P.Type <> 24 " & _
                  "GROUP BY P.SageAccountNumber) AS P ON P.SageAccountNumber = S.SupplierID " & _
               "WHERE TYPE = 'SUPPLIER' " & _
           ") AS X;"

'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

'######################################      CLIENT         ##############################################
   iIndex = UBound(szaSupplierBal, 2)

   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.ClientID, P.Dr " & _
               "FROM Client AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type = 6 Or P.Type = 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.ClientID " & _
               ") AS X;"

'Debug.Print szSQL
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String

   While Not adoPayDr.EOF
      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT X.ClientID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
           "FROM ( " & _
               "SELECT S.ClientID, P.Cr " & _
               "FROM Client AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type <> 6 And P.Type <> 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.ClientID " & _
               ") AS X;"

'Debug.Print szSQL
   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i <= iIndex - 1 Then
         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
         iIndex = iIndex + 1
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close
'
''######################################      LANDLORD       ##############################################
'   iIndex = UBound(szaSupplierBal, 2)
'
'   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
'           "FROM ( " & _
'               "SELECT S.LandlordID, P.Dr " & _
'               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type = 6 Or P.Type = 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.LandlordID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String
'
'   While Not adoPayDr.EOF
'      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
'      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
'      iIndex = iIndex + 1
'      adoPayDr.MoveNext
'   Wend
'
'   adoPayDr.Close
'
'   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
'           "FROM ( " & _
'               "SELECT S.LandlordID, P.Cr " & _
'               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type <> 6 And P.Type <> 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.LandlordID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoPayCr.EOF
'      For i = 0 To iIndex - 1
'         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'            Exit For
'         End If
'      Next i
'      If i <= iIndex - 1 Then
'         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
'      Else
'         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
'         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
'         iIndex = iIndex + 1
'      End If
'      adoPayCr.MoveNext
'   Wend
'
'   adoPayCr.Close

'######################################      AGENT       ##############################################
   iIndex = UBound(szaSupplierBal, 2)

   szSQL = "SELECT DISTINCT X.AgentID AS SageAccountNumber, IIF(ISNULL(X.Dr), 0, X.Dr) AS Dr " & _
           "FROM ( " & _
               "SELECT S.AgentID, P.Dr " & _
               "FROM Agent AS S LEFT JOIN ( " & _
                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Dr " & _
                     "FROM tlbPayment AS P " & _
                     "Where (P.Type = 6 Or P.Type = 24) " & _
                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
                           "P.SageAccountNumber = S.AgentID " & _
               ") AS X;"

'Debug.Print szSQL
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim Preserve szaSupplierBal(1, iIndex + adoPayDr.RecordCount) As String

   While Not adoPayDr.EOF
      szaSupplierBal(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBal(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close
'
'   szSQL = "SELECT DISTINCT X.LandlordID AS SageAccountNumber, IIF(ISNULL(X.Cr), 0, X.Cr) AS Cr " & _
'           "FROM ( " & _
'               "SELECT S.LandlordID, P.Cr " & _
'               "FROM PropertyLandlord AS S LEFT JOIN ( " & _
'                     "SELECT P.SageAccountNumber, SUM(P.Amount) AS Cr " & _
'                     "FROM tlbPayment AS P " & _
'                     "Where (P.Type <> 6 And P.Type <> 24) " & _
'                     "GROUP BY P.SageAccountNumber) AS P ON  " & _
'                           "P.SageAccountNumber = S.LandlordID " & _
'               ") AS X;"
'
''Debug.Print szSQL
'   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   While Not adoPayCr.EOF
'      For i = 0 To iIndex - 1
'         If szaSupplierBal(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
'            Exit For
'         End If
'      Next i
'      If i <= iIndex - 1 Then
'         szaSupplierBal(1, i) = szaSupplierBal(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
'      Else
'         szaSupplierBal(0, iIndex) = adoPayCr.Fields.Item("SageAccountNumber").Value
'         szaSupplierBal(1, iIndex) = adoPayCr.Fields.Item("Cr").Value
'         iIndex = iIndex + 1
'      End If
'      adoPayCr.MoveNext
'   Wend
'
'   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Sub

Private Sub HandleCommandButton(szButton As String)
   Select Case szButton
      Case "Save"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = True
         cmdCancel(0).Enabled = False

         flxPI.Enabled = True
         flxPI.col = 0
         flxPI.CellBackColor = vbWhite

         ConfigFlxPI

      Case "Add Invoice"
         cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = True
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = True

      Case "Edit"
         cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = False

      Case "Cancel"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdNew(0).Enabled = True
         cmdCancel(0).Enabled = False

      Case "Update Record"
         cmdSavePI.Enabled = True
         cmdNew(0).Enabled = False
         cmdCancel(0).Enabled = True
         flxPI.Enabled = True
         flxPI.row = 0
   End Select
End Sub

Private Sub PIComponents(ByVal c_mode As ComponentMode)
   Select Case c_mode

   Case ComponentMode.DefaultMode
      cmbSC.Enabled = True
      txtAc(0).text = ""
      txtSupplierName.text = ""
      txtInv(0).text = ""
      txtDueDate.text = ""
      txtUnit(0).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(0).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False

      txtPICNNet.text = "0.00"
      txtPICNVat.text = "0.00"
      txtPICNTotal.text = "0.00"

   Case ComponentMode.newLine
      txtUnit(0).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(0).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False

      txtPICNNet.text = "0.00"
      txtPICNVat.text = "0.00"
      txtPICNTotal.text = "0.00"

   Case ComponentMode.EditLine
      txtUnit(0).text = ""
      txtNC(0).text = ""
      txtNCName.text = ""
      txtDept(0).text = ""
      txtPFName.text = ""
      txtJobNo.text = ""
      txtSchedules.text = ""
      txtDetails_(0).text = ""
      txtNet_(0).text = ""
      txtVat_(0).text = ""
      txtTotal.text = ""
      txtRecoverable(0).text = ""
      chkRecover.Value = False
   End Select
End Sub

Public Sub UpdateTotalPICN()
   Dim i As Integer

   txtPICNNet.text = "0"
   txtPICNVat.text = "0"
   txtPICNTotal.text = "0"

   For i = 1 To flxPI.Rows - 1
      If flxPI.RowHeight(i) > 0 Then
         txtPICNNet.text = Val(txtPICNNet.text) + Val(flxPI.TextMatrix(i, 12))
         txtPICNVat.text = Val(txtPICNVat.text) + Val(flxPI.TextMatrix(i, 14))
         txtPICNTotal.text = Val(txtPICNTotal.text) + Val(flxPI.TextMatrix(i, 15))
      End If
   Next i

   txtPICNNet.text = Format(txtPICNNet.text, "0.00")
   txtPICNVat.text = Format(txtPICNVat.text, "0.00")
   txtPICNTotal.text = Format(txtPICNTotal.text, "0.00")
End Sub

Private Sub LoadNominalCode()
   flxSupplier(0).ColWidth(0) = 1500
   flxSupplier(0).ColWidth(1) = 2700
   flxSupplier(0).ColAlignment = vbLeftJustify

    '~~~ Added by Senthuran~~~ Configuring width and position of labels and search boxes.
   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1.Width = 2600
   lblSearch1.Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(0)

   txtSearch1.Width = 1400
   txtSearch1.Left = 40

   txtSearch2.Width = 2600
   txtSearch2.Left = txtSearch1.Left + flxSupplier(0).ColWidth(0)

   '~~~Added By Senthuran~~~ Code to configuer Label Caption
   lblSearch0(0).Caption = "N/C"
   lblSearch1.Caption = "Name"
   lblSearch2.Visible = False

' Error Handler
   On Error GoTo Error_Handler

   Dim adoConn As New ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString
   
   If frmMMain.IsRibbonVersion Then
      szSQL = "SELECT N.* " & _
              "FROM NominalLedger AS N " & _
              "WHERE N.ClientID = '" & txtClientID.text & "' AND " & _
                    "Posting AND (ISNULL(CAType) OR CAType='') " & _
              "ORDER BY N.Code;"
   Else
      szSQL = "SELECT N.* " & _
              "FROM NominalLedger AS N " & _
              "ORDER BY N.Code;"
   End If

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim iRows As Integer

   flxSupplier(0).Rows = 2
   iRows = 1
   While Not adoRst.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = adoRst.Fields.Item("Code").Value
      flxSupplier(0).TextMatrix(iRows, 1) = adoRst.Fields.Item("Name").Value
      If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRst.MoveNext
   Wend

   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing

   flxSupplier(0).RowHeight(0) = 0

   Exit Sub

' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If szCallerForm = "M" Then
      frmMaintenance.Enabled = True
   End If
   If szCallerForm = "P" Then
      frmPO.Enabled = True
      frmPO.cmdEdit(1).Enabled = True
   End If
   If szCallerForm = "I" Then
      frmPurchaseExpense.Enabled = True
   End If
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
End Sub

Private Sub txtDate_LostFocus()
   On Error Resume Next

   If txtDate.text <> "" Then TextBoxFormatDate txtDate
   If txtDate.text <> "" And bEditMode Then txtDueDate.text = DateAdd("d", iDayTerms, txtDate.text)
End Sub

Private Sub txtDueDate_Change()
   TextBoxChangeDate txtDueDate
End Sub

Private Sub txtDueDate_GotFocus()
   If txtDueDate.text = "dd/mm/yyyy" Then
      txtDueDate.text = ""
      Exit Sub
   End If
   If Len(txtDueDate.text) < 10 Then txtDueDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDueDate
End Sub

Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdClientSerc.SetFocus
    End If
   TextBoxKeyPrsDate txtDueDate, KeyAscii
End Sub

Private Sub txtDueDate_LostFocus()
   If txtDueDate.text <> "" Then TextBoxFormatDate txtDueDate
End Sub

Private Sub txtDate_GotFocus()
   If txtDate.text = "dd/mm/yyyy" Then
      txtDate.text = ""
      Exit Sub
   End If
   If Len(txtDate.text) < 10 Then
      txtDate.text = Format(Date, "dd/mm/yyyy")
'      lblPostingDate.ToolTipText = txtDate.text
   End If
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDueDate.SetFocus
    End If
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtNC_Change(Index As Integer)
'added by anol 06 Jan 2015
'issue 469
'  This change even is required when a supplier has a default NC.
'  System put the value of NC in the text box then this change event
'  look for the NC Name for the name text box.
   Dim i As Integer
   
   flxSupplier(0).Clear
   flxSupplier(0).Rows = 2
   txtNCName.text = ""
   LoadNominalCode
   For i = 1 To flxSupplier(0).Rows - 1
      If UCase(Left(flxSupplier(0).TextMatrix(i, 0), Len(txtNC(0).text))) = UCase(txtNC(0).text) Then
         txtNCName.text = flxSupplier(0).TextMatrix(i, 1)
         Exit For
      End If
   Next i
End Sub

Private Sub txtNet__GotFocus(Index As Integer)
   txtNet_(Index).SelStart = 0
   txtNet_(Index).SelLength = Len(txtNet_(Index).text)
End Sub

Private Sub txtNet__KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
   End If
   If KeyAscii = 13 Then ' Or KeyAscii = 10 Then txtNet__LostFocus (0)
        cmdTaxList(0).SetFocus
   End If
   DigitTextKeyPress txtNet_(0), KeyAscii
End Sub

Private Sub txtNet__LostFocus(Index As Integer)
   txtVat_(0).text = Format(IIf(txtNet_(0).text = "", 0, Val(txtNet_(0).text)) * (nTaxCode / 100), "0.00")
   txtNet_(0).text = Format(txtNet_(0).text, "0.00")
   txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
   txtTotal.text = Format(txtTotal.text, "0.00")
End Sub

Private Sub txtRecoverable_GotFocus(Index As Integer)
   If Index = 0 Then
      SelTxtInCtrl txtRecoverable(0)
'   Else
'      If chkRecover.Enabled Then cmdJobNo(0).SetFocus
   End If
End Sub
'
'Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
'   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, szClientID
'End Sub

Private Sub txtVat__KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
   End If
End Sub

Private Sub txtVat__LostFocus(Index As Integer)
   txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
   txtTotal.text = Format(txtTotal.text, "0.00")
End Sub

