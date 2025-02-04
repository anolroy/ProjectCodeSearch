VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance"
   ClientHeight    =   11490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   22860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   22860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frafixTransactions 
      Caption         =   "fix Transactions"
      Height          =   4200
      Left            =   9900
      TabIndex        =   64
      Top             =   7605
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Frame fraCheckSumPayment 
      Caption         =   "Report"
      Height          =   4200
      Left            =   180
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   8160
      Begin VB.CommandButton cmdFixWestbourne 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Fix Westbourne Pay Allocation Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4455
         Picture         =   "frmHelp.frx":0000
         TabIndex        =   66
         ToolTipText     =   "Click to fix the Data"
         Top             =   2070
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.CommandButton cmdfixEvron 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Fix Everon Receipt Allocation Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4410
         Picture         =   "frmHelp.frx":5D2FA
         TabIndex        =   65
         ToolTipText     =   "Click to fix the Data"
         Top             =   2700
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.CommandButton cmdFixPayAllocation 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Fix WPM Pay Allocation data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4455
         Picture         =   "frmHelp.frx":BA5F4
         TabIndex        =   60
         ToolTipText     =   "Click to fix the Data"
         Top             =   1485
         Visible         =   0   'False
         Width           =   3525
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   2985
         Left            =   180
         TabIndex        =   61
         Top             =   585
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   5265
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
      Begin VB.Shape Shape6 
         Height          =   3435
         Left            =   90
         Top             =   270
         Width           =   5415
      End
      Begin VB.Label Label13 
         Caption         =   "Report checksum Payment allocation for [WPM]"
         Height          =   510
         Left            =   135
         TabIndex        =   63
         Top             =   270
         Width           =   4560
      End
      Begin VB.Label Label12 
         Caption         =   "0"
         Height          =   195
         Left            =   4950
         TabIndex        =   62
         Top             =   450
         Width           =   330
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   "Report"
      Height          =   4155
      Left            =   11970
      TabIndex        =   45
      Top             =   6750
      Visible         =   0   'False
      Width           =   8115
      Begin VB.CommandButton cmdShowReport 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Show Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5760
         Picture         =   "frmHelp.frx":1178EE
         TabIndex        =   47
         ToolTipText     =   "Click to fix the Data"
         Top             =   3555
         Width           =   1590
      End
      Begin VB.CommandButton cmdFixTransactions 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Fix Transactions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3960
         Picture         =   "frmHelp.frx":174BE8
         TabIndex        =   46
         ToolTipText     =   "Click to fix the Data"
         Top             =   3555
         Width           =   1725
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1635
         Left            =   180
         TabIndex        =   48
         Top             =   765
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2884
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1635
         Left            =   2700
         TabIndex        =   49
         Top             =   765
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2884
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   1635
         Left            =   5265
         TabIndex        =   50
         Top             =   765
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   2884
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   90
         TabIndex        =   51
         Top             =   2880
         Visible         =   0   'False
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   240
         Left            =   90
         TabIndex        =   58
         Top             =   2565
         Visible         =   0   'False
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label trxCount 
         Caption         =   "0"
         Height          =   195
         Left            =   2205
         TabIndex        =   57
         Top             =   495
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label6 
         Caption         =   "PI & PC mismatch comparing tlbpurinv and NLposting :"
         Height          =   510
         Left            =   135
         TabIndex        =   56
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label trxCount2 
         Caption         =   "0"
         Height          =   195
         Left            =   4770
         TabIndex        =   55
         Top             =   495
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label trxCount3 
         Caption         =   "0"
         Height          =   195
         Left            =   7335
         TabIndex        =   54
         Top             =   495
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label7 
         Caption         =   "PP & PA mismatch comparing tlbpayment and NLposting:"
         Height          =   555
         Left            =   2655
         TabIndex        =   53
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "SR & SA mismatch comparing tlbReceipt and NLposting:"
         Height          =   420
         Left            =   5310
         TabIndex        =   52
         Top             =   270
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         Height          =   2265
         Left            =   90
         Top             =   225
         Width           =   2400
      End
      Begin VB.Shape Shape2 
         Height          =   2220
         Left            =   2610
         Top             =   225
         Width           =   2490
      End
      Begin VB.Shape Shape3 
         Height          =   2220
         Left            =   5175
         Top             =   225
         Width           =   2490
      End
   End
   Begin VB.Frame fraChecksum 
      Caption         =   "Report"
      Height          =   4200
      Left            =   900
      TabIndex        =   41
      Top             =   8820
      Visible         =   0   'False
      Width           =   8160
      Begin VB.CommandButton cmdIssue673 
         Caption         =   "Issue 673 WPM receipt Allocation fix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4500
         TabIndex        =   67
         Top             =   1890
         Width           =   3480
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   2985
         Left            =   180
         TabIndex        =   42
         Top             =   585
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   5265
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
      Begin VB.Label Label17 
         Caption         =   "0"
         Height          =   195
         Left            =   4950
         TabIndex        =   44
         Top             =   450
         Width           =   330
      End
      Begin VB.Label Label16 
         Caption         =   "Report checksum receipt allocation for [WPM]"
         Height          =   510
         Left            =   135
         TabIndex        =   43
         Top             =   270
         Width           =   4560
      End
      Begin VB.Shape Shape8 
         Height          =   3435
         Left            =   90
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.Frame fraVatamount 
      Height          =   4110
      Left            =   13725
      TabIndex        =   39
      Top             =   2295
      Visible         =   0   'False
      Width           =   8115
      Begin VB.CommandButton cmdFixVat 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Fix Transactions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6030
         Picture         =   "frmHelp.frx":1D1EE2
         TabIndex        =   40
         ToolTipText     =   "Click to fix the Data"
         Top             =   3060
         Width           =   1725
      End
   End
   Begin VB.ComboBox cboIssue 
      Height          =   315
      Left            =   3465
      TabIndex        =   0
      Top             =   1530
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4590
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   10170
      TabIndex        =   30
      Top             =   0
      Width           =   10200
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxIssueList 
         Height          =   4110
         Left            =   45
         TabIndex        =   31
         Top             =   405
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   7250
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
      Begin MSForms.Label Label11 
         Height          =   195
         Left            =   7965
         TabIndex        =   36
         Top             =   90
         Width           =   1545
         VariousPropertyBits=   8388627
         Caption         =   "Issue Number"
         Size            =   "2725;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   195
         Left            =   2025
         TabIndex        =   35
         Top             =   90
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Description"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label9 
         Height          =   195
         Left            =   405
         TabIndex        =   34
         Top             =   90
         Width           =   1545
         VariousPropertyBits=   8388627
         Caption         =   "Date"
         Size            =   "2725;344"
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
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   0
         Left            =   45
         Top             =   75
         Width           =   10035
      End
   End
   Begin VB.CommandButton cmdIssue681 
      Caption         =   "Issue 681"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10530
      TabIndex        =   29
      Top             =   315
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   10755
      ScaleHeight     =   4200
      ScaleWidth      =   6255
      TabIndex        =   18
      Top             =   945
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
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3525
         Left            =   45
         TabIndex        =   20
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6218
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
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
   Begin VB.Frame fraManualDemand 
      Height          =   4065
      Left            =   13995
      TabIndex        =   5
      Top             =   5625
      Visible         =   0   'False
      Width           =   8115
      Begin VB.CommandButton cmdConvert 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6120
         Picture         =   "frmHelp.frx":22F1DC
         TabIndex        =   38
         ToolTipText     =   "Click to fix the Data"
         Top             =   3330
         Width           =   1725
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4545
         TabIndex        =   4
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtStart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   3
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "End demand No"
         Height          =   240
         Left            =   3240
         TabIndex        =   7
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Start demand No"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdUnlock 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Unlock Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13680
      TabIndex        =   1
      ToolTipText     =   "Supplier"
      Top             =   315
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00E9E8E4&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16965
      TabIndex        =   2
      ToolTipText     =   "Supplier"
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Frame FraReconciliation 
      Height          =   4065
      Left            =   5625
      TabIndex        =   9
      Top             =   5220
      Visible         =   0   'False
      Width           =   8070
      Begin VB.CheckBox chkRollBackCheck3 
         Caption         =   "I have printed/saved the latest Reconciled transactions Bank reconciliation report."
         Height          =   420
         Left            =   405
         TabIndex        =   70
         Top             =   2115
         Width           =   7395
      End
      Begin VB.CheckBox chkRollBackCheck2 
         Caption         =   "I have printed/saved the latest unreconciled transactions Bank reconciliation report."
         Height          =   420
         Left            =   405
         TabIndex        =   69
         Top             =   1800
         Width           =   7395
      End
      Begin VB.CheckBox chkRollBackCheck1 
         Caption         =   "I have backed up my Prestige data."
         Height          =   420
         Left            =   405
         TabIndex        =   68
         Top             =   1485
         Width           =   7395
      End
      Begin VB.CommandButton cmdRollbackBreconciliation 
         BackColor       =   &H00E9E8E4&
         Caption         =   "Rollback Bank Reconciliation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3195
         Picture         =   "frmHelp.frx":28C4D6
         TabIndex        =   37
         ToolTipText     =   "Click to fix the Data"
         Top             =   2700
         Width           =   3705
      End
      Begin VB.CommandButton cmdBC 
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
         Left            =   6705
         TabIndex        =   15
         Top             =   315
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
         Left            =   3510
         TabIndex        =   13
         Top             =   315
         Width           =   300
      End
      Begin MSForms.TextBox txtPassword 
         Height          =   285
         Left            =   4815
         TabIndex        =   28
         Top             =   855
         Width           =   1890
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "3334;503"
         PasswordChar    =   42
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   17
         Top             =   315
         Width           =   735
      End
      Begin MSForms.TextBox txtBC 
         Height          =   285
         Left            =   4815
         TabIndex        =   16
         Top             =   315
         Width           =   1890
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "3334;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1035
         TabIndex        =   14
         Top             =   315
         Width           =   2475
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4366;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         Caption         =   "Client"
         Height          =   240
         Left            =   495
         TabIndex        =   12
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Reconciliation Date"
         Height          =   240
         Left            =   495
         TabIndex        =   11
         Top             =   855
         Width           =   1455
      End
      Begin MSForms.ComboBox cboCurStDt 
         Height          =   315
         Left            =   2025
         TabIndex        =   10
         Top             =   810
         Width           =   1770
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3122;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Shape Shape5 
      Height          =   4155
      Left            =   10305
      Top             =   45
      Width           =   8205
   End
   Begin VB.Label Label3 
      Caption         =   "Issue"
      Height          =   240
      Left            =   540
      TabIndex        =   8
      Top             =   315
      Width           =   915
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTextBox As String
'Some principles of this software
'In the table Paytransaction ToTran is Invoice
'                    FromTran is payment
'*single payment can fill many invoices in PI
Private Const IDC_APPSTARTING = 32650&
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Const IDC_CROSS = 32515&
Private Const IDC_IBEAM = 32513&
Private Const IDC_ICON = 32641&
Private Const IDC_NO = 32648&
Private Const IDC_SIZE = 32640&
Private Const IDC_SIZEALL = 32646&
Private Const IDC_SIZENESW = 32643&
Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZENWSE = 32642&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_UPARROW = 32516&
Private Const IDC_WAIT = 32514&

Private Declare Function LoadCursorLong Lib "user32" Alias "LoadCursorA" _
  (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Declare Function SetCursor Lib "user32" _
  (ByVal hCursor As Long) As Long
Private Sub loadflxIssueList()
        Dim iRow As Integer
        iRow = 1
         flxIssueList.RowHeight(0) = 0
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Convert Demand Manual to Auto [WPM]"
         flxIssueList.TextMatrix(iRow, 3) = "" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
        
         flxIssueList.TextMatrix(iRow, 1) = "" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "VAT amount in demandsplit is not in total Amount [WESTGATE]"
         flxIssueList.TextMatrix(iRow, 3) = "" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
        
         flxIssueList.TextMatrix(iRow, 1) = "" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Rollback Bank reconciliation"
         flxIssueList.TextMatrix(iRow, 3) = "" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
        
         flxIssueList.TextMatrix(iRow, 1) = "2017/09/01" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Payment Edit, Receipt Edit was not posting Data to NL FIX"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 452" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
'         flxIssueList.TextMatrix(5, 0) = "" 'for selection
'         flxIssueList.TextMatrix(5, 1) = "" 'issue Number
'         flxIssueList.TextMatrix(5, 2) = "" 'Date
'         flxIssueList.TextMatrix(5, 3) = "Booking Batch Receipt with lessee filter not updating lessee balance correctly-filtering on lesse created a problem"
'         flxIssueList.TextMatrix(5, 4) = "" 'Empty
'         flxIssueList.RowHeight(5) = 500
'         iRow = iRow + 1
         
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Clear Record Locking table"
         flxIssueList.TextMatrix(iRow, 3) = " " 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2018/11/29" 'Date
'         flxIssueList.TextMatrix(iRow, 2) = "Report checksum receipt allocation for current database(fix button is for [WPM],it is partially hardcoded)"
            flxIssueList.TextMatrix(iRow, 2) = "Report checksum receipt allocation for current database"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 673" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
                    
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2018/11/13" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Duplicate Invoice PI5221 [WPM]"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 681" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2019/05/23" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Report checksum Payment allocation for current database (fix button is hardcoded for [WPM])"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 769" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2018/12/06" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Report checksum Payment allocation for AC(DB location and fix button hardcoded)"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 696" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2018/12/06" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Receipt allocation for AC(DB location and fix button hardcoded)"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 696" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2019/05/10" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Batch Receipts not Bank Reconciling"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 764" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         flxIssueList.RowHeight(iRow) = 0
         iRow = iRow + 1
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2019/06/10" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Inconsistency in Bank Reconciliation ‘ORCHGREEN’"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 784" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/04/18" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "WPM: The Lessee WR151A, Transaction SR62072.OSamount Incorrect"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/04/18" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         
          flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/04/20" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Client statement: Paid to client isnt marked with statement ID"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/04/20" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/04/20" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "WPM: The Lessee GEO6A, Transaction SR60665.OSamount Incorrect"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/04/20" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/05/10" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "WPM: Some transactions are not posted to NLPOSTING, post them"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/05/10" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/05/15" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "WPM: fix inconsistent data ob Batch Payment"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/05/15" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         
         flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/05/26" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "MNC: Update ALL Retention Bankcode from Client statement"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/05/26" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
          flxIssueList.AddItem ""
         flxIssueList.TextMatrix(iRow, 0) = "" 'for selection
         flxIssueList.TextMatrix(iRow, 1) = "2023/08/23" 'Date
         flxIssueList.TextMatrix(iRow, 2) = "Clear all locks"
         flxIssueList.TextMatrix(iRow, 3) = "Issue 2023/05/23" 'issue Number
         flxIssueList.TextMatrix(iRow, 4) = "" 'Empty
         iRow = iRow + 1
         
         

'
                    
'        cboIssue.AddItem "Convert Demand Manual to Auto [WPM]"
'        cboIssue.AddItem "VAT amount in demandsplit is not in total Amount [WESTGATE]"
'        cboIssue.AddItem "Rollback Bank reconciliation"
'        cboIssue.AddItem "Payment Edit, Receipt Edit was not posting Data to NL FIX"
'        cboIssue.AddItem "Booking Batch Receipt with lessee filter not updating lessee balance correctly"
'        cboIssue.AddItem "Show checksum report of receipt allocation"
End Sub
Private Function SetMouseCursor(CursorType As Long)
  Dim hCursor As Long
  hCursor = LoadCursorLong(0&, CursorType)
  hCursor = SetCursor(hCursor)
End Function
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

Private Sub cmdIssue_Click()
    
End Sub

Private Sub chkRollBackCheck1_Click()
    If chkRollBackCheck1.Value = 1 And chkRollBackCheck2.Value = 1 And chkRollBackCheck3.Value = 1 Then
        cmdRollbackBreconciliation.Enabled = True
    Else
        cmdRollbackBreconciliation.Enabled = False
    End If
End Sub

Private Sub chkRollBackCheck2_Click()
    If chkRollBackCheck1.Value = 1 And chkRollBackCheck2.Value = 1 And chkRollBackCheck3.Value = 1 Then
        cmdRollbackBreconciliation.Enabled = True
    Else
        cmdRollbackBreconciliation.Enabled = False
    End If
End Sub

Private Sub chkRollBackCheck3_Click()
    If chkRollBackCheck1.Value = 1 And chkRollBackCheck2.Value = 1 And chkRollBackCheck3.Value = 1 Then
        cmdRollbackBreconciliation.Enabled = True
    Else
        cmdRollbackBreconciliation.Enabled = False
    End If
End Sub

Private Sub cmdConvert_Click()
       Dim adoconn As New ADODB.Connection
     'For WPM demand Manual to Automatic
            If Trim(txtStart.text) = "" Or Trim(txtEnd.text) = "" Then
                    MsgBox "PLease enter the range correcly", vbInformation, "Warning"
                    Exit Sub
            End If
            adoconn.Open getConnectionString
            If MsgBox("Do you wish to convert WPM Demand Manual to Auto from " & Trim(txtStart.text) & " to " & Trim(txtEnd.text) & " ?", vbYesNo, "Warning") = vbYes Then
                 adoconn.Execute "update DemandRecords INNER JOIN DemandSplitRecords ON DemandRecords.DemandID = DemandSplitRecords.DemandID " & _
                 "set A_M='A' WHERE DemandRecords.dmDslno >=" & Trim(txtStart.text) & " and DemandRecords.dmDslno <=" & Trim(txtEnd.text) & "  "
                 'Freqency ID is not writing for some Demand split records
                 'Fixed by anol on 03 Jun 2016
                 adoconn.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LServiceCharges LS SET DS.FrequencyID=LS.SCFrequency,DS.ChargingMethod=LS.ChargingMethod " & _
                 "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=2 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
                 "AND isnull(LS.Spare3)"
                 adoconn.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LRentCharges LS SET DS.FrequencyID=LS.BRFrequency,DS.ChargingMethod=LS.spare1 " & _
                 "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=1 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
                 "AND isnull(LS.Spare3)"
                  adoconn.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LInsuranceCharges LS SET DS.FrequencyID=LS.InsuranceFrequency,DS.ChargingMethod=LS.ChargingType " & _
                 "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=3 AND L.SageAccountNumber=DR.SageAccountNumber AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
                 "AND isnull(LS.Spare3)"
            End If
            adoconn.Close
            Set adoconn = Nothing
End Sub

Private Sub cmdfixEvron_Click()
    Exit Sub
    Dim adoconn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    
    adoconn.Open "DSN=PrestigeBMcNS002;UID=;PWD=" & accessDBPws & ""  'PrestigeBMcNS002
    'adoConn.Execute "Update tlbReceipt set Osamount=500,ReceiptView=true where transactionId =1848"
    'adoConn.Execute "Update tlbReceiptsplit set Osamount=500 where transactionId ='1406111111320135972'"
    adoRst.Open "Select * from RptTransactions", adoconn, adOpenDynamic, adLockOptimistic
'    If adorst.EOF Then
        adoRst.AddNew
        adoRst!TranType = "AL"
        adoRst!TransactionID = MaxAllocID(adoconn)
        adoRst!Alloc_Unalloc = 1
        adoRst!FromTran = 1904
        adoRst!ToTran = 1848
        adoRst!AllocDate = "30 June 2014"
        adoRst!receiptAmount = 500
        adoRst!Discount = 0
        adoRst!IsSageUpdate = False
        adoRst!UpdateSage = False
        adoRst!Unalloc = 0
        adoRst!BankCode = "1200"
        adoRst!nominalCode = "1200"
        adoRst!SlNumber = 888
        adoRst!Reconciled = 0
        adoRst!ReconNow = ""
'        adorst!Exp2Sage
'        adorst!FundID
        adoRst.Update
        MsgBox "Transactions has been fixed for Everon"
'    Else
'        MsgBox "record already exists (Everon)"
'     End If
    adoRst.Close
    adoconn.Close
    Set adoconn = Nothing
    
End Sub
Private Function MaxAllocID(adoconn) As Integer
     Dim adorst2 As New ADODB.Recordset
     adorst2.Open "Select max(TransactionID)+1 as TranID from RptTransactions", adoconn, adOpenKeyset, adLockReadOnly
     If Not adorst2.EOF Then
            MaxAllocID = adorst2("tranID").Value
     End If
     
End Function
Private Sub cmdFixPayAllocation_Click()
   If MsgBox("Paytransactions HEATH(PI7150) will be fixed if you run this procedure. are you sure?", vbYesNo, "Confirm?") = vbYes Then
   
        Dim adoconn As New ADODB.Connection
        'adoconn.Open "DSN=PrestigeBMcNS001;UID=;PWD=" & accessDBPws & ""   'PrestigeBMcNS017
        adoconn.Open getConnectionString
        adoconn.Execute "Delete from Paytransactions where transactionId=10723"
        adoconn.Execute "update Paytransactions set NominalCode='1210', BankCode='1210' where transactionId=10722"
        adoconn.Execute "update tlbPayment set osamount=0 where transactionID=13266"
        
       
            
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Paytransactions HEATH(PI7150) has been fixed for WPM."
        MSHFlexGrid5.Clear
        MSHFlexGrid5.Rows = 2
        Label12.Caption = "0"
        cmdFixPayAllocation.Visible = False
        
   'JGC(PI6390)  Fix
'        Dim adoConn As New ADODB.Connection
'        'adoconn.Open "DSN=PrestigeBMcNS001;UID=;PWD=" & accessDBPws & ""   'PrestigeBMcNS017
'        adoConn.Open getConnectionString
'        adoConn.Execute "Delete from Paytransactions where transactionId=9779"
'        adoConn.Execute "update Paytransactions set NominalCode='1210', BankCode='1210' where transactionId=9778"
'        adoConn.Execute "update tlbPayment set osamount=0 where transactionID=11817"
'
'        'adoConn.Execute "Delete from Paytransactions where transactionId in (5683,2964,6829,7669,7670,8691)"
'        'adoConn.Execute "Update tlbpayment set OSamount=0 where TransactionID in (1209,1448,4664,12446,13810)" 'tlbpaymentsplit is fiyne osamount=0
'        adoConn.Close
'        Set adoConn = Nothing
'        MsgBox "Paytransactions JGC(PI6390) has been fixed for WPM."
'        MSHFlexGrid5.Clear
'        MSHFlexGrid5.Rows = 2
'        Label12.Caption = "0"
'        cmdFixPayAllocation.Visible = False
    End If
End Sub

Private Sub cmdFixVat_Click()
        Exit Sub
        Dim adoconn As New ADODB.Connection
        
        adoconn.Open getConnectionString
        Dim rsReceipt As New ADODB.Recordset
        If MsgBox("Do you wish to run WESTGATE VAT Problem fix routine?", vbYesNo, "Warning") = vbYes Then
              adoconn.Execute "Update DemandSplitRecords SET TotalAmount=Amount+VATAmount Where A_M='B'"
             Dim rsAllNLamt As New ADODB.Recordset
             Dim amt As Double
              rsReceipt.Open "SELECT cstr(R.TransactionID) AS TRID,demandref  FROM tlbReceipt AS R, (SELECT D.DemandID,  SUM(S.TotalAmount) AS DT FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID GROUP BY D.DemandID) AS S WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND ROUND(R.Amount, 4) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 4)", adoconn, adOpenKeyset, adLockReadOnly
               While Not rsReceipt.EOF
                     adoconn.Execute "Update tlbReceiptsplit,DemandSplitRecords,tlbReceipt  SET  tlbReceiptsplit.Amount=DemandSplitRecords.TotalAmount,tlbReceiptsplit.OSAmount=DemandSplitRecords.TotalAmount, tlbReceipt.Amount=DemandSplitRecords.TotalAmount, tlbReceipt.OSAmount=DemandSplitRecords.TotalAmount  where tlbReceipt.DemandRef=DemandSplitRecords.DemandID AND tlbReceipt.TransactionID =tlbReceiptsplit.rptHeader AND rptHeader =" & rsReceipt.Fields("TRID").Value & ";"
                     rsAllNLamt.Open "Select sum(AMOUNT)as amt from NLPOSTING where TRANS_ID='" & rsReceipt.Fields("demandref").Value & "' AND NOMINAL_CODE<>'LL4000' and Deleteflag=0", adoconn, adOpenKeyset, adLockReadOnly
                     amt = -rsAllNLamt.Fields(0).Value
                     rsAllNLamt.Close
                     adoconn.Execute "update NLPOSTING set AMOUNT =" & amt & " where TRANS_ID='" & rsReceipt.Fields("demandref").Value & "' AND NOMINAL_CODE='LL4000' and Deleteflag=0"
                     amt = 0
                     rsReceipt.MoveNext
              Wend
              'UpdateNL adoConn
               rsReceipt.Close
        End If
        adoconn.Close
        Set adoconn = Nothing
End Sub

Private Sub cmdFixWestbourne_Click()
    Exit Sub
    Dim adoconn As New ADODB.Connection
    adoconn.Open "DSN=PrestigeBMcNS017;UID=;PWD=" & accessDBPws & ""  'PrestigeBMcNS017
    adoconn.Execute "Delete from Paytransactions where transactionId in (252,253,254)"
    adoconn.Close
    Set adoconn = Nothing
    MsgBox "Transactions has been fixed for Westbourne"
End Sub

Private Sub cmdIssue681_Click()
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String
    Dim lSlNumber As Long
    adoconn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    'Type 6 Need to find the maximum SLnumber to be inputted
    'update tblPurIv Where slnumber=5221 AND SUPP_AC='AXIS'and transactiontype=6
    'Update tlbpayment slnumber=newslnumber where slnumber=5221 AND SUPP_AC='AXIS'and type=6
    'Update NLposting for that transaction with new Slnumber
    'allocation is fyne we are not doing anything with the transaction number
    'adoRst.Open "Select * from NLPOSTING where ACCOUNT_NUMBER='AXIS' and TRANSACTION_TYPE=6 AND TRANSACTION_REF='5221'", adoConn, adOpenStatic, adLockReadOnly
    adoRst.Open "select count(slnumber) as CNT from tblPurInv  Where slnumber=5221 and TransactionType=6", adoconn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
        If adoRst("CNT").Value > 1 Then
            lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
            adoconn.Execute "update tblPurInv set slnumber=" & lSlNumber & " Where slnumber=5221 AND SUPP_AC='AXIS'and TransactionType=6"
            adoconn.Execute "Update tlbpayment set slnumber=" & lSlNumber & "  where slnumber=5221 AND sageaccountnumber='AXIS'and type=6"
            adoconn.Execute "Update NLPOSTING set TRANSACTION_REF='" & lSlNumber & "' AND TRANS_ID='" & lSlNumber & "' where ACCOUNT_NUMBER='AXIS' and TRANSACTION_TYPE=6 AND TRANSACTION_REF='5221'"
             MsgBox "Duplicate PI5221 has been fixed "
         Else
             MsgBox "Duplicate serial for PI5221 not found"
         End If
    Else
         MsgBox "Duplicate serial for PI5221 not found"
    End If
    adoRst.Close
    Set adoRst = Nothing
    adoconn.Close
    Set adoconn = Nothing
   
End Sub

Private Sub cmdIssue673_Click()
    Exit Sub
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String
    Dim lSlNumber As Long
    adoconn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
    Dim i As Integer
    'if Osamount is zero and there is entry in the allocation table and that is not equal to receipt amount then you need to unallocate first becasue there may have bank reconciliation
    ' and then you  need to update the OSamount manually.
    'case 1 for which have entry in the allocation
    adoRst.Open " Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,Allocamt from tlbReceipt R,(select Sum(ReceiptAmount) as Allocamt," & _
            "ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-Allocamt),2)<>round(osamount,2) ", adoconn, adOpenKeyset, adLockReadOnly
    
    While Not adoRst.EOF
        If adoRst("osamount").Value >= 0 And (adoRst("amount").Value - adoRst("Allocamt").Value) >= 0 Then '2nd clause is for prevent overallocation
            adoconn.Execute "Update tlbReceipt Set OSAmount =" & (adoRst("amount").Value - adoRst("Allocamt").Value) & " where TransactionID=" & adoRst("transactionID").Value & ""
            adoconn.Execute "Update tlbReceiptsplit Set OSAmount = " & (adoRst("amount").Value - adoRst("Allocamt").Value) & " where rptheader=" & adoRst("transactionID").Value & ""
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
        End If
        adoRst.MoveNext
    Wend
    adoRst.Close
    Set adoRst = Nothing
    'case 2 for which records does not have entry in the allocation but osamount>amount
    adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
  'introduce this trap while closing form demand and batch receipt
   'I am putting following execution Here so that it runs only once validity is only 2018, because later on if he runs this procedure by mistake
    'it shall not run because of this date validity
    
   ' If Not adorst.EOF Then 'And Year(Date) = 2018
        'this is the receipt and unallocated so OSamount=amount
        adoconn.Execute "Update tlbReceipt set OSamount=amount,receiptview=true where transactionID=16969"
        adoconn.Execute "Update tlbReceiptSplit set OSamount=amount where rptHeader=16969"
        adoconn.Execute "Delete from RptTransactions where transactionID=13906"
        szTran2Fix = szTran2Fix + "SI1782"
        i = i + 1
        '17429 buck04a

        adoconn.Execute "Update tlbReceipt set OSamount=amount,receiptview=true where transactionID=17429"
        adoconn.Execute "Update tlbReceiptSplit set OSamount=amount where rptHeader=17429"
        adoconn.Execute "Delete from RptTransactions where transactionID=14458"
         szTran2Fix = szTran2Fix + "SI2617"
         i = i + 1
         'BUCK01A

        adoconn.Execute "Update tlbReceipt set OSamount=amount,receiptview=true where transactionID=25885"
        adoconn.Execute "Update tlbReceiptSplit set OSamount=amount where rptHeader=25885"
        adoconn.Execute "Delete from RptTransactions where transactionID=21830"
        szTran2Fix = szTran2Fix + "SI2614"
         i = i + 1
   ' End If
    
    While Not adoRst.EOF
        adoconn.Execute "Update tlbReceipt Set OSAmount = " & adoRst("amount").Value & " where TransactionID=" & adoRst("transactionID").Value & ""
        adoconn.Execute "Update tlbReceiptsplit Set OSAmount = " & adoRst("amount").Value & " where rptheader=" & adoRst("transactionID").Value & ""
        szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
        adoRst.MoveNext
        i = i + 1
    Wend
    adoRst.Close
    Set adoRst = Nothing
    
    MsgBox "issue 673 has been fixed.Transactions are: " & Chr(13) & szTran2Fix & Chr(13) & " Total: " & i & " Records"
   
    
    adoconn.Close
    Set adoconn = Nothing
   
End Sub

Private Sub cmdRollbackBreconciliation_Click()
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    'Written by anoll 21 July 2016
        If txtPassword.text <> "xx" Then
            MsgBox "Password is incorrect. Please enter a valid password"
            Exit Sub
        End If
        
        If IsDate(cboCurStDt.Value) = False Then
            MsgBox "Date is invalid"
            Exit Sub
        End If
        If MsgBox("Do you wish to rollback all reconciled bank reconciliations back to and including the following date : " & cboCurStDt.Value & "? ", vbYesNo, "Warning") = vbYes Then
            adoconn.Execute "DELETE FROM tlbBankReconClosingBal where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & "' AND StatementDate>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "#"
            adoconn.Execute "DELETE FROM tlbBankReconcilation where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & "' AND ReconDate>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "# "
'            adoconn.Execute "Update tlbReceipt SET ReconNow= NULL,Reconciled= NULL where tlbReceipt.ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
'                "' AND  CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "# AND Reconnow is not NULL"
                
                'modified by anol 20170117 some of the reconciliation where clinet ID was null  was not updating.
              adoconn.Execute "Update tlbReceipt,Units,Property SET tlbReceipt.ReconNow= NULL,tlbReceipt.Reconciled= NULL where tlbReceipt.UnitID=Units.UnitNumber AND Units.PropertyID=Property.PropertyID and Property.ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
                "' AND  CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "# AND Reconnow is not NULL"
   
               
            adoconn.Execute "Update tlbPayment SET ReconNow= NULL, Reconciled = NULL where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & _
            "' AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "#"
            adoconn.Execute "Update tlbBankPayment SET ReconNow= NULL,Reconciled=NULL where ClientID='" & txtClientList.Tag & "' AND Bank_AC='" & txtBC.Tag & _
            "' AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(cboCurStDt.Value, "dd MMM yyyy") & "#"
            adoRst.Open "Select *  FROM tlbBankReconClosingBal where ClientID='" & txtClientList.Tag & "' AND BankCode='" & txtBC.Tag & "' order by StatementDate ", adoconn, adOpenKeyset, adLockReadOnly
            If adoRst.EOF Then
                  adoconn.Execute "Update tlbClientBanks set spare2=NULL , ClosingBal=0, SOB=0,PCB=0  where Client_ID='" & txtClientList.Tag & "' AND NominalCode='" & txtBC.Tag & "'"
            Else
                  adoRst.MoveLast
                  adoconn.Execute "Update tlbClientBanks set spare2='" & Format(adoRst.Fields("StatementDate").Value, "dd MMM yyyy") & "' , SOB=" & adoRst.Fields("ProjClbal").Value & ",PCB=0,ClosingBal=" & adoRst.Fields("ProjClbal").Value & " where CLIENT_ID='" & txtClientList.Tag & "' AND NominalCode='" & txtBC.Tag & "'"
                  ',PCB=" & adoRst.Fields("ProjClbal").Value & "
            End If
            MsgBox "Routine completed successfully."
        End If
     adoconn.Close
     Set adoconn = Nothing
End Sub

Private Sub cmdShowCheckSumreport_Click()
    Call ChecksumValidationOnAllocation
End Sub

Private Sub cmdShowReport_Click()
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Cols = 2
    MSHFlexGrid1.ColWidth(0) = 200
    MSHFlexGrid1.ColWidth(1) = 1800
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.RowHeight(0) = 0
    
    
    MSHFlexGrid2.Clear
    MSHFlexGrid2.Cols = 2
    MSHFlexGrid2.ColWidth(0) = 200
    MSHFlexGrid2.ColWidth(1) = 1800
    MSHFlexGrid2.Rows = 2
    MSHFlexGrid2.RowHeight(0) = 0
    
    MSHFlexGrid3.Clear
    MSHFlexGrid3.Cols = 2
    fraReport.Visible = True
    MSHFlexGrid3.ColWidth(0) = 200
    MSHFlexGrid3.ColWidth(1) = 1800
    MSHFlexGrid3.Rows = 2
    MSHFlexGrid3.RowHeight(0) = 0
    'MSHFlexGrid1.AddItem "1", 2
    Call FIX_PaymentEdit_ReceiptEdit_Report(getConnectionString)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub flxIssueList_Click()
    Dim iRow As Integer
    Dim adodata As New ADODB.Recordset
    Dim i As Integer
    Dim iSelRow As Integer
    Dim adocon As New ADODB.Connection
    Dim adoconn As New ADODB.Connection
    Call SelectOnly1RowFlxGrid(flxIssueList, flxIssueList.row, 0)
    For iRow = 1 To flxIssueList.Rows - 1
         If flxIssueList.TextMatrix(iRow, 0) = "X" Then
            iSelRow = iRow
         End If
    Next iRow
    
    If iSelRow = 1 Then
        fraManualDemand.Visible = True
    Else
        fraManualDemand.Visible = False
    End If
     If iSelRow = 2 Then
        fraVatamount.Visible = True
    Else
        fraVatamount.Visible = False
    End If
    If iSelRow = 3 Then
        FraReconciliation.Visible = True
        'cmdFixTransactions.Caption = "RollBack"
    Else
        FraReconciliation.Visible = False
        'cmdFixTransactions.Caption = "Fix it"
    End If
    
    If iSelRow = 4 Then
         fraReport.Visible = True
         'Call ChecksumValidationOnAllocation
     Else
        fraReport.Visible = False
     End If
     If iSelRow = 5 Then
        cmdUnlock.Visible = True
     Else
        cmdUnlock.Visible = False
     End If
     If iSelRow = 6 Then
         fraChecksum.Visible = True
         Call ChecksumValidationOnAllocation 'RECEIPT PART
     Else
         fraChecksum.Visible = False
     End If
     If iSelRow = 7 Then
         cmdIssue681.Visible = True
     Else
         cmdIssue681.Visible = False
     End If
     If iSelRow = 8 Then
         fraCheckSumPayment.Visible = True
         Call ChecksumValidationOnPayAllocation 'PAYMENT PART
     ElseIf iSelRow = 9 Then
         fraCheckSumPayment.Visible = True
         'On Error Resume Next
         adoconn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=PBMControl.mdb;DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Backup _20190115;Uid=;Pwd=RDSWKDPP;"
         'adoConn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
         adodata.Open "Select * from databases", adoconn, adOpenStatic, adLockReadOnly
             i = 1
           While Not adodata.EOF
               MSHFlexGrid5.AddItem ""
               'MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
               MSHFlexGrid5.TextMatrix(i, 1) = adodata("SCName").Value
               i = i + 1
               adocon.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & adodata("dbName").Value & ";DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Backup _20190115;Uid=;Pwd=RDSWKDPP;"
               'adocon.Open "DSN=" & adodata("AccessDsn").Value & ";UID=;PWD=" & accessDBPws & ""
               Call ChecksumValidationOnPayAllocationAC(adocon, i)
               adodata.MoveNext
           Wend
     Else
         fraCheckSumPayment.Visible = False
     End If
'     If iSelRow = 10 Then
'          frafixTransactions.Visible = True
'     Else
'          frafixTransactions.Visible = False
'     End If
     If iSelRow = 10 Then
         fraChecksum.Visible = True
         'On Error Resume Next
         adoconn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=PBMControl.mdb;DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Backup _20190115;Uid=;Pwd=RDSWKDPP;"
         'adoConn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
         adodata.Open "Select * from databases", adoconn, adOpenStatic, adLockReadOnly
             i = 1
           While Not adodata.EOF
                MSHFlexGrid4.Visible = True
               'MSHFlexGrid5.AddItem ""
               'MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
              ' MSHFlexGrid5.TextMatrix(i, 1) = adodata("SCName").Value
               i = i + 1
               adocon.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & adodata("dbName").Value & ";DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Data 20181206;Uid=;Pwd=RDSWKDPP;"
               'adocon.Open "DSN=" & adodata("AccessDsn").Value & ";UID=;PWD=" & accessDBPws & ""
               Call ChecksumValidationOnAllocationAC(adocon, i)
               Debug.Print adodata.RecordCount
               adodata.MoveNext
           Wend
     Else
         fraChecksum.Visible = False
     End If
     If iSelRow = 11 Then 'Batch Receipts not Bank Reconciling issue 764 date 2019/05/10
        If MsgBox("Do you want to run the routine to fix 'Batch Receipts not Bank Reconciling issue for'OAKTREE' and also rollback Bank reconciliation on '14/11/2018'?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update tlbreceipt set Reconnow='11/12/2017#Full',Reconciled=   90.00   where TransactionID=   16488   "
        adoconn.Execute "Update tlbreceipt set Reconnow='12/09/2017#Full',Reconciled=   90.00   where TransactionID=   13259   "
        adoconn.Execute "Update tlbreceipt set Reconnow='12/09/2017#Full',Reconciled=   90.00   where TransactionID=   13261   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/10/2017#Full',Reconciled=   90.00   where TransactionID=   14094   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/10/2017#Full',Reconciled=   90.00   where TransactionID=   14095   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/10/2017#Full',Reconciled=   90.00   where TransactionID=   14096   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/10/2017#Full',Reconciled=   90.00   where TransactionID=   14098   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   90.00   where TransactionID=   14842   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   2025.00    where TransactionID=   14843   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   2025.00    where TransactionID=   14844   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   90.00   where TransactionID=   14845   "
        adoconn.Execute "Update tlbreceipt set Reconnow='02/03/2018#Full',Reconciled=   90.00   where TransactionID=   17890   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   90.00   where TransactionID=   14848   "
        adoconn.Execute "Update tlbreceipt set Reconnow='24/08/2017#Full',Reconciled=   90.00   where TransactionID=   13253   "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/12/2017#Full',Reconciled=   90.00   where TransactionID=   16490   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2017#Full',Reconciled=   90.00   where TransactionID=   9722    "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/12/2017#Full',Reconciled=   90.00   where TransactionID=   16492   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2018#Full',Reconciled=   90.00   where TransactionID=   16502   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2018#Full',Reconciled=   90.00   where TransactionID=   16503   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2018#Full',Reconciled=   90.00   where TransactionID=   16505   "
        adoconn.Execute "Update tlbreceipt set Reconnow='07/02/2018#Full',Reconciled=   90.00   where TransactionID=   17885   "
        adoconn.Execute "Update tlbreceipt set Reconnow='07/02/2018#Full',Reconciled=   90.00   where TransactionID=   17887   "
        adoconn.Execute "Update tlbreceipt set Reconnow='07/02/2018#Full',Reconciled=   90.00   where TransactionID=   17889   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/11/2017#Full',Reconciled=   90.00   where TransactionID=   14847   "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/05/2017#Full',Reconciled=   90.00   where TransactionID=   9742    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2017#Full',Reconciled=   90.00   where TransactionID=   9723    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2017#Full',Reconciled=   90.00   where TransactionID=   9724    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2017#Full',Reconciled=   90.00   where TransactionID=   9726    "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/02/2017#Full',Reconciled=   90.00   where TransactionID=   9727    "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/02/2017#Full',Reconciled=   90.00   where TransactionID=   9729    "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/02/2017#Full',Reconciled=   90.00   where TransactionID=   9731    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/03/2017#Full',Reconciled=   90.00   where TransactionID=   9732    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/03/2017#Full',Reconciled=   90.00   where TransactionID=   9734    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/03/2017#Full',Reconciled=   90.00   where TransactionID=   9736    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/04/2017#Full',Reconciled=   90.00   where TransactionID=   9737    "
        adoconn.Execute "Update tlbreceipt set Reconnow='12/09/2017#Full',Reconciled=   90.00   where TransactionID=   13257   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/04/2017#Full',Reconciled=   90.00   where TransactionID=   9739    "
        adoconn.Execute "Update tlbreceipt set Reconnow='24/08/2017#Full',Reconciled=   90.00   where TransactionID=   13256   "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/05/2017#Full',Reconciled=   90.00   where TransactionID=   9743    "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/05/2017#Full',Reconciled=   90.00   where TransactionID=   9744    "
        adoconn.Execute "Update tlbreceipt set Reconnow='11/05/2017#Full',Reconciled=   90.00   where TransactionID=   9745    "
        adoconn.Execute "Update tlbreceipt set Reconnow='14/06/2017#Full',Reconciled=   90.00   where TransactionID=   11323   "
        adoconn.Execute "Update tlbreceipt set Reconnow='14/06/2017#Full',Reconciled=   90.00   where TransactionID=   11325   "
        adoconn.Execute "Update tlbreceipt set Reconnow='14/06/2017#Full',Reconciled=   90.00   where TransactionID=   11327   "
        adoconn.Execute "Update tlbreceipt set Reconnow='19/07/2017#Full',Reconciled=   5190.00    where TransactionID=   13244   "
        adoconn.Execute "Update tlbreceipt set Reconnow='19/07/2017#Full',Reconciled=   90.00   where TransactionID=   13245   "
        adoconn.Execute "Update tlbreceipt set Reconnow='19/07/2017#Full',Reconciled=   90.00   where TransactionID=   13248   "
        adoconn.Execute "Update tlbreceipt set Reconnow='24/08/2017#Full',Reconciled=   90.00   where TransactionID=   13251   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/01/2018#Full',Reconciled=   90.00   where TransactionID=   16501   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/04/2017#Full',Reconciled=   90.00   where TransactionID=   9738    "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/09/2018#Full',Reconciled=   90.00   where TransactionID=   24220   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/07/2018#Full',Reconciled=   90.00   where TransactionID=   21626   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/07/2018#Full',Reconciled=   90.00   where TransactionID=   21627   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/07/2018#Full',Reconciled=   90.00   where TransactionID=   21629   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/08/2018#Full',Reconciled=   90.00   where TransactionID=   22815   "
        adoconn.Execute "Update tlbreceipt set Reconnow='02/03/2018#Full',Reconciled=   90.00   where TransactionID=   17892   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/08/2018#Full',Reconciled=   90.00   where TransactionID=   22820   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/09/2018#Full',Reconciled=   90.00   where TransactionID=   24215   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/09/2018#Full',Reconciled=   90.00   where TransactionID=   24217   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/08/2018#Full',Reconciled=   90.00   where TransactionID=   22817   "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/10/2018#Full',Reconciled=   90.00   where TransactionID=   24221   "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/10/2018#Full',Reconciled=   90.00   where TransactionID=   24223   "
        adoconn.Execute "Update tlbreceipt set Reconnow='09/10/2018#Full',Reconciled=   90.00   where TransactionID=   24224   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/09/2018#Full',Reconciled=   90.00   where TransactionID=   24216   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/05/2018#Full',Reconciled=   90.00   where TransactionID=   20193   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/07/2018#Full',Reconciled=   90.00   where TransactionID=   21625   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/05/2018#Full',Reconciled=   90.00   where TransactionID=   20191   "
        adoconn.Execute "Update tlbreceipt set Reconnow='17/04/2018#Full',Reconciled=   90.00   where TransactionID=   19436   "
        adoconn.Execute "Update tlbreceipt set Reconnow='17/04/2018#Full',Reconciled=   90.00   where TransactionID=   19438   "
        adoconn.Execute "Update tlbreceipt set Reconnow='17/04/2018#Full',Reconciled=   90.00   where TransactionID=   19435   "
        adoconn.Execute "Update tlbreceipt set Reconnow='10/05/2018#Full',Reconciled=   90.00   where TransactionID=   20189   "
        adoconn.Execute "Update tlbreceipt set Reconnow='15/06/2018#Full',Reconciled=   90.00   where TransactionID=   21620   "
        adoconn.Execute "Update tlbreceipt set Reconnow='17/04/2018#Full',Reconciled=   90.00   where TransactionID=   19434   "
        adoconn.Execute "Update tlbreceipt set Reconnow='02/03/2018#Full',Reconciled=   90.00   where TransactionID=   17893   "
        adoconn.Execute "Update tlbreceipt set Reconnow='15/06/2018#Full',Reconciled=   90.00   where TransactionID=   21622   "
        adoconn.Execute "Update tlbreceipt set Reconnow='15/06/2018#Full',Reconciled=   90.00   where TransactionID=   21624   "
        adoconn.Execute "Update tlbreceipt set Reconnow='03/07/2017#Full',Reconciled=   90.00   where TransactionID=   13250   "
        
        
        
         'Now rollback last bank recociliation
        Dim szRecondate As String
        Dim adoRst As New ADODB.Recordset
        szRecondate = "14/11/2018"
        adoconn.Execute "DELETE FROM tlbBankReconClosingBal where ClientID='OAKTREE' AND BankCode='1210' AND StatementDate>=#" & Format(szRecondate, "dd MMM yyyy") & "#"
        adoconn.Execute "DELETE FROM tlbBankReconcilation where ClientID='OAKTREE' AND BankCode='1210' AND ReconDate>=#" & Format(szRecondate, "dd MMM yyyy") & "# "
        adoconn.Execute "Update tlbReceipt,Units,Property SET tlbReceipt.ReconNow= NULL,tlbReceipt.Reconciled= NULL where tlbReceipt.UnitID=Units.UnitNumber" & _
                " AND Units.PropertyID=Property.PropertyID and Property.ClientID='OAKTREE' AND BankCode='1210' " & _
                "AND  CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(szRecondate, "dd MMM yyyy") & "# AND Reconnow is not NULL"
        adoconn.Execute "Update tlbPayment SET ReconNow= NULL, Reconciled = NULL where ClientID='OAKTREE' AND BankCode='1210' " & _
                "AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(szRecondate, "dd MMM yyyy") & "#"
        adoconn.Execute "Update tlbBankPayment SET ReconNow= NULL,Reconciled=NULL where ClientID='OAKTREE' AND Bank_AC='1210' " & _
                "AND CDate( Left(iif(isnull(Reconnow),#31 Mar 1930#,Reconnow),10))>=#" & Format(szRecondate, "dd MMM yyyy") & "#"

        adoRst.Open "Select *  FROM tlbBankReconClosingBal where ClientID='OAKTREE' AND BankCode='1210' order by StatementDate ", adoconn, adOpenKeyset, adLockReadOnly
        If adoRst.EOF Then
           adoconn.Execute "Update tlbClientBanks set spare2=NULL , ClosingBal=0, SOB=0,PCB=0  where Client_ID='OAKTREE' AND NominalCode='1210'"
        Else
           adoRst.MoveLast
           adoconn.Execute "Update tlbClientBanks set spare2='" & Format(adoRst.Fields("StatementDate").Value, "dd MMM yyyy") & "' , SOB=" & _
           adoRst.Fields("ProjClbal").Value & ",PCB=0,ClosingBal=" & adoRst.Fields("ProjClbal").Value & " where CLIENT_ID='OAKTREE' AND NominalCode='1210'"
           ',PCB=" & adoRst.Fields("ProjClbal").Value & "
        End If
        
        

         adoconn.Close
          MsgBox "Routine completed successfully."
     End If
     If iSelRow = 12 Then 'Inconsistency in Bank Reconciliation ‘ORCHGREEN’ date 2019/06/10
        If MsgBox("Do you want to run the routine to fix 'Inconsistency in Bank Reconciliation ‘ORCHGREEN’ '?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update NLPOSTING SET NOMINAL_CODE='1210' where UNIQUE_REFERENCE_NO=104709"
        adoconn.Execute "Update tlbPayment SET NOMINALCODE='1210' where TransactionID=16511"
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     If iSelRow = 13 Then '"The Lessee WR151A, Transaction SR62072.OSamount Incorrect"
        If MsgBox("Do you want to run the routine to fix SR62072 ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update tlbReceipt SET OSAmount=amount,ReceiptView=true where TransactionID=77094"
         adoconn.Execute "Update tlbReceiptSplit SET OSAmount=amount where RptHeader=77094"
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
      If iSelRow = 14 Then 'Marking paid to client with statementID
        If MsgBox("Do you want to run the routine to fix 'Marking paid to client with statementID' ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update tlbpaymentsplit SET clientstatementID=341 where Payheader=5373 "
         
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     If iSelRow = 15 Then 'Marking paid to client with statementID
        If MsgBox("Do you want to run the routine to fix SR62072 ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update tlbReceipt SET OSAmount=amount,ReceiptView=true where TransactionID=75117"
         adoconn.Execute "Update tlbReceiptSplit SET OSAmount=amount where RptHeader=75117"
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     If iSelRow = 16 Then 'Marking paid to client with statementID
        If MsgBox("Do you want to run the routine to fix 'Some data are not posted into NLPosting' ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        Call fixReceipts(adoconn)
        Call fixPayments(adoconn)
        Call fixBankPayments(adoconn)
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     If iSelRow = 17 Then 'Marking paid to client with statementID
        If MsgBox("Do you want to run the routine to fix 'fix inconsistent data ob Batch Payment' ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update tlbPaymentSplit set OSAmount=0 where PayHeader=47696"
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     
     If iSelRow = 18 Then 'Marking paid to client with statementID
        If MsgBox("Do you want to update all Retention Bank Code from client statement ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
        adoconn.Execute "Update Retentiondetails R, RentSummaryStatement S set R.BankCode=S.BankCode where R.StatementID=S.StatementID and R.BankCode is Null"
        
        adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     If iSelRow = 19 Then
        If MsgBox("Do you want to clear all locks in the system ?", vbYesNo) = vbNo Then Exit Sub
        If adocon.State = 1 Then
            adocon.Close
        End If
        adoconn.Open getConnectionString
            adoconn.Execute "UPDATE tlbPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
            adoconn.Execute "UPDATE tlbReceipt P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
            adoconn.Execute "UPDATE tlbBankPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
            adoconn.Execute "UPDATE NJ_Header P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
            adoconn.Close
        Set adoconn = Nothing
        MsgBox "Routine completed successfully."
     End If
     
     'ChecksumValidationOnAllocationAC
'     fraCheckSumPayment.Visible = True
'         On Error Resume Next
'         adoConn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=PBMControl.mdb;DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Backup _20190115;Uid=;Pwd=RDSWKDPP;"
'         'adoConn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
'         adodata.Open "Select * from databases", adoConn, adOpenStatic, adLockReadOnly
'             i = 1
'           While Not adodata.EOF
'               MSHFlexGrid5.AddItem ""
'               'MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
'               MSHFlexGrid5.TextMatrix(i, 1) = adodata("SCName").Value
'               i = i + 1
'               adocon.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & adodata("dbName").Value & ";DefaultDir=U:\Support Issues\Austin Chambers\ALL Prestige Data For Testing\Prestige Data 20181206;Uid=;Pwd=RDSWKDPP;"
'               'adocon.Open "DSN=" & adodata("AccessDsn").Value & ";UID=;PWD=" & accessDBPws & ""
'               Call ChecksumValidationOnPayAllocationAC(adocon, i)
'               adodata.MoveNext
'           Wend
     
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SetMouseCursor IDC_ARROW
End Sub
Private Sub fixReceipts(adoconn As ADODB.Connection)
    Dim rsCheckReceipt As New ADODB.Recordset
    Dim szTransactionID  As String
     rsCheckReceipt.Open "SELECT R.TransactionID ,TYPE from tlbReceipt R  LEFT join NLposting N   ON cstr(R.TransactionID)=N.TRANSACTION_REF where " & _
            "TRANSACTION_TYPE in (1,2,3,4,23) and DeleteFlag=false and TRANSACTION_ref is NULL", adoconn, adOpenDynamic, adLockReadOnly

                If Not rsCheckReceipt.EOF Then
                        szTransactionID = SQL2String(rsCheckReceipt, 1)
                        rsCheckReceipt.MoveFirst
                 End If
                 While Not rsCheckReceipt.EOF
                    adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckReceipt("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & ""
                     rsCheckReceipt.MoveNext
                 Wend
                 
                MsgBox RecordCount(rsCheckReceipt) & " Receipt record(s) have been posted "
                If Len(szTransactionID) > 0 Then
                        adoconn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  in ( " & szTransactionID & ")"
                End If
                rsCheckReceipt.Close
                Set rsCheckReceipt = Nothing
End Sub
Private Sub fixPayments(adoconn As ADODB.Connection)
    Dim rsCheckPayments As New ADODB.Recordset
    Dim szTransactionID  As String
     rsCheckPayments.Open "SELECT P.* from tlbPayment P  LEFT join NLposting N   ON cstr(P.TransactionID)=N.TRANSACTION_REF " & _
                    "where TRANSACTION_TYPE in (6,7,8,9,24) and DeleteFlag=false and TRANSACTION_ref is NULL", adoconn, adOpenDynamic, adLockReadOnly

                If Not rsCheckPayments.EOF Then
                        szTransactionID = SQL2String(rsCheckPayments, 1)
                        rsCheckPayments.MoveFirst
                 End If
                 MsgBox RecordCount(rsCheckPayments) & " Payment record(s) have been posted "
                 While Not rsCheckPayments.EOF
                    adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckPayments("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckPayments("Type").Value & ""
                    rsCheckPayments.MoveNext
                 Wend
                If Len(szTransactionID) > 0 Then
                        adoconn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  in ( " & szTransactionID & ")"
                End If
                rsCheckPayments.Close
                Set rsCheckPayments = Nothing
End Sub
Private Sub fixBankPayments(adoconn As ADODB.Connection)
    Dim rsCheckBankPayments As New ADODB.Recordset
    Dim szTransactionID  As String
     rsCheckBankPayments.Open "SELECT BP.* from tlbBankPayment BP  LEFT join NLposting N   ON cstr(BP.TRAN_ID)=N.TRANSACTION_REF where  " & _
                    "TRANSACTION_TYPE in (11,12) and DeleteFlag=false and TRANSACTION_ref is NULL ", adoconn, adOpenDynamic, adLockReadOnly

                If Not rsCheckBankPayments.EOF Then
                        szTransactionID = SQL2String(rsCheckBankPayments, 1)
                        rsCheckBankPayments.MoveFirst
                 End If
                 MsgBox RecordCount(rsCheckBankPayments) & " BankPayment record(s) have been posted "
                 While Not rsCheckBankPayments.EOF
                    adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckBankPayments("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckBankPayments("Type").Value & ""
                    rsCheckBankPayments.MoveNext
                 Wend
                If Len(szTransactionID) > 0 Then
                        adoconn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  in ( " & szTransactionID & ")"
                End If
                rsCheckBankPayments.Close
                Set rsCheckBankPayments = Nothing
End Sub
Private Sub loadcboCurStDt()
    '#
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String
    adoconn.Open getConnectionString
    Dim adoRst As New ADODB.Recordset
    Dim ReConDates()        As String
   
    szSQL = "SELECT StatementDate, BankCode " & _
           "FROM tlbBankReconClosingBal " & _
           "WHERE BankCode = '" & txtBC.Tag & "' " & _
           "AND ClientID = '" & txtClientList.Tag & "' GROUP BY StatementDate, BankCode " & _
           "ORDER BY StatementDate DESC;"
   
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim ReConDates(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol
         If Not IsNull(adoRst.Fields(j).Value) And adoRst.Fields(j).Value <> "" Then
            ReConDates(j, i) = adoRst.Fields(j).Value
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   cboCurStDt.Column() = ReConDates()
  

   adoRst.Close
   adoconn.Close
End Sub

Private Sub cboIssue_Change()
'    If cboIssue.ListIndex = 0 Then
'        Frame1.Visible = True
'     Else
'        Frame1.Visible = False
'    End If
'    If cboIssue.ListIndex = 2 Then
'        FraReconciliation.Visible = True
'        cmdSupplier.Caption = "RollBack"
'     Else
'        FraReconciliation.Visible = False
'         cmdSupplier.Caption = "Fix it"
'
'    End If
'    If cboIssue.ListIndex = 3 Then
'        cmdShowReport.Visible = True
'     Else
'        cmdShowReport.Visible = False
'     End If
'     If cboIssue.ListIndex = 5 Then
'        Call ChecksumValidationOnAllocation
'     End If
End Sub
Private Function ChecksumValidationOnAllocation() As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    Dim adoconn As New ADODB.Connection
    'this function is written by anol 20181123 when found that (issue 673 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This function shall compare allocation with receipt amount and outstanding amount
    adoconn.Open getConnectionString
    MSHFlexGrid4.Clear
    MSHFlexGrid4.Cols = 2
    MSHFlexGrid4.Rows = 2
    MSHFlexGrid4.ColWidth(1) = 3000
    'Exit Function
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
     Dim i As Integer
    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID  AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoconn, adOpenStatic, adLockReadOnly
                    
'        If rsChecksum.EOF Then
'                ChecksumValidationOnAllocation = True
'        Else
'        rsChecksum.MoveFirst
                While Not rsChecksum.EOF
                    szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    MSHFlexGrid4.AddItem ""
                    MSHFlexGrid4.TextMatrix(i, 1) = "SI" + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    i = i + 1
                    rsChecksum.MoveNext
                Wend
'        End If
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL  AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             MSHFlexGrid4.AddItem ""
             MSHFlexGrid4.TextMatrix(i, 1) = "SI" + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
                ChecksumValidationOnAllocation = True
                MsgBox "No problematic records found"
        Else
'                MsgBox "Transaction that has problem: " & Chr(13) & szTran2Fix & "Total records: " & i, vbInformation, "Report of SI!"
        End If
        Label17.Caption = i
        adoconn.Close
        Set adoconn = Nothing
                
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)

'receipt check is fine on allocation
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt, FromTran from
'RptTransactions  group By FromTran ) as A where A.FromTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Private Function ChecksumValidationOnPayAllocation() As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
    Dim adoconn As New ADODB.Connection
    'this function is written by anol 20181123 when found that (issue 695 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This function shall compare allocation with payment amount and outstanding amount
    adoconn.Open getConnectionString
    MSHFlexGrid5.Clear
    MSHFlexGrid5.Cols = 2
    MSHFlexGrid5.Rows = 2
    MSHFlexGrid5.ColWidth(1) = 3000
    'Exit Function
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
     Dim i As Integer
     
'    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
'                    "Where a.ToTran = r.TransactionID  AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
     rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
                  " ToTran from PayTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)", adoconn, adOpenStatic, adLockReadOnly
                    
'        If rsChecksum.EOF Then
'                ChecksumValidationOnPayAllocation = True
'        Else
'        rsChecksum.MoveFirst
                While Not rsChecksum.EOF
                    szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    MSHFlexGrid5.AddItem ""
                    MSHFlexGrid5.TextMatrix(i, 1) = "PI" + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    i = i + 1
                    rsChecksum.MoveNext
                Wend
'        End If
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
        adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbPayment R LEFT JOIN PayTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL  AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             MSHFlexGrid5.AddItem ""
             MSHFlexGrid5.TextMatrix(i, 1) = "PI" + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
                ChecksumValidationOnPayAllocation = True
                MsgBox "No problematic records found"
        Else
'                MsgBox "Transaction that has problem: " & Chr(13) & szTran2Fix & "Total records: " & i, vbInformation, "Report of SI!"
                cmdFixPayAllocation.Visible = True
        End If
        Label12.Caption = i
        adoconn.Close
        Set adoconn = Nothing
                
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)

'receipt check is fine on allocation
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt, FromTran from
'RptTransactions  group By FromTran ) as A where A.FromTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Private Function ChecksumValidationOnAllocationAC(adoconn As ADODB.Connection, i As Integer) As Boolean  ' if returns true then true means data is fine and false means there is some inconsistent data
    'Dim adoConn As New ADODB.Connection
    'this function is written by anol 20181123 when found that (issue 673 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This function shall compare allocation with receipt amount and outstanding amount
   
'    MSHFlexGrid4.Clear
'    MSHFlexGrid4.Cols = 2
'    MSHFlexGrid4.Rows = 2
    MSHFlexGrid4.ColWidth(1) = 3000
    'Exit Function
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
     'Dim i As Integer
    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID  AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoconn, adOpenStatic, adLockReadOnly
                    
'        If rsChecksum.EOF Then
'                ChecksumValidationOnAllocation = True
'        Else
'        rsChecksum.MoveFirst
                While Not rsChecksum.EOF
                    szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    MSHFlexGrid4.AddItem ""
                    MSHFlexGrid4.TextMatrix(i, 1) = "SI" + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    i = i + 1
                    rsChecksum.MoveNext
                Wend
'        End If
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL  AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             MSHFlexGrid4.AddItem ""
             MSHFlexGrid4.TextMatrix(i, 1) = "SI" + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
'                ChecksumValidationOnAllocation = True
               ' MsgBox "No problematic records found"
        Else
'                MsgBox "Transaction that has problem: " & Chr(13) & szTran2Fix & "Total records: " & i, vbInformation, "Report of SI!"
        End If
        Label17.Caption = i
        adoconn.Close
        Set adoconn = Nothing
                
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)

'receipt check is fine on allocation
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt, FromTran from
'RptTransactions  group By FromTran ) as A where A.FromTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function

Private Function ChecksumValidationOnPayAllocationAC(adoconn As ADODB.Connection, i As Integer) As Boolean ' if returns true then true means data is fine and false means there is some inconsistent data
   ' Dim
    'this function is written by anol 20181123 when found that (issue 695 )Updated OS amount 102 extra incorrectly
    'This function shall prevent saving the data if when outstading amount on receipt is not updated.
    'This function shall compare allocation with payment amount and outstanding amount
    'adoconn.Open getConnectionString
   ' MSHFlexGrid5.Clear
    MSHFlexGrid5.Cols = 2
    'MSHFlexGrid5.Rows = 2
    MSHFlexGrid5.ColWidth(1) = 3000
    'Exit Function
    Dim rsChecksum As New ADODB.Recordset
    Dim adoRst As New ADODB.Recordset
    Dim szTran2Fix As String
     'Dim i As Integer
     
     'i = MSHFlexGrid5.Rows + 1
     
'    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
'                    "Where a.ToTran = r.TransactionID  AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoConn, adOpenStatic, adLockReadOnly
     rsChecksum.Open "Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbPayment R,(select Sum(PaymentAmount) as amt," & _
                  " ToTran from PayTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)", adoconn, adOpenStatic, adLockReadOnly
                    
'        If rsChecksum.EOF Then
'                ChecksumValidationOnPayAllocation = True
'        Else
'        rsChecksum.MoveFirst
                While Not rsChecksum.EOF
                    szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    MSHFlexGrid5.AddItem ""
                    MSHFlexGrid5.TextMatrix(i, 1) = "PI" + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    i = i + 1
                    rsChecksum.MoveNext
                Wend
'        End If
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
        adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbPayment R LEFT JOIN PayTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL  AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "PI", ",PI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             MSHFlexGrid5.AddItem ""
             MSHFlexGrid5.TextMatrix(i, 1) = "PI" + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
'                ChecksumValidationOnPayAllocation = True
'                MsgBox "No problematic records found"
        Else
'                MsgBox "Transaction that has problem: " & Chr(13) & szTran2Fix & "Total records: " & i, vbInformation, "Report of SI!"
        End If
        Label12.Caption = i
'        adoconn.Close
'        Set adoconn = Nothing
        
       
    rsChecksum.Open "Select SlNumber,amount,osamount,amt,R.sageaccountnumber from tlbReceipt R,(select Sum(ReceiptAmount) as amt,ToTran from RptTransactions  group By ToTran ) as A " & _
                    "Where a.ToTran = r.TransactionID  AND Round((amount - amt), 2) <> Round(OSAmount, 2)", adoconn, adOpenStatic, adLockReadOnly
                    
'        If rsChecksum.EOF Then
'                ChecksumValidationOnAllocation = True
'        Else
'        rsChecksum.MoveFirst
                While Not rsChecksum.EOF
                    szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    MSHFlexGrid5.AddItem ""
                    MSHFlexGrid5.TextMatrix(i, 1) = "SI" + CStr(rsChecksum("SlNumber").Value) + " (" + rsChecksum("sageaccountnumber").Value + ") "
                    i = i + 1
                    rsChecksum.MoveNext
                Wend
'        End If
        rsChecksum.Close
        Set rsChecksum = Nothing
     'for which records does not have entry in the allocation but osamount>amount
         adoRst.Open "Select R.transactionID,R.SlNumber,sageaccountnumber,amount,osamount from tlbReceipt R LEFT JOIN RptTransactions A ON " & _
                "A.Totran=R.transactionID where A.Totran IS NULL  AND osamount>amount", adoconn, adOpenKeyset, adLockReadOnly
        While Not adoRst.EOF
             szTran2Fix = szTran2Fix + IIf(szTran2Fix = "", "SI", ",SI") + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             MSHFlexGrid5.AddItem ""
             MSHFlexGrid5.TextMatrix(i, 1) = "SI" + CStr(adoRst("SlNumber").Value) + " (" + adoRst("sageaccountnumber").Value + ") "
             i = i + 1
             adoRst.MoveNext
        Wend
        adoRst.Close
        Set adoRst = Nothing
        If szTran2Fix = "" Then
'                ChecksumValidationOnAllocation = True
               ' MsgBox "No problematic records found"
        Else
'                MsgBox "Transaction that has problem: " & Chr(13) & szTran2Fix & "Total records: " & i, vbInformation, "Report of SI!"
        End If
        Label17.Caption = i
        adoconn.Close
        Set adoconn = Nothing
                
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt,
'                   ToTran from RptTransactions  group By ToTran ) as A where A.ToTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)

'receipt check is fine on allocation
'Select  R.transactionID,R.SlNumber,'',R.sageaccountnumber,amount,osamount,amt from tlbReceipt R,(select Sum(ReceiptAmount) as amt, FromTran from
'RptTransactions  group By FromTran ) as A where A.FromTran=R.transactionID AND round((amount-amt),2)<>round(osamount,2)
                    
End Function
Private Sub cboIssue_Click()
    cboIssue_Change
End Sub

Private Sub cboIssue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtStart.SetFocus
    End If
End Sub

Private Sub cmdBC_Click()
'    picClient.Left = 1355.029
'    picClient.Top = 155.299
    picClient.Left = 10755
    picClient.Top = 345
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    sTextBox = "2"
    ConfigureFlxBank
    Dim szAllBankBalance As String
    szAllBankBalance = BankAndBalance(adoconn)
    adoconn.Close
    Set adoconn = Nothing
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

Private Sub configflxissueList()
    flxIssueList.Clear
    flxIssueList.Cols = 5
    flxIssueList.Rows = 9
    flxIssueList.ColWidth(0) = 200 'for viewing selection
    flxIssueList.ColWidth(1) = 1200 'for Viewing date
    flxIssueList.ColWidth(2) = 7000 'Description
    flxIssueList.ColWidth(3) = 1200 'for viewing issue number
    flxIssueList.ColWidth(4) = 0

End Sub
Private Sub ConfigureFlxBank()
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 5
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.ColWidth(3) = 0
   flxClient.ColWidth(4) = 0
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment = vbLeftJustify
   lblClientID.Caption = "Bank Code"
   lblClientName.Caption = "Bank Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
End Sub
Private Function BankAndBalance(adoconn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim rRow As Integer
   
  
         szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
 
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
   'Modified by anol 22 Feb 2015
            If txtClientList.text <> "Consolidated" Then
                    MsgBox "Please setup your Client Bank Accounts." & Chr(13) & _
                           "Please also check the nominal chart of account for the client."
             End If
             
   Else

                rRow = 0
                While Not adoRst.EOF
                    flxClient.row = 1
                    flxClient.TextMatrix(rRow, 0) = ""
                    flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("BNC").Value
                    flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("BNN").Value
                    flxClient.TextMatrix(rRow, 3) = adoRst.Fields.Item("ID").Value
                    flxClient.TextMatrix(rRow, 4) = adoRst.Fields.Item("CLIENT_ID").Value
                    flxClient.RowHeight(rRow) = 280
                    adoRst.MoveNext
                    If Not adoRst.EOF Then flxClient.AddItem ""
                    rRow = rRow + 1
                 Wend
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   'LoadAdoBank

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    cmdClientList.SetFocus
End Sub

Private Sub cmdUnlock_Click()
    Dim adoconn As New ADODB.Connection
    
    adoconn.Open getConnectionString
    adoconn.Execute "Delete From recordlocking"
    adoconn.Close
    MsgBox "Screen has been unlocked.", vbInformation, "Unlocked"
End Sub

Private Sub flxClient_Click()
    If sTextBox = "1" Then
        txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
        txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
        txtBC.Tag = ""
        txtBC.text = ""
        cboCurStDt.Clear
        cmdBC.SetFocus
    ElseIf sTextBox = "2" Then
        txtBC.Tag = flxClient.TextMatrix(flxClient.row, 1)
        txtBC.text = flxClient.TextMatrix(flxClient.row, 1)
        Call loadcboCurStDt
    End If
    picClient.Visible = False
        
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
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
           cmdClientList.SetFocus
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
Private Sub cmdClientList_Click()
    picClient.Left = 10755
    picClient.Top = 345
    sTextBox = "1"
    LoadflxClient
    
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub FIX_PaymentEdit_ReceiptEdit(strCon As String)
On Error GoTo Err
   Dim rsCheckPayment As New ADODB.Recordset
   Dim rsCheckReceipt As New ADODB.Recordset
   Dim adoRst As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection
   adoconn.Open strCon
   Dim rsFixPaymentNominalCode  As New ADODB.Recordset
   Dim szTransactionID As String
   Dim szTHIS_RECORD As String
   Dim rsPLC As New ADODB.Recordset
   Dim rsSLC As New ADODB.Recordset
   Dim rsPITransactions As New ADODB.Recordset
   Dim rsPITransactionsZeroValue As New ADODB.Recordset
   ProgressBar2.Min = 0
   ProgressBar2.Max = 100
   ProgressBar2.Visible = True
        'Issue 440 Invoice number ,receipt and payment number was showing incorrectly in NLhistory
        'added by anol 20170807
        'IN NLPosting table transaction ref was empty so they needs to be updated from the receipt ,payment and bank transaction table.
        'By this order as well
        ' IN the NLposting table there was some empty Transaction_ref. That was neeed to build up SI no or PI no or receipt
        'I am updating that
        adoconn.Execute "UPDATE NLPOSTING, tlbReceipt SET TRANSACTION_REF = slNumber " & _
        "WHERE NLPOSTING.TRANS_ID=cstr(tlbReceipt.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbReceipt.Type AND TRANSACTION_REF is NULL;"
        'issue 552
        adoconn.Execute "Update  tblPurInv INNER JOIN tlbPayment ON tblPurInv.MY_ID = tlbPayment.PI SET tlbPayment.slnumber=tblPurInv.SlNumber " & _
        "WHERE  tlbPayment.slnumber<>tblPurInv.SlNumber AND (tblPurInv.TransactionType=6 Or tblPurInv.TransactionType=7);"
        '    Type 24,8,9
        adoconn.Execute "UPDATE NLPOSTING, tlbPayment SET TRANSACTION_REF = tlbPayment.slNumber ,NLPOSTING.TRANS_ID=tlbPayment.slNumber " & _
        "WHERE NLPOSTING.PARENT_RECORD=cstr(tlbPayment.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbPayment.Type AND " & _
        "NLPOSTING.TRANSACTION_REF is NULL AND  NLPOSTING.TRANS_ID IS NULL;"
        'Type 6,7,11,12
        adoconn.Execute "UPDATE NLPOSTING SET TRANSACTION_REF = TRANS_ID where NLPOSTING.TRansaction_TYpe in (11,12,24,8,9,6,7) AND TRANSACTION_REF is NULL"
        
        'Type 1,2
        adoconn.Execute "UPDATE NLPOSTING, DemandRecords SET TRANSACTION_REF = DmdSlNo " & _
        "WHERE NLPOSTING.TRANS_ID=cstr(DemandRecords.DemandID) AND NLPOSTING.TRansaction_TYpe=DemandRecords.TransactionType " & _
        "AND TRANSACTION_REF is NULL; "
        fraReport.Visible = True
        trxCount.Visible = True
        trxCount.Caption = "0"
        trxCount2.Visible = True
        trxCount2.Caption = "0"
        trxCount3.Visible = True
        trxCount3.Caption = "0"
        
'1)In the help menu I need to work with this SQL and make them zero in the tlbpayment table
'2)In tlbPurINvSplit should insert a deleted split with nominal code But this does not have nominal code
'    So you canot post it again to NL , you need to make them zero as they are inconsistent data no further posting
 'if you found this problem exists then there is some transaction in tblPurInv which are inconsitent
        'Now by comparing with the NLposting this transaction will be zerorized
        'To prevent happening this again there is check I have implemented in PI
        rsPITransactionsZeroValue.Open "SELECT P.TransactionType, P.SlNumber, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, " & _
        "P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID, tblPurInvSRec.TRAN_ID, tblPurInvSRec.NOMINAL_CODE, " & _
        "tblPurInvSRec.DESCRIPTION FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID  " & _
        "WHERE ((P.TransactionType)=6 Or (P.TransactionType)=7) AND tran_ID is null;", adoconn, adOpenKeyset, adLockReadOnly
         While Not rsPITransactionsZeroValue.EOF
            adoconn.Execute "Update tblPurINV SET  tblPurINV.NLPOST=true,TOTAL_AMOUNT=0 where tblPurINV.TransactionType=" & rsPITransactionsZeroValue("TransactionType").Value & _
                " AND SlNumber=" & rsPITransactionsZeroValue("SlNumber").Value & ""
            adoconn.Execute "Update tlbPayment SET Amount=0,osAmount=0 where SlNumber=" & rsPITransactionsZeroValue("SlNumber").Value & _
                " AND Type=" & rsPITransactionsZeroValue("TransactionType").Value & " AND Reconciled is NULL"
            rsPITransactionsZeroValue.MoveNext
         Wend
    'Now code start for Correcting mismatch PI type 6 and 7
    'Find the Purchase Ledger control first.Loop client and get a PI LD CTRL CODE
    'The FIND THE PI numbers that are not posted ,then repost them
        rsPLC.Open "SELECT Code,ClientID FROM NominalLedger WHERE CAName = 'Purchase Ledger Control'", adoconn, adOpenKeyset, adLockReadOnly
        While Not rsPLC.EOF
            If rsPLC("ClientID").Value = "TWINOAKS" Then
                Debug.Print ""
            End If
            rsPITransactions.Open "SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE,Y.TYPE,Y.NominalCode,Y.PI FROM " & _
            "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where " & _
            "DELETEFLAG=FALSE AND NOMINAL_CODE='" & rsPLC("Code").Value & "' AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(TRANSACTION_TYPE=6 OR TRANSACTION_TYPE=7) GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X " & _
            "Right Join " & _
            "(SELECT IIF((R.Type=6) ,(-R.AMOUNT),(R.AMOUNT)) as AMOUNT2 ,R.Type, R.SageAccountNumber, " & _
            "R.SlNumber,R.NominalCode,R.PI FROM tlbPayment R WHERE AMOUNT<>0 AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(R.Type=6 OR R.Type=7)) AS Y " & _
            "ON X.TRANS_ID=cstr(Y.SlNumber) AND X.TRANSACTION_TYPE=Y.Type where X.AMOUNT1 IS NULL", adoconn, adOpenKeyset, adLockReadOnly
            While Not rsPITransactions.EOF
                    adoconn.Execute "Update tblPurINV SET  tblPurINV.NLPOST=False where tblPurINV.MY_ID='" & rsPITransactions("PI").Value & "'"
                    adoconn.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & rsPITransactions("TRANS_ID").Value & _
                                    "' AND TRANSACTION_TYPE=" & rsPITransactions("TYPE").Value & " "
                     'Debug.Print rsPLC("ClientID").Value
                     'Debug.Print rsPITransactions("PI").Value
                rsPITransactions.MoveNext
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount.Caption = Val(trxCount.Caption) + 1
                trxCount.Refresh
                If ProgressBar2.Value > 98 Then
                        ProgressBar2.Max = 3000
                End If
            Wend
            rsPITransactions.Close
            rsPLC.MoveNext
        Wend
        'after tblPurINv.NLPOST=False I should delete the tlbpayment entries and create the again
        'The FIND THE PI numbers that are mismatched ,then repost them, Type 6,7
        rsPLC.MoveFirst
        While Not rsPLC.EOF
            rsPITransactions.Open "SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE,Y.TYPE,Y.NominalCOde,Y.PI FROM " & _
            "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where " & _
            "DELETEFLAG=FALSE AND NOMINAL_CODE='" & rsPLC("Code").Value & "' AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(TRANSACTION_TYPE=6 OR TRANSACTION_TYPE=7) GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X " & _
            "INNER Join " & _
            "(SELECT IIF((R.Type=6) ,(-R.AMOUNT),(R.AMOUNT)) as AMOUNT2 ,R.Type, R.SageAccountNumber, " & _
            "R.SlNumber,R.NominalCOde,R.PI FROM tlbPayment R WHERE ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(R.Type=6 OR R.Type=7)) AS Y " & _
            "ON X.TRANS_ID=cstr(Y.SlNumber) AND X.TRANSACTION_TYPE=Y.Type where Y.AMOUNT2<> X.AMOUNT1", adoconn, adOpenKeyset, adLockReadOnly
            While Not rsPITransactions.EOF
                    adoconn.Execute "Update tblPurINV SET  tblPurINv.NLPOST=False where tblPurINV.MY_ID='" & rsPITransactions("PI").Value & "'"
                    adoconn.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & rsPITransactions("TRANS_ID").Value & _
                                    "' AND TRANSACTION_TYPE=" & rsPITransactions("TYPE").Value & " "
                rsPITransactions.MoveNext
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount.Caption = Val(trxCount.Caption) + 1
                trxCount.Refresh
                If ProgressBar2.Value > 98 Then
                        ProgressBar2.Max = 3000
                End If
            Wend
            rsPITransactions.Close
            rsPLC.MoveNext
        Wend
        
       
         
        adoconn.Close
        adoconn.Open strCon
         Export_PInPC_2_NL adoconn
    '        correcting nominalCode from bank account that was correct
        'Select B.*,A.*  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8 AND A.deleteflag=false
'        adoConn.Execute "Update tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) Set B.NLPost=false,A.deleteflag=true,B.NominalCode=B.bankCode  " & _
                        "where A.clientID=B.clientID and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8"
          'updating two table step by step
'          adoConn.Execute "Update tlbpayment B set B.NominalCode=B.bankCode, B.NLPost=false where B.TransactionID IN " & _
'                            "(Select B.*,A.*  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID and " & _
'                                "A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8 AND A.deleteflag=false)"
'1.  Problem : bankCode not NominalCode equal nominal Code. This is because nomnal code was not updating when user was changing the bank code.
'so you need to set set B.NominalCode=B.bankCode

          rsFixPaymentNominalCode.Open "Select B.TransactionID,A.THIS_RECORD  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID " & _
                            "and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND A.deleteflag=false", adoconn, adOpenDynamic, adLockReadOnly
          If Not rsFixPaymentNominalCode.EOF Then
                szTransactionID = SQL2String(rsFixPaymentNominalCode, 0)
                rsFixPaymentNominalCode.MoveFirst
          End If
          If Not rsFixPaymentNominalCode.EOF Then
                szTHIS_RECORD = SQL2StringQuote(rsFixPaymentNominalCode, 1)
          End If
          If Len(szTransactionID) > 0 Then
                adoconn.Execute "Update tlbpayment B set B.NominalCode=B.bankCode, B.NLPost=false where B.TransactionID IN (" & szTransactionID & ")"
          End If
          If Len(szTHIS_RECORD) > 0 Then
                adoconn.Execute "Update NLPOSTING N set N.deleteflag=true where N.THIS_RECORD IN (" & szTHIS_RECORD & ")"
          End If
'         adoConn.Execute "Update tlbpayment B SET NominalCode=bankCode  where bankCode<>NominalCode AND Type=8"
                        'client er joint nae
                        
                        
        'reversing the payment which have mismatch amount
        rsCheckPayment.Open "Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type from (Select * from tlbPayment A  INNER join (Select sum( amount)as " & _
                            "NLamount,NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, " & _
                            "clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID " & _
                            "AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode)", adoconn, adOpenDynamic, adLockReadOnly
                           
'         Query For amount mismatch with NLposting from payment  Edit :
'         Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type from (Select * from tlbPayment A  INNER join (Select sum( amount)as NLamount,
'         NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE,
'         clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND
'         abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode)
        'Problem 2) - payment edit amount not same In NLpsting
        '1.  reversing the payment which have mismatched amount with tlb receipt and NLpostosting.
        '2.  Select the transaction that are mismatched . set NLPost=false for tlbpayment
        '3.  Post those transaction again.

         adoconn.Execute "Update tlbPayment B  SET  B.NLPost=false  where  TransactionID  in ( " & _
                            "Select  A.TransactionID  from (Select * from tlbPayment A  INNER join (Select sum( amount)as " & _
                            "NLamount,NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, " & _
                            "clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID " & _
                            "AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) ) "
                            
        While Not rsCheckPayment.EOF
               adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANSACTION_REF='" & rsCheckPayment("TRANSACTION_REF").Value & "' AND TRANSACTION_TYPE=" & rsCheckPayment("Type").Value & "" & _
                    " AND N.clientID ='" & rsCheckPayment("ClientID").Value & "'"
               rsCheckPayment.MoveNext
               ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount2.Caption = Val(trxCount2.Caption) + 1
                trxCount2.Refresh
               If ProgressBar2.Value > 98 Then
                        ProgressBar2.Max = 3000
                End If
        Wend
        rsCheckPayment.Close
        Set rsCheckPayment = Nothing
        adoconn.Close
        adoconn.Open strCon
        Export_PPnPPR_2_NL adoconn
        
        
        'Now for receipt edit error
        adoconn.Close
        adoconn.Open strCon
        'Query for Sales receipt mismatch amount with NL amount:
''        SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE FROM
''(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where DELETEFLAG=FALSE AND
''(TRANSACTION_TYPE=3 OR TRANSACTION_TYPE=4 OR TRANSACTION_TYPE=23)
''GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X LEFT JOIN
''(SELECT IIF( (R.Type=23),-R.AMOUNT,R.AMOUNT) as AMOUNT2 ,R.Type, R.SageAccountNumber,R.UnitID, R.RDate, R.Ref, R.PostingDate, R.BankCode,
''
''R.ExtRef, R.NominalCode,
''R.TransactionID,  R.SlNumber FROM tlbReceipt R WHERE (R.Type=3 Or R.Type=4 Or R.Type=23)) AS Y
''ON X.TRANS_ID=cstr(Y.TransactionID) AND X.TRANSACTION_TYPE=Y.Type AND X.NOMINAL_CODE=Y.BankCode where X.AMOUNT1<>Y.AMOUNT2
''
''Nominal COde ,back code thek update dite hobe tarpor EI SQL GULO chalate hobe
''
''Left Join
''SELECT X.TRANS_ID,Y.TransactionID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE FROM
''(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where DELETEFLAG=FALSE
''AND NOMINAL_CODE<>'1100' AND
''(TRANSACTION_TYPE=3 OR TRANSACTION_TYPE=4 OR TRANSACTION_TYPE=23)
''GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X LEFT JOIN
''(SELECT IIF( (R.Type=23),-R.AMOUNT,R.AMOUNT) as AMOUNT2 ,R.Type, R.SageAccountNumber,R.UnitID, R.RDate, R.Ref, R.PostingDate, R.BankCode,
''
''R.ExtRef, R.NominalCode,
''R.TransactionID,  R.SlNumber FROM tlbReceipt R WHERE (R.Type=3 Or R.Type=4 Or R.Type=23)) AS Y
''ON X.TRANS_ID=cstr(Y.TransactionID) AND X.TRANSACTION_TYPE=Y.Type AND X.NOMINAL_CODE=Y.BankCode where Y.TransactionID IS NULL
''
''Right Join
''SELECT X.TRANS_ID,Y.TransactionID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE FROM
''(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where DELETEFLAG=FALSE
''AND NOMINAL_CODE<>'1100' AND
''(TRANSACTION_TYPE=3 OR TRANSACTION_TYPE=4 OR TRANSACTION_TYPE=23)
''GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X RIGHT JOIN
''(SELECT IIF( (R.Type=23),-R.AMOUNT,R.AMOUNT) as AMOUNT2 ,R.Type, R.SageAccountNumber,R.UnitID, R.RDate, R.Ref, R.PostingDate, R.BankCode,
''
''R.ExtRef, R.NominalCode,
''R.TransactionID,  R.SlNumber FROM tlbReceipt R WHERE (R.Type=3 Or R.Type=4 Or R.Type=23)) AS Y
''ON X.TRANS_ID=cstr(Y.TransactionID) AND X.TRANSACTION_TYPE=Y.Type AND X.NOMINAL_CODE=Y.BankCode where  X.TRANS_ID IS NULL


        'Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE,
        'clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON
        'X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode)
               
        adoconn.Execute "Update tlbReceipt B  SET NominalCode=BankCode where NominalCode<>BankCode"
        ''' mismatch amount with NL amount posting:
''        rsCheckReceipt.Open "Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE," & _
''        "clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON " & _
''        "X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode)", adoconn, adOpenDynamic, adLockReadOnly
''        'updating NLposting for deleting
''        While Not rsCheckReceipt.EOF
''           'Deleteing mismatched record from NLPosting
''                adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANSACTION_REF='" & rsCheckReceipt("TRANSACTION_REF").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & "" & _
''                    " AND N.clientID ='" & rsCheckReceipt("ClientID").Value & "'"
''                rsCheckReceipt.MoveNext
''                ProgressBar2.Value = ProgressBar2.Value + 1
''                trxCount3.Caption = Val(trxCount3.Caption) + 1
''                trxCount3.Refresh
''                If ProgressBar2.Value > 98 Then
''                        ProgressBar2.Max = 3000
''                End If
''        Wend
''        rsCheckReceipt.Close
''        Set rsCheckReceipt = Nothing
        'From receipt table we need to post that transaction
        'A.Those who had amount mismatch
         
         Debug.Print "Select  A.TransactionID,A.Type,A.Slnumber from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE," & _
                "clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON " & _
                "X.TRANSACTION_REF =cstr(A.slnumber) AND X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) "
           'this works only for 3,4,23 because in receipt table nominal code is empty for other types
        rsCheckReceipt.Open "Select  A.TransactionID,A.Type,A.Slnumber from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE," & _
        "clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON " & _
        "X.TRANSACTION_REF =cstr(A.slnumber) AND X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) " _
        , adoconn, adOpenDynamic, adLockReadOnly
         While Not rsCheckReceipt.EOF
                adoconn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  =" & rsCheckReceipt("TransactionID").Value & ""
                adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckReceipt("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & ""
                Debug.Print rsCheckReceipt("TransactionID").Value
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount3.Caption = Val(trxCount3.Caption) + 1
                trxCount3.Refresh
                rsCheckReceipt.MoveNext
        Wend
        rsCheckReceipt.Close
        Set rsCheckReceipt = Nothing

         'B.Those who was never posted.
        rsSLC.Open "SELECT Code,ClientID FROM NominalLedger WHERE CAName = 'Sales Ledger Control'", adoconn, adOpenKeyset, adLockReadOnly
        While Not rsSLC.EOF
                 rsCheckReceipt.Open "SELECT X.TRANS_ID,Y.TransactionID,X.AMOUNT1,Y.AMOUNT2,Y.TYPE,X.NOMINAL_CODE FROM " & _
                "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where DELETEFLAG=FALSE " & _
                "AND NOMINAL_CODE<>'" & rsSLC("Code").Value & "' AND ClientID='" & rsSLC("ClientID").Value & "' AND " & _
                "(TRANSACTION_TYPE=3 OR TRANSACTION_TYPE=4 OR TRANSACTION_TYPE=23) " & _
                "GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X RIGHT JOIN " & _
                "(SELECT IIF( (R.Type=23),-R.AMOUNT,R.AMOUNT) as AMOUNT2 ,R.Type, R.SageAccountNumber,R.UnitID, R.RDate, R.Ref, R.PostingDate, R.BankCode, " & _
                "R.ExtRef, R.NominalCode,R.TransactionID,  R.SlNumber FROM tlbReceipt R WHERE ClientID='" & rsSLC("ClientID").Value & "' AND (R.Type=3 Or R.Type=4 Or R.Type=23)) AS Y " & _
                "ON X.TRANS_ID=cstr(Y.TransactionID) AND X.TRANSACTION_TYPE=Y.Type AND X.NOMINAL_CODE=Y.BankCode where  X.TRANS_ID IS NULL", adoconn, adOpenDynamic, adLockReadOnly
                 If Not rsCheckReceipt.EOF Then
                        szTransactionID = SQL2String(rsCheckReceipt, 1)
                        rsCheckReceipt.MoveFirst
                  End If
                 While Not rsCheckReceipt.EOF
                   'Deleting mismatched record from NLPosting
                        adoconn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckReceipt("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & ""
                        Debug.Print rsCheckReceipt("TransactionID").Value
                        rsCheckReceipt.MoveNext
                        ProgressBar2.Value = ProgressBar2.Value + 1
                        trxCount3.Caption = Val(trxCount3.Caption) + 1
                        trxCount3.Refresh
                        Debug.Print rsSLC("ClientID").Value
                        If ProgressBar2.Value > 98 Then
                                ProgressBar2.Max = 3000
                        End If
                Wend
                If Len(szTransactionID) > 0 Then
                        adoconn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  in ( " & szTransactionID & ")"
                End If
                rsCheckReceipt.Close
                Set rsCheckReceipt = Nothing
            rsSLC.MoveNext
        Wend
        adoconn.Close
        adoconn.Open strCon
        Export_SRnSRR_2_NL adoconn
        
        adoconn.Close
         ProgressBar2.Visible = False
        Exit Sub
Err:
        MsgBox Err.description
End Sub
Private Sub FIX_PaymentEdit_ReceiptEdit_Report(strCon As String)
On Error GoTo Err
   Dim rsCheckPayment As New ADODB.Recordset
   Dim rsCheckReceipt As New ADODB.Recordset
   Dim adoRst As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection
   Dim iRow  As Integer
   Dim TransactionType As String
   adoconn.Open strCon
   Dim rsFixPaymentNominalCode  As New ADODB.Recordset
   Dim szTransactionID As String
   Dim szTHIS_RECORD As String
   Dim rsPLC As New ADODB.Recordset
   Dim rsSLC As New ADODB.Recordset
   Dim rsPITransactions As New ADODB.Recordset
   Dim rsPITransactionsZeroValue As New ADODB.Recordset
   ProgressBar2.Min = 0
   ProgressBar2.Max = 100
   ProgressBar2.Visible = True
        'Issue 440 Invoice number ,receipt and payment number was showing incorrectly in NLhistory
        'added by anol 20170807
        'IN NLPosting table transaction ref was empty so they needs to be updated from the receipt ,payment and bank transaction table.
        'By this order as well
        ' IN the NLposting table there was some empty Transaction_ref. That was neeed to build up SI no or PI no or receipt
        'I am updating that
        adoconn.Execute "UPDATE NLPOSTING, tlbReceipt SET TRANSACTION_REF = slNumber " & _
        "WHERE NLPOSTING.TRANS_ID=cstr(tlbReceipt.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbReceipt.Type AND TRANSACTION_REF is NULL;"
        'issue 552
        adoconn.Execute "Update  tblPurInv INNER JOIN tlbPayment ON tblPurInv.MY_ID = tlbPayment.PI SET tlbPayment.slnumber=tblPurInv.SlNumber " & _
        "WHERE  tlbPayment.slnumber<>tblPurInv.SlNumber AND (tblPurInv.TransactionType=6 Or tblPurInv.TransactionType=7);"
        '    Type 24,8,9
        adoconn.Execute "UPDATE NLPOSTING, tlbPayment SET TRANSACTION_REF = tlbPayment.slNumber ,NLPOSTING.TRANS_ID=tlbPayment.slNumber " & _
        "WHERE NLPOSTING.PARENT_RECORD=cstr(tlbPayment.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbPayment.Type AND " & _
        "NLPOSTING.TRANSACTION_REF is NULL AND  NLPOSTING.TRANS_ID IS NULL;"
        'Type 6,7,11,12
        adoconn.Execute "UPDATE NLPOSTING SET TRANSACTION_REF = TRANS_ID where NLPOSTING.TRansaction_TYpe in (11,12,24,8,9,6,7) AND TRANSACTION_REF is NULL"
        
        'Type 1,2
        adoconn.Execute "UPDATE NLPOSTING, DemandRecords SET TRANSACTION_REF = DmdSlNo " & _
        "WHERE NLPOSTING.TRANS_ID=cstr(DemandRecords.DemandID) AND NLPOSTING.TRansaction_TYpe=DemandRecords.TransactionType " & _
        "AND TRANSACTION_REF is NULL; "
        fraReport.Visible = True
        trxCount.Visible = True
        trxCount.Caption = "0"
        trxCount2.Visible = True
        trxCount2.Caption = "0"
        trxCount3.Visible = True
        trxCount3.Caption = "0"
        
'1)In the help menu I need to work with this SQL and make them zero in the tlbpayment table
'2)In tlbPurINvSplit should insert a deleted split with nominal code But this does not have nominal code
'    So you canot post it again to NL , you need to make them zero as they are inconsistent data no further posting
 'if you found this problem exists then there is some transaction in tblPurInv which are inconsitent
        'Now by comparing with the NLposting this transaction will be zerorized
        'To prevent happening this again there is check I have implemented in PI
        rsPITransactionsZeroValue.Open "SELECT P.TransactionType, P.SlNumber, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, " & _
        "P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID, tblPurInvSRec.TRAN_ID, tblPurInvSRec.NOMINAL_CODE, " & _
        "tblPurInvSRec.DESCRIPTION FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID  " & _
        "WHERE ((P.TransactionType)=6 Or (P.TransactionType)=7) AND tran_ID is null;", adoconn, adOpenKeyset, adLockReadOnly
         While Not rsPITransactionsZeroValue.EOF
            adoconn.Execute "Update tblPurINV SET  tblPurINV.NLPOST=true,TOTAL_AMOUNT=0 where tblPurINV.TransactionType=" & rsPITransactionsZeroValue("TransactionType").Value & _
                " AND SlNumber=" & rsPITransactionsZeroValue("SlNumber").Value & ""
            adoconn.Execute "Update tlbPayment SET Amount=0,osAmount=0 where SlNumber=" & rsPITransactionsZeroValue("SlNumber").Value & _
                " AND Type=" & rsPITransactionsZeroValue("TransactionType").Value & " AND Reconciled is NULL"
            rsPITransactionsZeroValue.MoveNext
         Wend
    'Now code start for Correcting mismatch PI type 6 and 7
    'Find the Purchase Ledger control first.Loop client and get a PI LD CTRL CODE
    'The FIND THE PI numbers that are not posted ,then repost them
        rsPLC.Open "SELECT Code,ClientID FROM NominalLedger WHERE CAName = 'Purchase Ledger Control'", adoconn, adOpenKeyset, adLockReadOnly
        iRow = 1
        While Not rsPLC.EOF
'            If rsPLC("ClientID").Value = "TWINOAKS" Then
'                Debug.Print ""
'            End If
            rsPITransactions.Open "SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE,Y.TYPE,Y.NominalCode,Y.PI,Y.SLNumber FROM " & _
            "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where " & _
            "DELETEFLAG=FALSE AND NOMINAL_CODE='" & rsPLC("Code").Value & "' AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(TRANSACTION_TYPE=6 OR TRANSACTION_TYPE=7) GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X " & _
            "Right Join " & _
            "(SELECT IIF((R.Type=6) ,(-R.AMOUNT),(R.AMOUNT)) as AMOUNT2 ,R.Type, R.SageAccountNumber, " & _
            "R.SlNumber,R.NominalCode,R.PI FROM tlbPayment R WHERE AMOUNT<>0 AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(R.Type=6 OR R.Type=7)) AS Y " & _
            "ON X.TRANS_ID=cstr(Y.SlNumber) AND X.TRANSACTION_TYPE=Y.Type where X.AMOUNT1 IS NULL", adoconn, adOpenKeyset, adLockReadOnly
            
            While Not rsPITransactions.EOF
'                    adoConn.Execute "Update tblPurINV SET  tblPurINV.NLPOST=False where tblPurINV.MY_ID='" & rsPITransactions("PI").Value & "'"
'                    adoConn.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & rsPITransactions("TRANS_ID").Value & _
'                                    "' AND TRANSACTION_TYPE=" & rsPITransactions("TYPE").Value & " "
                     'Debug.Print rsPLC("ClientID").Value
                     'Debug.Print rsPITransactions("PI").Value
                     MSHFlexGrid1.TextMatrix(iRow, 1) = IIf(rsPITransactions("TYPE").Value = 6, "PI", "PC") & rsPITransactions("SLNumber").Value
                     iRow = iRow + 1
                     MSHFlexGrid1.AddItem ""
                rsPITransactions.MoveNext
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount.Caption = Val(trxCount.Caption) + 1
                trxCount.Refresh
                If ProgressBar2.Value > 98 Then
                        ProgressBar2.Max = 3000
                End If
            Wend
            rsPITransactions.Close
            rsPLC.MoveNext
        Wend
        'after tblPurINv.NLPOST=False I should delete the tlbpayment entries and create the again
        'The FIND THE PI numbers that are mismatched ,then repost them, Type 6,7
        rsPLC.MoveFirst
        While Not rsPLC.EOF
            rsPITransactions.Open "SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE,Y.TYPE,Y.NominalCOde,Y.PI,Y.SLNumber FROM " & _
            "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where " & _
            "DELETEFLAG=FALSE AND NOMINAL_CODE='" & rsPLC("Code").Value & "' AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(TRANSACTION_TYPE=6 OR TRANSACTION_TYPE=7) GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X " & _
            "INNER Join " & _
            "(SELECT IIF((R.Type=6) ,(-R.AMOUNT),(R.AMOUNT)) as AMOUNT2 ,R.Type, R.SageAccountNumber, " & _
            "R.SlNumber,R.NominalCOde,R.PI FROM tlbPayment R WHERE ClientID='" & rsPLC("ClientID").Value & "' AND " & _
            "(R.Type=6 OR R.Type=7)) AS Y " & _
            "ON X.TRANS_ID=cstr(Y.SlNumber) AND X.TRANSACTION_TYPE=Y.Type where Y.AMOUNT2<> X.AMOUNT1", adoconn, adOpenKeyset, adLockReadOnly
            'iRow = 1
            While Not rsPITransactions.EOF
'                    adoConn.Execute "Update tblPurINV SET  tblPurINv.NLPOST=False where tblPurINV.MY_ID='" & rsPITransactions("PI").Value & "'"
'                    adoConn.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & rsPITransactions("TRANS_ID").Value & _
'                                    "' AND TRANSACTION_TYPE=" & rsPITransactions("TYPE").Value & " "
                     MSHFlexGrid1.TextMatrix(iRow, 1) = IIf(rsPITransactions("TYPE").Value = 6, "PI", "PC") & rsPITransactions("SLNumber").Value
                     iRow = iRow + 1
                     MSHFlexGrid1.AddItem ""
                rsPITransactions.MoveNext
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount.Caption = Val(trxCount.Caption) + 1
                trxCount.Refresh
                If ProgressBar2.Value > 98 Then
                        ProgressBar2.Max = 3000
                End If
            Wend
            rsPITransactions.Close
            rsPLC.MoveNext
        Wend
        
       
         
        adoconn.Close
        adoconn.Open strCon
        ' Export_PInPC_2_NL adoConn
    '        correcting nominalCode from bank account that was correct
        'Select B.*,A.*  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8 AND A.deleteflag=false
'        adoConn.Execute "Update tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) Set B.NLPost=false,A.deleteflag=true,B.NominalCode=B.bankCode  " & _
                        "where A.clientID=B.clientID and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8"
          'updating two table step by step
'          adoConn.Execute "Update tlbpayment B set B.NominalCode=B.bankCode, B.NLPost=false where B.TransactionID IN " & _
'                            "(Select B.*,A.*  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID and " & _
'                                "A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND Type=8 AND A.deleteflag=false)"
'1.  Problem : bankCode not NominalCode equal nominal Code. This is because nominal code was not updating when user was changing the bank code.
'so you need to set set B.NominalCode=B.bankCode

          rsFixPaymentNominalCode.Open "Select B.TransactionID,A.THIS_RECORD  from  tlbpayment B INNER JOIN NLPosting A ON A.TRANSACTION_REF =cstr(B.slnumber) where A.clientID=B.clientID " & _
                            "and A.TRANSACTION_TYPE in(8,9,24) AND bankCode<>NominalCode AND A.deleteflag=false", adoconn, adOpenDynamic, adLockReadOnly
          If Not rsFixPaymentNominalCode.EOF Then
                szTransactionID = SQL2String(rsFixPaymentNominalCode, 0)
                rsFixPaymentNominalCode.MoveFirst
          End If
          If Not rsFixPaymentNominalCode.EOF Then
                szTHIS_RECORD = SQL2StringQuote(rsFixPaymentNominalCode, 1)
          End If
          If Len(szTransactionID) > 0 Then
                adoconn.Execute "Update tlbpayment B set B.NominalCode=B.bankCode, B.NLPost=false where B.TransactionID IN (" & szTransactionID & ")"
          End If
          If Len(szTHIS_RECORD) > 0 Then
                adoconn.Execute "Update NLPOSTING N set N.deleteflag=true where N.THIS_RECORD IN (" & szTHIS_RECORD & ")"
          End If
'         adoConn.Execute "Update tlbpayment B SET NominalCode=bankCode  where bankCode<>NominalCode AND Type=8"
                        'client er joint nae
                        
                        
        'reversing the payment which have mismatch amount
        rsCheckPayment.Open "Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type,A.SLNumber from (Select * from tlbPayment A  INNER join (Select sum( amount)as " & _
                            "NLamount,NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, " & _
                            "clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID " & _
                            "AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) ", adoconn, adOpenStatic, adLockReadOnly
                           
'         Query For amount mismatch with NLposting from payment  Edit :
'         Select  a.amount,X.NLamount,A.ClientID,TRANSACTION_REF,A.Type from (Select * from tlbPayment A  INNER join (Select sum( amount)as NLamount,
'         NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE,
'         clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND
'         abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode)
        'Problem 2) - payment edit amount not same In NLpsting
        '1.  reversing the payment which have mismatched amount with tlb receipt and NLpostosting.
        '2.  Select the transaction that are mismatched . set NLPost=false for tlbpayment
        '3.  Post those transaction again.

'         adoConn.Execute "Update tlbPayment B  SET  B.NLPost=false  where  TransactionID  in ( " & _
'                            "Select  A.TransactionID  from (Select * from tlbPayment A  INNER join (Select sum( amount)as " & _
'                            "NLamount,NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, " & _
'                            "clientID, deleteflag ,TRANSACTION_REF) as X ON X.TRANSACTION_REF =cstr(A.slnumber )  AND  X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID " & _
'                            "AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) ) "
        iRow = 1
        MSHFlexGrid2.Rows = rsCheckPayment.RecordCount + 1
        
        While Not rsCheckPayment.EOF
'               adoConn.Execute "Update NLPosting N SET N.deleteflag=true where TRANSACTION_REF='" & rsCheckPayment("TRANSACTION_REF").Value & "' AND TRANSACTION_TYPE=" & rsCheckPayment("Type").Value & "" & _
'                    " AND N.clientID ='" & rsCheckPayment("ClientID").Value & "'"
                     If rsCheckPayment("Type").Value = 24 Then
                            TransactionType = "PPR"
                     ElseIf rsCheckPayment("Type").Value = 6 Then
                            TransactionType = "PI"
                     ElseIf rsCheckPayment("Type").Value = 7 Then
                            TransactionType = "PC"
                     ElseIf rsCheckPayment("Type").Value = 8 Then
                            TransactionType = "PP"
                     ElseIf rsCheckPayment("Type").Value = 9 Then
                            TransactionType = "PA"
                     Else
                        TransactionType = ""
                     End If
                     
                     MSHFlexGrid2.TextMatrix(iRow, 1) = TransactionType & rsCheckPayment("SLNumber").Value
                     iRow = iRow + 1
                    ' MSHFlexGrid2.AddItem ""
               rsCheckPayment.MoveNext
               ProgressBar2.Value = ProgressBar2.Value + 1
               trxCount2.Caption = Val(trxCount2.Caption) + 1
               trxCount2.Refresh
               If ProgressBar2.Value > 98 Then
                     ProgressBar2.Max = 3000
               End If
        Wend
        rsCheckPayment.Close
        Set rsCheckPayment = Nothing
        adoconn.Close
         TransactionType = ""
'        adoConn.Open strCon
'        Export_PPnPPR_2_NL adoConn
'
'
'        'Now for receipt edit error
'        adoConn.Close
        adoconn.Open strCon
       
               
        adoconn.Execute "Update tlbReceipt B  SET NominalCode=BankCode where NominalCode<>BankCode"
       
         
'         Debug.Print "Select  A.TransactionID,A.Type,A.Slnumber from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE," & _
'                "clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON " & _
'                "X.TRANSACTION_REF =cstr(A.slnumber) AND X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) "
           'this works only for 3,4,23 because in receipt table nominal code is empty for other types
        rsCheckReceipt.Open "Select  A.TransactionID,A.Type,A.Slnumber,A.Type from (Select * from tlbreceipt A  INNER join (Select sum( amount)as NLamount,NOMINAL_CODE,TRANSACTION_TYPE," & _
        "clientID, deleteflag ,TRANSACTION_REF   from NLPosting where deleteflag=false  group by NOMINAL_CODE,TRANSACTION_TYPE, clientID, deleteflag ,TRANSACTION_REF) as X ON " & _
        "X.TRANSACTION_REF =cstr(A.slnumber) AND X.TRANSACTION_TYPE=A.Type AND A.ClientID=X.ClientID AND abs(a.amount)<>abs(X.NLamount) AND X.NOMINAL_CODE=A.NominalCode) " _
        , adoconn, adOpenDynamic, adLockReadOnly
        iRow = 1
         While Not rsCheckReceipt.EOF
'                adoConn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  =" & rsCheckReceipt("TransactionID").Value & ""
'                adoConn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckReceipt("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & ""
'                Debug.Print rsCheckReceipt("TransactionID").Value
                     If rsCheckReceipt("Type").Value = 3 Then
                            TransactionType = "SR"
                     ElseIf rsCheckReceipt("Type").Value = 4 Then
                            TransactionType = "SA"
                     ElseIf rsCheckReceipt("Type").Value = 23 Then
                            TransactionType = "SRR"
                     Else
                            TransactionType = ""
                     End If
                     
                     MSHFlexGrid3.TextMatrix(iRow, 1) = TransactionType & rsCheckReceipt("SLNumber").Value
                     iRow = iRow + 1
                     MSHFlexGrid3.AddItem ""
                ProgressBar2.Value = ProgressBar2.Value + 1
                trxCount3.Caption = Val(trxCount3.Caption) + 1
                trxCount3.Refresh
                rsCheckReceipt.MoveNext
        Wend
        rsCheckReceipt.Close
        Set rsCheckReceipt = Nothing

         'B.Those who was never posted.
        rsSLC.Open "SELECT Code,ClientID FROM NominalLedger WHERE CAName = 'Sales Ledger Control'", adoconn, adOpenKeyset, adLockReadOnly
        While Not rsSLC.EOF
                 rsCheckReceipt.Open "SELECT X.TRANS_ID,Y.TransactionID,X.AMOUNT1,Y.AMOUNT2,Y.TYPE,X.NOMINAL_CODE,Y.slnumber FROM " & _
                "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where DELETEFLAG=FALSE " & _
                "AND NOMINAL_CODE<>'" & rsSLC("Code").Value & "' AND ClientID='" & rsSLC("ClientID").Value & "' AND " & _
                "(TRANSACTION_TYPE=3 OR TRANSACTION_TYPE=4 OR TRANSACTION_TYPE=23) " & _
                "GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X RIGHT JOIN " & _
                "(SELECT IIF( (R.Type=23),-R.AMOUNT,R.AMOUNT) as AMOUNT2 ,R.Type, R.SageAccountNumber,R.UnitID, R.RDate, R.Ref, R.PostingDate, R.BankCode, " & _
                "R.ExtRef, R.NominalCode,R.TransactionID,  R.SlNumber FROM tlbReceipt R WHERE ClientID='" & rsSLC("ClientID").Value & "' AND (R.Type=3 Or R.Type=4 Or R.Type=23)) AS Y " & _
                "ON X.TRANS_ID=cstr(Y.TransactionID) AND X.TRANSACTION_TYPE=Y.Type AND X.NOMINAL_CODE=Y.BankCode where  X.TRANS_ID IS NULL", adoconn, adOpenDynamic, adLockReadOnly
                 If Not rsCheckReceipt.EOF Then
                        szTransactionID = SQL2String(rsCheckReceipt, 1)
                        rsCheckReceipt.MoveFirst
                  End If
                 While Not rsCheckReceipt.EOF
                   'Deleting mismatched record from NLPosting
'                        adoConn.Execute "Update NLPosting N SET N.deleteflag=true where TRANS_ID='" & rsCheckReceipt("TransactionID").Value & "' AND TRANSACTION_TYPE=" & rsCheckReceipt("Type").Value & ""
'                        Debug.Print rsCheckReceipt("TransactionID").Value
                            If rsCheckReceipt("Type").Value = 3 Then
                                TransactionType = "SR"
                            ElseIf rsCheckReceipt("Type").Value = 4 Then
                                   TransactionType = "SA"
                            ElseIf rsCheckReceipt("Type").Value = 23 Then
                                   TransactionType = "SRR"
                            Else
                                    TransactionType = ""
                            End If
                     
                     MSHFlexGrid3.TextMatrix(iRow, 1) = TransactionType & rsCheckReceipt("SLNumber").Value
                     iRow = iRow + 1
                     MSHFlexGrid3.AddItem ""
                     
                        rsCheckReceipt.MoveNext
                        ProgressBar2.Value = ProgressBar2.Value + 1
                        trxCount3.Caption = Val(trxCount3.Caption) + 1
                        trxCount3.Refresh
'                        Debug.Print rsSLC("ClientID").Value
                        If ProgressBar2.Value > 98 Then
                                ProgressBar2.Max = 3000
                        End If
                Wend
'                If Len(szTransactionID) > 0 Then
'                        adoConn.Execute "Update tlbReceipt B  SET  B.NLPost=false  where  TransactionID  in ( " & szTransactionID & ")"
'                End If
                rsCheckReceipt.Close
                Set rsCheckReceipt = Nothing
            rsSLC.MoveNext
        Wend
        adoconn.Close
'        adoConn.Open strCon
'        Export_SRnSRR_2_NL adoConn
'
'        adoConn.Close
         ProgressBar2.Visible = False
        Exit Sub
Err:
        MsgBox Err.description
End Sub

Private Sub cmdFixtransactions_Click()
    MSHFlexGrid1.Visible = False
    MSHFlexGrid2.Visible = False
    MSHFlexGrid3.Visible = False
    
   Dim adoconn As New ADODB.Connection
   cmdFixTransactions.Enabled = False
   adoconn.Open getConnectionString
   Dim adoRst As New ADODB.Recordset
   
   Dim rsFixPaymentNominalCode  As New ADODB.Recordset
   Dim szTransactionID As String
   Dim szTHIS_RECORD As String
   Dim iCount  As Integer
   adoconn.Execute "UPDATE ShoppingCentre SET Field1 = '', Field2 = '';"
'   If MsgBox("Do you wish to run delete blank tables routine?", vbYesNo, "Warning") = vbYes Then
'         Delit adoConn
'   End If
'    If cboIssue.ListIndex = 4 Then
'        adoConn.Execute "Update  tlbreceipt A LEFT JOIN tlbReceiptsplit B on A.transactionID=B.rptheader SET A.OSamount=A.amount,ReceiptView=true  " & _
'        "where B.rptheader is null and A.transactionID not in ( Select Fromtran from rpttransactions)"
'    End If
    If cboIssue.ListIndex = 3 Then 'Payment Edit was not posting Data to NL FIX
            adoconn.Close
            adoconn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
            If MsgBox("Do you want to fix the data for all the databases", vbYesNo, "") = vbYes Then
                 ProgressBar1.Min = 0
                 ProgressBar1.Max = 30
                 ProgressBar1.Visible = True
                 adoRst.Open "Select * from Databases", adoconn, adOpenDynamic, adLockOptimistic
                 While Not adoRst.EOF
                    iCount = iCount + 1
                    ProgressBar1.Value = iCount
                    Call FIX_PaymentEdit_ReceiptEdit("DSN=" & adoRst.Fields("AccessDsn").Value & ";UID=;PWD=" & accessDBPws & "")
                    
'                    MsgBox adoRST.AbsolutePosition
                    adoRst.MoveNext
                 Wend
                  ProgressBar1.Visible = False
                 adoRst.Close
                
            Else
                 Call FIX_PaymentEdit_ReceiptEdit(getConnectionString)
            End If
           
    End If
    If cboIssue.ListIndex = 2 Then
    
    End If
''for westgate only
    If cboIssue.ListIndex = 1 Then
          
          
   End If
    If cboIssue.ListIndex = 0 Then
           
   End If
   
   
   adoconn.Close
   Set adoconn = Nothing
   cmdFixTransactions.Enabled = True
   ShowMsgInTaskBar "DONE", "Y", "Y"
'    MsgBox "DONE"
End Sub
'Private Sub UpdateNL(con As ADODB.Connection)
'     ' Dim rsAllNL As New ADODB.Recordset
'    rsAllNL.Open "Select PARENT_RECORD,AMOUNT,NOMINAL_CODE from NLPOSTING", con, adOpenKeyset, adLockReadOnly
'    While Not rsAllNL.EOF
'
'        rsAllNL.MoveNext
'    Wend
'    MsgBox "Done"
'End Sub
Private Sub Delit(con As ADODB.Connection)
con.Execute "Drop table  Agent;"
con.Execute "Drop table  AttachedFile;"
con.Execute "Drop table  Batches;"
con.Execute "Drop table  ChargeTypes;"
con.Execute "Drop table  ClientGDFees_;"
con.Execute "Drop table  ClientGlobalData;"
con.Execute "Drop table  ClientProAgr;"
con.Execute "Drop table  Contacts;"
con.Execute "Drop table  DemandCategory_;"
con.Execute "Drop table  DemandRecPreview;"
con.Execute "Drop table  DemandSplPreview;"
con.Execute "Drop table  GlobalInsurance;"
con.Execute "Drop table  GlobalRC;"
con.Execute "Drop table  GlobalSCDtls;"
con.Execute "Drop table  InterestRates;"
con.Execute "Drop table  Landlord;"
con.Execute "Drop table  LeaseAssignments;"
con.Execute "Drop table  LeaseBreaches;"
con.Execute "Drop table  LInsuranceCharges;"
con.Execute "Drop table  LUtilityUsage;"
con.Execute "Drop table  MemoDetails;"
con.Execute "Drop table  NJ_CC;"
con.Execute "Drop table  NJ_Header;"
con.Execute "Drop table  NJ_Split;"
con.Execute "Drop table  PayableTypes;"
con.Execute "Drop table  PayTransactions;"
con.Execute "Drop table  PropertyAnalysis;"
con.Execute "Drop table  PropertyInsurance;"
con.Execute "Drop table  PropertyLandlord;"
con.Execute "Drop table  PropertyMaintHistory;"
con.Execute "Drop table  PropertySafety;"
con.Execute "Drop table  PropertyUtilities;"
con.Execute "Drop table  RentAnalysis;"
con.Execute "Drop table  ReportCategory;"
con.Execute "Drop table  Schedule;"
con.Execute "Drop table  SpareTable1;"
con.Execute "Drop table  SpareTable10;"
con.Execute "Drop table  SpareTable2;"
con.Execute "Drop table  SpareTable3;"
con.Execute "Drop table  SpareTable4;"
con.Execute "Drop table  SpareTable5;"
con.Execute "Drop table  SpareTable6;"
con.Execute "Drop table  SpareTable7;"
con.Execute "Drop table  SpareTable8;"
con.Execute "Drop table  SpareTable9;"
con.Execute "Drop table  SubUnitType;"
con.Execute "Drop table  tblBatchPayment;"
con.Execute "Drop table  tblBatchReceipt;"
con.Execute "Drop table  tblBatchTransaction;"
con.Execute "Drop table  tblBtRptTran;"
con.Execute "Drop table  tblPoA;"
con.Execute "Drop table  tblPrevGLU;"
con.Execute "Drop table  tblPurInv;"
con.Execute "Drop table  tblPurInvSRec;"
con.Execute "Drop table  TemplateUnitSelection;"
con.Execute "Drop table  TenantBankDetails;"
con.Execute "Drop table  TenantDeposit;"
con.Execute "Drop table  TenantEventHistory;"
con.Execute "Drop table  tlbAgreement;"
con.Execute "Drop table  tlbBankPayment;"
con.Execute "Drop table  tlbBankReconcilation;"
con.Execute "Drop table  tlbBankReconClosingBal;"
con.Execute "Drop table  tlbChildDemandRecord;"
con.Execute "Drop table  tlbCreditNote;"
con.Execute "Drop table  tlbFloor;"
con.Execute "Drop table  tlbImages;"
con.Execute "Drop table  tlbLetterReports;"
con.Execute "Drop table  tlbMemo;"
con.Execute "Drop table  tlbPayable;"
con.Execute "Drop table  tlbPayment;"
con.Execute "Drop table  tlbPaymentSplit;"
con.Execute "Drop table  tlbRecharged;"
con.Execute "Drop table  tlbRechargePre;"
con.Execute "Drop table  tlbSupplierPayment_;"
con.Execute "Drop table  UnitInsurance;"
con.Execute "Drop table  UnitMaintHistory;"
con.Execute "Drop table  UnitSafety;"
con.Execute "Drop table  UnitUtilities;"
con.Execute "ALTER Table FUND ADD column FundList text(255);"
con.Execute "ALTER Table Tenants ADD column SLControl text(50);"
con.Execute "ALTER Table Tenants ADD column DefaultNC text(15);"
con.Execute "ALTER Table Tenants ADD column VAT_CODE text(5);"
MsgBox "Droping the blank table and adding column tenant and fund done"
'    con.Execute "Drop table ChargingMethod;"
'con.Execute "Drop table Client;"
'con.Execute "Drop table DemandRecords;"
'con.Execute "Drop table DemandSplitRecords;"
'con.Execute "Drop table DemandTypes;"
'con.Execute "Drop table FinancialYear;"
'con.Execute "Drop table Frequencies;"
'con.Execute "Drop table Fund;"
'con.Execute "Drop table GlobalData;"
'con.Execute "Drop table GlobalSC;"
'con.Execute "Drop table LeaseDetails;"
'con.Execute "Drop table LeaseHistory;"
'con.Execute "Drop table LRentCharges;"
'con.Execute "Drop table LServiceCharges;"
'con.Execute "Drop table NLCategory;"
'con.Execute "Drop table NLPosting;"
'con.Execute "Drop table NLSubTypes;"
'con.Execute "Drop table NLType;"
'con.Execute "Drop table NominalLedger;"
''con.Execute "Drop table NominalLedger_OLD;"
'con.Execute "Drop table PaymentDates;"
'con.Execute "Drop table Periods;"
'con.Execute "Drop table PrimaryCode;"
'con.Execute "Drop table Property;"
'con.Execute "Drop table RptTransactions;"
'con.Execute "Drop table SecondaryCode;"
'con.Execute "Drop table ShoppingCentre;"
'con.Execute "Drop table Supplier;"
'con.Execute "Drop table Template;"
'con.Execute "Drop table Tenants;"
'con.Execute "Drop table tlbBank;"
'con.Execute "Drop table tlbClientBanks;"
'con.Execute "Drop table tlbDRCurrentPrint;"
'con.Execute "Drop table tlbID;"
'con.Execute "Drop table tlbReceipt;"
'con.Execute "Drop table tlbReceiptSplit;"
'con.Execute "Drop table tlbRef;"
'con.Execute "Drop table tlbTransactionTypes;"
'con.Execute "Drop table tlbVatCode;"
'con.Execute "Drop table UnitAnalysis;"
'con.Execute "Drop table Units;"
'con.Execute "Drop table UserNames;"

End Sub

Public Sub SelectOnly1RowFlxGrid(conFlxGrid As Control, iNewRow As Integer, Optional iColID As Integer = 0)
   Dim iRow       As Integer
   Dim iCol       As Integer
   Dim iColPaint  As Integer

   iColPaint = IIf(iColID = 0, 1, 0)
   
   For iRow = 1 To conFlxGrid.Rows - 1
      If conFlxGrid.TextMatrix(iRow, iColID) = "X" Then
         If iRow = iNewRow And conFlxGrid.TextMatrix(iRow, iColID) = "X" Then Exit Sub
         conFlxGrid.TextMatrix(iRow, iColID) = ""
         conFlxGrid.row = iRow
         For iCol = iColPaint To conFlxGrid.Cols - 1
            conFlxGrid.col = iCol
            conFlxGrid.CellBackColor = vbWhite
         Next iCol
      End If
   Next iRow

   conFlxGrid.TextMatrix(iNewRow, iColID) = "X"
   conFlxGrid.row = iNewRow

   For iCol = iColPaint To conFlxGrid.Cols - 1
      conFlxGrid.col = iCol
      conFlxGrid.CellBackColor = RGB(174, 179, 233)
   Next iCol
End Sub
Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
    Me.Height = 5325
    Me.Width = 18630
    
    FraReconciliation.Left = 10395
    FraReconciliation.Top = 90
    
    fraManualDemand.Left = 10395
    fraManualDemand.Top = 90
    
    fraVatamount.Left = 10395
    fraVatamount.Top = 90
    
    fraChecksum.Left = 10395
    fraChecksum.Top = 90
    
    fraReport.Left = 10395
    fraReport.Top = 90
    
    fraCheckSumPayment.Left = 10395
    fraCheckSumPayment.Top = 90
    
    frafixTransactions.Left = 10395
    frafixTransactions.Top = 90
    cmdRollbackBreconciliation.Enabled = False
    
    Call configflxissueList
    Call loadflxIssueList           'This is the main fucntion for loading the list of issues
    cboIssue.AddItem "Convert Demand Manual to Auto (WPM)"
    cboIssue.AddItem "VAT amount in demandsplit is not in total Amount (WESTGATE)"
    cboIssue.AddItem "Rollback Bank reconciliation"
    cboIssue.AddItem "Payment Edit, Receipt Edit was not posting Data to NL FIX"
    'remming this on 20181128 this issue is same as issue 673
    'cboIssue.AddItem "Booking Batch Receipt with lessee filter not updating lessee balance correctly"
    
    cboIssue.AddItem "Show checksum report of receipt allocation"
    cboIssue.ListIndex = 0
    
    '#
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    '#
    Dim szSQL As String
    Dim adoRst As New ADODB.Recordset
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
   End If
   adoRst.Close
    szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
                  "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
              "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
              "WHERE N.ClientID = CB.CLIENT_ID AND CB.NominalCode = N.Code AND " & _
                  "CB.CLIENT_ID = '" & txtClientList.Tag & "' " & _
              "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
     adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If Not adoRst.EOF Then
        txtBC.Tag = adoRst.Fields("BNC").Value
        txtBC.text = adoRst.Fields("BNC").Value
   End If
   adoRst.Close
 
   '#
    Dim ReConDates()        As String
   
    szSQL = "SELECT StatementDate, BankCode " & _
           "FROM tlbBankReconClosingBal " & _
           "WHERE BankCode = '" & txtBC.Tag & "' " & _
           "AND ClientID = '" & txtClientList.Tag & "' GROUP BY StatementDate, BankCode " & _
           "ORDER BY StatementDate DESC;"
   
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim ReConDates(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
      For j = 0 To TotalCol
         If Not IsNull(adoRst.Fields(j).Value) And adoRst.Fields(j).Value <> "" Then
            ReConDates(j, i) = adoRst.Fields(j).Value
         End If
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   cboCurStDt.Column() = ReConDates()
  

   adoRst.Close
   adoconn.Close
   '#
   Call WheelHook(Me.hWnd)
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdFixTransactions.SetFocus
    End If
    DigitTextKeyPress txtEnd, KeyAscii
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEnd.SetFocus
    End If
    DigitTextKeyPress txtStart, KeyAscii
End Sub
