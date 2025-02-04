VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGLU 
   BackColor       =   &H00FFEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Lease Update"
   ClientHeight    =   12300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGLU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12300
   ScaleMode       =   0  'User
   ScaleWidth      =   15105
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   6
      Left            =   7065
      TabIndex        =   45
      Top             =   8010
      Width           =   6375
      Begin VB.CheckBox chkAll 
         Caption         =   "Select All Lease"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtGLU 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFF0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   3480
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxGLU 
         Height          =   1815
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblTotalSelectedLease 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   5640
         TabIndex        =   59
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   19
         Left            =   3720
         TabIndex        =   58
         Top             =   3000
         Visible         =   0   'False
         Width           =   1500
         VariousPropertyBits=   276824083
         Caption         =   "Total Selected Lease:"
         Size            =   "2646;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   3525
         ForeColor       =   128
         VariousPropertyBits=   8388627
         Caption         =   "Select Lease to update:"
         Size            =   "6218;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Index           =   17
         Left            =   1800
         TabIndex        =   55
         Top             =   120
         Width           =   4425
         VariousPropertyBits=   276824083
         Size            =   "7805;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   615
         Index           =   16
         Left            =   5160
         TabIndex        =   52
         Top             =   480
         Width           =   675
         VariousPropertyBits=   276824083
         Caption         =   "New yearly value"
         Size            =   "1191;1085"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   555
         Index           =   15
         Left            =   3720
         TabIndex        =   51
         Top             =   480
         Width           =   840
         VariousPropertyBits=   276824083
         Caption         =   "Current yearly value"
         Size            =   "1482;979"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   14
         Left            =   1680
         TabIndex        =   50
         Top             =   840
         Width           =   855
         VariousPropertyBits=   276824083
         Caption         =   "Lease Name"
         Size            =   "1508;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   49
         Top             =   840
         Width           =   750
         VariousPropertyBits=   276824083
         Caption         =   "Unit Name"
         Size            =   "1323;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   255
         VariousPropertyBits=   276824083
         Caption         =   "ID"
         Size            =   "450;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   11
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   1290
         VariousPropertyBits=   276824083
         Caption         =   "Property Name:"
         Size            =   "2275;688"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   7
      Left            =   0
      TabIndex        =   38
      Top             =   5160
      Width           =   6375
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Run Preview Report"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   600
         Width           =   2895
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   8
         Left            =   600
         TabIndex        =   41
         Top             =   2040
         Width           =   5025
         ForeColor       =   192
         VariousPropertyBits=   276824083
         Caption         =   "TO SKIP THE PREVIEW REPORT PLEASE CLICK 'FINISH'. THIS WILL APPLY THE CHANGES FOR IMMEDIATE USE."
         Size            =   "8864;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   7
         Left            =   600
         TabIndex        =   40
         Top             =   1320
         Width           =   4950
         ForeColor       =   32768
         VariousPropertyBits=   276824083
         Caption         =   "IT IS RECOMMENDED THAT YOU VIEW THE PREVIEW REPORT BEFORE IMPLEMENTING YOUR CHANGES."
         Size            =   "8731;688"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   6
         Left            =   600
         TabIndex        =   39
         Top             =   600
         Width           =   1515
         VariousPropertyBits=   276824083
         Caption         =   "Preview the changes:"
         Size            =   "2672;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   5
      Left            =   6750
      TabIndex        =   13
      Top             =   4995
      Width           =   6375
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3225
         TabIndex        =   25
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0080C0FF&
         Height          =   2295
         Index           =   1
         Left            =   480
         Top             =   480
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Height          =   2295
         Index           =   0
         Left            =   480
         Top             =   480
         Width           =   4095
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   56
         ToolTipText     =   "Replace existing value to the new amount."
         Top             =   2760
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Specific Amt."
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   540
         Index           =   5
         Left            =   600
         TabIndex        =   44
         ToolTipText     =   "Manually replace individual amounts with new values"
         Top             =   2760
         Width           =   3495
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "6165;952"
         Value           =   "0"
         Caption         =   "Manually enter individual value amounts"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   10
         Left            =   4800
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         Width           =   795
         ForeColor       =   16512
         VariousPropertyBits=   276824083
         Caption         =   "Opt - Index"
         Size            =   "1402;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   31
         Top             =   2400
         Width           =   1725
         VariousPropertyBits=   276824083
         Caption         =   "Please enter the amount"
         Size            =   "3043;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   30
         ToolTipText     =   "Replace existing value to the new amount."
         Top             =   1920
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "New Amt."
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   540
         Index           =   4
         Left            =   600
         TabIndex        =   24
         ToolTipText     =   "Replace existing value or percentage with a new amount."
         Top             =   1875
         Width           =   2850
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "5027;952"
         Value           =   "0"
         Caption         =   "Apply a new percentage or Value"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   29
         ToolTipText     =   "Decrease existing charges by deducting an amount."
         Top             =   1560
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Dec. by value"
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   345
         Index           =   3
         Left            =   600
         TabIndex        =   23
         ToolTipText     =   "Decrease existing charges by deducting an amount."
         Top             =   1560
         Width           =   1695
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2990;609"
         Value           =   "0"
         Caption         =   "Decrease by value"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   28
         ToolTipText     =   "Decrease existing charges by a percentage."
         Top             =   1200
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Dec. by %"
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   345
         Index           =   2
         Left            =   600
         TabIndex        =   22
         ToolTipText     =   "Decrease existing charges by a percentage."
         Top             =   1200
         Width           =   2370
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4180;609"
         Value           =   "0"
         Caption         =   "Decrease by percentage (%)"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   27
         ToolTipText     =   "Increase existing charges by adding an amount."
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Inc. by value"
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   21
         ToolTipText     =   "Increase existing charges by adding an amount."
         Top             =   840
         Width           =   1620
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2857;609"
         Value           =   "0"
         Caption         =   "Increase by value"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblMethodTypeHelp 
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   26
         ToolTipText     =   "Increase existing charges by a percentage."
         Top             =   480
         Visible         =   0   'False
         Width           =   1005
         ForeColor       =   16512
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Inc. by %"
         Size            =   "1773;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optIncPCT 
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   20
         ToolTipText     =   "Increase existing charges by a percentage."
         Top             =   480
         Width           =   2295
         VariousPropertyBits=   1015023635
         BackColor       =   16772846
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "4048;609"
         Value           =   "1"
         Caption         =   "Increase by percentage (%)"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
         VariousPropertyBits=   276824083
         Caption         =   "Method Type:"
         Size            =   "1720;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   8550
      Width           =   6375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChargingMethod 
         Height          =   2490
         Left            =   225
         TabIndex        =   63
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4392
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label Label1 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2100
         VariousPropertyBits=   276824083
         Caption         =   "Select Charging Method:"
         Size            =   "3704;476"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   3
      Left            =   7455
      TabIndex        =   11
      Top             =   2400
      Width           =   6375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypes 
         Height          =   2490
         Left            =   225
         TabIndex        =   62
         Top             =   585
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4392
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label Label1 
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2280
         VariousPropertyBits=   276824083
         Caption         =   "Select Demand Types:"
         Size            =   "4022;397"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1575
      Index           =   8
      Left            =   0
      TabIndex        =   15
      Top             =   3480
      Width           =   6375
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   375
         Left            =   3540
         TabIndex        =   36
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   1
         X1              =   0
         X2              =   6840
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   6840
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   0
         ScaleHeight     =   825
         ScaleWidth      =   6345
         TabIndex        =   1
         Top             =   0
         Width           =   6375
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the Global Changes"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   3975
         End
      End
      Begin MSForms.Label Label1 
         Height          =   390
         Index           =   9
         Left            =   240
         TabIndex        =   42
         Top             =   3000
         Width           =   5610
         ForeColor       =   255
         BackColor       =   16777215
         VariousPropertyBits=   276824091
         Caption         =   "PLEASE ENSURE ALL OTHER USERS ARE LOGGED OUT OF PRESTIGE BEFORE RUNNING THIS PROCEDURE"
         Size            =   "9895;688"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   8
         Top             =   2640
         Width           =   2895
         ForeColor       =   128
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "- Apply the changes for immediate use"
         Size            =   "5106;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Top             =   2400
         Width           =   4455
         ForeColor       =   128
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "- Preview the changes before implementing them on your data"
         Size            =   "7858;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   6
         Top             =   2160
         Width           =   2775
         ForeColor       =   128
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "- Determine how you want to change it"
         Size            =   "4895;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   2535
         ForeColor       =   128
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "- Choose what you want to change"
         Size            =   "4471;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
         ForeColor       =   128
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "This form will help you to:"
         Size            =   "3413;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   5895
         BackColor       =   16772846
         VariousPropertyBits=   8388627
         Caption         =   "Use this form to globally amend Rent Charges, Service Charge and Insurance Charges on your leases."
         Size            =   "10398;873"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   2
      Left            =   6960
      TabIndex        =   10
      Top             =   1320
      Width           =   6375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperty 
         Height          =   2490
         Left            =   225
         TabIndex        =   61
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4392
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1140
         VariousPropertyBits=   276824083
         Caption         =   "Select Property:"
         Size            =   "2011;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3255
      Index           =   1
      Left            =   6570
      TabIndex        =   9
      Top             =   240
      Width           =   6375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
         Height          =   2490
         Left            =   225
         TabIndex        =   60
         Top             =   585
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   4392
         _Version        =   393216
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
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   960
         VariousPropertyBits=   276824083
         Caption         =   "Select Client:"
         Size            =   "1693;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label lblFrameIndex 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "1"
      Height          =   195
      Left            =   6600
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmGLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_NUMBER_FRAME_INDEX = 8
Private bPreviewShowed As Boolean
Private szDemandCharges As String
Dim iCurRow As Integer
Dim baChangesMade() As Boolean
Dim szUndoText As String, iflxGLUCol As Integer

Private Sub cmdBack_Click()
   lblFrameIndex.Caption = Val(lblFrameIndex.Caption) - 1
   Frame1(Val(lblFrameIndex.Caption) - 1).ZOrder 0
   cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
   cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
   cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
   lblTotalSelectedLease.Caption = "0"
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Are you sure you want to cancel the routine?", vbQuestion + vbYesNo, "Global Lease Update") = vbNo Then Exit Sub
   Unload Me
End Sub

Private Sub cmdNext_Click()
   If Val(lblFrameIndex.Caption) = 1 Then flxClients.SetFocus

   If Val(lblFrameIndex.Caption) = 2 Then
      If flxClients.TextMatrix(flxClients.row, 0) = "" Or flxClients.TextMatrix(flxClients.row, 1) = "" Or IsNull(flxClients.TextMatrix(flxClients.row, 1)) Then
         MsgBox "Please select a client from the list.", vbCritical + vbOKOnly, "Select Client"
         flxClients.SetFocus
         Exit Sub
      End If

      LoadProperties
      flxProperty.SetFocus
   End If

   If Val(lblFrameIndex.Caption) = 3 Then
      If flxProperty.TextMatrix(flxProperty.row, 0) = "" Or flxProperty.TextMatrix(flxProperty.row, 1) = "" Or IsNull(flxProperty.TextMatrix(flxProperty.row, 1)) Then
         MsgBox "Please select a property from the list.", vbCritical + vbOKOnly, "Select Property"
         flxProperty.SetFocus
         Exit Sub
      End If

      LoadDemandType
      flxDemandTypes.SetFocus
   End If

   If Val(lblFrameIndex.Caption) = 4 Then
      If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "" Or flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) = "" Or IsNull(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) Then
         MsgBox "Please select a demand type from the list.", vbCritical + vbOKOnly, "Select Demand Type"
         flxDemandTypes.SetFocus
         Exit Sub
      End If
      LoadChargingMethod
      flxChargingMethod.SetFocus
   End If

   If Val(lblFrameIndex.Caption) = 5 Then
      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 0) = "" Or flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) = "" Or IsNull(flxChargingMethod.TextMatrix(flxChargingMethod.row, 1)) Then
         MsgBox "Please select a charging method from the list.", vbCritical + vbOKOnly, "Select Charging Method"
         flxChargingMethod.SetFocus
         Exit Sub
      End If
   End If

   If Val(lblFrameIndex.Caption) = 6 Then
      If optIncPCT(5).Value Then
         txtValue.Enabled = False
         chkAll.Enabled = False
         Label1(18).Caption = "Enter specific amount for eash lease:"
         Label1(19).Visible = False
         lblTotalSelectedLease.Visible = False
      Else
         If txtValue.text = "" Or Val(txtValue.text) <= 0 Then
            MsgBox "Please enter the value.", vbCritical + vbOKOnly, "Charging Amount"
            txtValue.Enabled = True
            txtValue.SetFocus
            SelTxtInCtrl txtValue
            chkAll.Enabled = True
            Exit Sub
         Else
            Label1(18).Caption = "Select Lease to update:"
            Label1(19).Visible = True
            lblTotalSelectedLease.Visible = True
         End If
      End If
      optIncPCT(0).SetFocus
      cmdPreview.SetFocus
      LoadLeaseList
      flxGLU.SetFocus
   End If

   Frame1(Val(lblFrameIndex.Caption)).Top = 860
   Frame1(Val(lblFrameIndex.Caption)).Left = Frame1(Val(lblFrameIndex.Caption) - 1).Left
   Frame1(Val(lblFrameIndex.Caption)).ZOrder 0
   lblFrameIndex.Caption = Val(lblFrameIndex.Caption) + 1

   cmdBack.Enabled = IIf(Val(lblFrameIndex.Caption) > 1, True, False)
   cmdFinish.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, True, False)
   cmdNext.Enabled = IIf(Val(lblFrameIndex.Caption) > MAX_NUMBER_FRAME_INDEX - 1, False, True)
End Sub

Private Sub LoadChargingMethod()
   Dim szSQL As String, r As Integer, iDemandCategory As Byte
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szHeader As String
   flxChargingMethod.Clear

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT DT.CategoryCode " & _
           "FROM DemandTypes AS DT " & _
           "WHERE DT.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ";"
'Debug.Print szSQL
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iDemandCategory = adoRST.Fields.Item(0).Value
   adoRST.Close
'issue 731 Glabal lease update not working for insurence
'   If iDemandCategory = 3 Then      'demand type is Insurance
'      szSQL = "SELECT CODE, VALUE AS V " & _
'              "FROM SecondaryCode AS SC " & _
'              "WHERE PrimaryCode = 'ICRG';"
'   Else
'      szSQL = "SELECT ChargingMethodID AS CODE, ChargingMethod AS V " & _
'           "FROM CHARGINGMETHOD;"
'   End If
  szSQL = "SELECT ChargingMethodID AS CODE, ChargingMethod AS V " & _
           "FROM CHARGINGMETHOD;"

   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   szHeader$ = "|< ID|< Charging Method"
   flxChargingMethod.FormatString = szHeader$
   flxChargingMethod.ColWidth(0) = 200
   flxChargingMethod.ColWidth(1) = 1500
   flxChargingMethod.ColWidth(2) = 4000
   flxChargingMethod.Rows = 2
   flxChargingMethod.Cols = 3
   r = 1
   While Not adoRST.EOF
      flxChargingMethod.TextMatrix(r, 1) = adoRST.Fields.Item("CODE").Value
      flxChargingMethod.TextMatrix(r, 2) = adoRST.Fields.Item("V").Value
      flxChargingMethod.AddItem ""
      r = r + 1
      adoRST.MoveNext
   Wend
   flxChargingMethod.TextMatrix(r, 1) = "5"
   flxChargingMethod.TextMatrix(r, 2) = "ALL"
   flxChargingMethod.AddItem ""
   
   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub LoadDemandType()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szHeader As String
   flxDemandTypes.Clear

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT ID, TYPE " & _
           "FROM DEMANDTYPES " & _
           "WHERE PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' OR PropertyID = 'ALL';"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
   szHeader$ = "|< ID|< TYPE"
   flxDemandTypes.FormatString = szHeader$
   flxDemandTypes.ColWidth(0) = 200
   flxDemandTypes.ColWidth(1) = 1500
   flxDemandTypes.ColWidth(2) = 4000
   flxDemandTypes.Rows = 2
   flxDemandTypes.Cols = 3
   r = 1
   While Not adoRST.EOF
      flxDemandTypes.TextMatrix(r, 1) = adoRST.Fields.Item("ID").Value
      flxDemandTypes.TextMatrix(r, 2) = adoRST.Fields.Item("TYPE").Value
      flxDemandTypes.AddItem ""
      r = r + 1
      adoRST.MoveNext
   Wend

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub LoadProperties()
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   Dim szHeader As String
   flxProperty.Clear

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT PROPERTYID, PROPERTYNAME " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & flxClients.TextMatrix(flxClients.row, 1) & "';"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
   
    szHeader$ = "|< PROPERTY ID|< PROPERTY NAME"
    flxProperty.FormatString = szHeader$
    flxProperty.ColWidth(0) = 200
    flxProperty.ColWidth(1) = 1500
    flxProperty.ColWidth(2) = 4000
    flxProperty.Rows = 2
    flxProperty.Cols = 3
    r = 1
   While Not adoRST.EOF
      flxProperty.TextMatrix(r, 1) = adoRST.Fields.Item("PROPERTYID").Value
      flxProperty.TextMatrix(r, 2) = adoRST.Fields.Item("PROPERTYNAME").Value
      flxProperty.AddItem ""
      r = r + 1
      adoRST.MoveNext
   Wend

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
End Sub

Private Sub cmdFinish_Click()
   If MsgBox("Are you sure you want to run the routine?", vbQuestion + vbYesNo, "Global Lease Update") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection
   Dim adoFROM As New ADODB.Recordset, adoTO As New ADODB.Recordset
   Dim szSQL As String

'   connect to database
   adoConn.Open getConnectionString

   If Not bPreviewShowed Then CreatePreview

   If szDemandCharges = "RENT" Then
      adoConn.Execute "UPDATE LRentCharges, tblPrevGLU " & _
                      "Set LRentCharges.spare2 = tblPrevGLU.NewValue, " & _
                           "LRentCharges.BRTotal = tblPrevGLU.NewYrTotal, " & _
                           "LRentCharges.BRAmount = tblPrevGLU.NewDueEachPeriod " & _
                      "WHERE tblPrevGLU.ChildID = LRentCharges.RentCharges AND tblPrevGLU.Value > 0;"
   End If
   If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
      adoConn.Execute "UPDATE LServiceCharges, tblPrevGLU " & _
                      "Set LServiceCharges.CMFigure = tblPrevGLU.NewValue, " & _
                           "LServiceCharges.SCTotal = tblPrevGLU.NewYrTotal, " & _
                           "LServiceCharges.SCAmount = tblPrevGLU.NewDueEachPeriod " & _
                      "WHERE tblPrevGLU.ChildID = LServiceCharges.ServiceCharge AND tblPrevGLU.Value > 0;"
   End If
   If szDemandCharges = "INSURANCE" Then
      adoConn.Execute "UPDATE LInsuranceCharges, tblPrevGLU " & _
                      "Set LInsuranceCharges.ChargingFigure = tblPrevGLU.NewValue, " & _
                           "LInsuranceCharges.TotalYearlyInsurance = tblPrevGLU.NewYrTotal, " & _
                           "LInsuranceCharges.InsuranceEachPeriod = tblPrevGLU.NewDueEachPeriod " & _
                      "WHERE tblPrevGLU.ChildID = LInsuranceCharges.InsCharges AND tblPrevGLU.Value > 0;"
   End If

   adoConn.Close
   Set adoConn = Nothing

   MsgBox "The selected leases have been updated.", vbInformation + vbOKOnly, "Global Lease Update"
   Unload Me
End Sub

Private Sub CreatePreview()
   Dim szSQL As String, iCount As Integer, dTemp As Double
   Dim adoConn As New ADODB.Connection
   Dim adoFROM As New ADODB.Recordset, adoTO As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

'  Delete the temporary table to print the preview list
   szSQL = "DELETE * FROM tblPrevGLU;"
   adoConn.Execute szSQL

   szSQL = "SELECT * FROM tblPrevGLU;"
   adoTO.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   szDemandCharges = DemandCategory(adoConn, flxDemandTypes.TextMatrix(flxDemandTypes.row, 1))

   If szDemandCharges = "RENT" Then
      szSQL = "SELECT LD.LEASEID, LD.SAGEACCOUNTNUMBER AS SAN, " & _
                  "RC.BRFrequency AS FREQ, RC.BRDemandType AS DMDTYPE, " & _
                  "CM.ChargingMethod AS C_METHOD, RC.spare2 AS C_VALUE, " & _
                  "RC.BRTotal AS YR_TOTAL, RC.BRAmount AS DUE_PERIOD, " & _
                  "RC.RentChargeDept AS DEPT, U.UnitNumber as UnitNumber, " & _
                  "RC.RentCharges AS CHILDID, RC.spare1 AS C_METHOD_ID " & _
              "FROM LeaseDetails AS LD, Property AS P, Units AS U, " & _
                  "LRentCharges AS RC, ChargingMethod AS CM "
      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LD.LeaseID = RC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "RC.spare1 = '" & flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) & "' AND " & _
                     "RC.BRDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "VAL(RC.spare1) = CM.ChargingMethodID AND " & _
                     "ISNULL(RC.spare3);"
      Else
         szSQL = szSQL & _
                 "WHERE LD.LeaseID = RC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "RC.BRDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "VAL(RC.spare1) = CM.ChargingMethodID AND " & _
                     "ISNULL(RC.spare3);"
      End If
   End If

   If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
      szSQL = "SELECT LD.LEASEID, LD.SAGEACCOUNTNUMBER AS SAN, " & _
                  "SC.SCFrequency AS FREQ, SC.SCDemandType AS DMDTYPE, " & _
                  "CM.ChargingMethod AS C_METHOD, SC.CMFigure AS C_VALUE, " & _
                  "SC.SCTotal AS YR_TOTAL, SC.SCAmount AS DUE_PERIOD, " & _
                  "SC.ServiceChargeDept AS DEPT, U.UnitNumber as UnitNumber, " & _
                  "SC.ServiceCharge AS CHILDID, SC.ChargingMethod AS C_METHOD_ID " & _
              "FROM LeaseDetails AS LD, Property AS P, Units AS U, " & _
                  "LServiceCharges AS SC, ChargingMethod AS CM "

      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LD.LeaseID = SC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "SC.ChargingMethod = " & CInt(flxChargingMethod.TextMatrix(flxChargingMethod.row, 1)) & " AND " & _
                     "SC.SCDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "SC.ChargingMethod = CM.ChargingMethodID AND " & _
                     "ISNULL(SC.spare3);"
      Else
         szSQL = szSQL & _
                 "WHERE LD.LeaseID = SC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "SC.SCDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "SC.ChargingMethod = CM.ChargingMethodID AND " & _
                     "ISNULL(SC.spare3);"
      End If
   End If
'Debug.Print szSQL
   If szDemandCharges = "INSURANCE" Then
   ' szSQL = ""
      szSQL = "SELECT LD.LEASEID, LD.SAGEACCOUNTNUMBER AS SAN, " & _
                  "IC.InsuranceFrequency AS FREQ, IC.InsuranceDemandType AS DMDTYPE, " & _
                  "CM.ChargingMethod AS C_METHOD, IC.ChargingFigure AS C_VALUE, " & _
                  "IC.TotalYearlyInsurance AS YR_TOTAL, IC.InsuranceEachPeriod AS DUE_PERIOD, " & _
                  "IC.InsuranceDemandType AS DEPT, U.UnitNumber as UnitNumber, " & _
                  "IC.InsCharges AS CHILDID, IC.ChargingType AS C_METHOD_ID " & _
              "FROM LeaseDetails AS LD, Property AS P, Units AS U, LInsuranceCharges AS IC, ChargingMethod AS CM  "
      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LD.STATUS = TRUE AND " & _
                     "LD.LeaseID = IC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "IC.ChargingType = " & CInt(flxChargingMethod.TextMatrix(flxChargingMethod.row, 1)) & " AND " & _
                     "IC.InsuranceDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "IC.ChargingType = CM.ChargingMethodID AND " & _
                     "ISNULL(IC.spare3);"
                     
                     
      Else
         szSQL = szSQL & _
                 "WHERE LD.STATUS = TRUE AND " & _
                     "LD.LeaseID = IC.LeaseID AND " & _
                     "P.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "P.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "P.PropertyID = U.PropertyID AND " & _
                     "U.UnitNumber = LD.UnitNumber AND " & _
                     "IC.InsuranceDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "IC.ChargingType = CM.ChargingMethodID AND " & _
                     "ISNULL(IC.spare3);"
      End If
   End If
'Debug.Print szSQL
   adoFROM.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   iCount = 1
   With adoTO
      While Not adoFROM.EOF
         .AddNew
         .Fields.Item("ID").Value = iCount
         .Fields.Item("LEASEID").Value = adoFROM.Fields.Item("LEASEID").Value
         .Fields.Item("CHILDID").Value = adoFROM.Fields.Item("CHILDID").Value
         .Fields.Item("SAN").Value = adoFROM.Fields.Item("SAN").Value
         .Fields.Item("UNITNUMBER").Value = adoFROM.Fields.Item("UNITNUMBER").Value
         .Fields.Item("FREQ").Value = adoFROM.Fields.Item("FREQ").Value
         .Fields.Item("DMDTYPE").Value = adoFROM.Fields.Item("DMDTYPE").Value
         .Fields.Item("CHARGINGMETHOD").Value = adoFROM.Fields.Item("C_METHOD").Value
         .Fields.Item("METHODTYPE").Value = lblMethodTypeHelp(CInt(Label1(10))).Caption
         If Label1(10).Caption = 5 Then
            If flxGLU.TextMatrix(iCount, 1) = adoFROM.Fields.Item("LEASEID").Value Then
               .Fields.Item("VALUE").Value = flxGLU.TextMatrix(iCount, 5)
            End If
         Else
            If flxGLU.TextMatrix(iCount, 0) = "X" Then
               .Fields.Item("VALUE").Value = txtValue.text
            Else
               .Fields.Item("VALUE").Value = "0.00"
            End If
         End If
         .Fields.Item("EXVALUE").Value = Val(adoFROM.Fields.Item("C_VALUE").Value) 'CCur
         .Fields.Item("EXYRTOTAL").Value = Val(adoFROM.Fields.Item("YR_TOTAL").Value) 'CCur
         .Fields.Item("EXDUEEACHPERIOD").Value = Val(adoFROM.Fields.Item("DUE_PERIOD").Value) 'CCur

         If Label1(10).Caption = "0" Then                                                       'Increase by percentage (%)
            .Fields.Item("NEWVALUE").Value = Val(adoFROM.Fields.Item("C_VALUE").Value) * (Val(txtValue.text) + 100) / 100 'CCur
            If szDemandCharges = "RENT" Then
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 1 Then _
                  .Fields.Item("NEWYRTOTAL").Value = RCPricePerSqFoot(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                   .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                   dTemp)
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 2 Then _
                  .Fields.Item("NEWYRTOTAL").Value = RCPercentage(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                   .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                   adoFROM.Fields.Item("DEPT").Value, dTemp)
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 3 Then _
                  .Fields.Item("NEWYRTOTAL").Value = RCAnnual(adoConn, .Fields.Item("NEWVALUE").Value, _
                                                   .Fields.Item("FREQ").Value, dTemp)
            End If
            If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 1 Then _
                  .Fields.Item("NEWYRTOTAL").Value = SCPricePerSqFoot(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                   .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                   dTemp)
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 2 Then _
                  .Fields.Item("NEWYRTOTAL").Value = SCPercentage(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                   .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                   adoFROM.Fields.Item("DEPT").Value, dTemp)
               If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 3 Then _
                  .Fields.Item("NEWYRTOTAL").Value = SCAnnual(adoConn, .Fields.Item("NEWVALUE").Value, _
                                                   .Fields.Item("FREQ").Value, dTemp)
            
            End If
            If szDemandCharges = "INSURANCE" Then
               .Fields.Item("NEWYRTOTAL").Value = InsuranceUpdate(CInt(adoFROM.Fields.Item("C_METHOD_ID").Value), adoConn, _
                                                   .Fields.Item("UNITNUMBER").Value, _
                                                   .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                   adoFROM.Fields.Item("DEPT").Value, dTemp)
            End If
            .Fields.Item("NEWDUEEACHPERIOD").Value = dTemp
         End If

'        Increase by value
         If Label1(10).Caption = "1" Then _
            .Fields.Item("NEWVALUE").Value = Val(adoFROM.Fields.Item("C_VALUE").Value) + (Val(txtValue.text)) 'ccur
'        Decrease by percentage (%)
         If Label1(10).Caption = "2" Then _
            .Fields.Item("NEWVALUE").Value = Val(adoFROM.Fields.Item("C_VALUE").Value) * (100 - Val(txtValue.text)) / 100 'ccur
'        Decrease by value
         If Label1(10).Caption = "3" Then _
            .Fields.Item("NEWVALUE").Value = Val(adoFROM.Fields.Item("C_VALUE").Value) - (Val(txtValue.text)) 'ccur
'        Change to a fixed amount
         If Label1(10).Caption = "4" Then _
            .Fields.Item("NEWVALUE").Value = Val(txtValue.text) 'ccur
         If Label1(10).Caption = "5" Then _
            .Fields.Item("NEWVALUE").Value = Val(.Fields.Item("VALUE").Value) 'ccur

         If szDemandCharges = "RENT" Then
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 1 Then _
               .Fields.Item("NEWYRTOTAL").Value = RCPricePerSqFoot(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                dTemp)
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 2 Then _
               .Fields.Item("NEWYRTOTAL").Value = RCPercentage(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                adoFROM.Fields.Item("DEPT").Value, dTemp)
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 3 Then _
               .Fields.Item("NEWYRTOTAL").Value = RCAnnual(adoConn, .Fields.Item("NEWVALUE").Value, _
                                                .Fields.Item("FREQ").Value, dTemp)
         End If
         If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 1 Then _
               .Fields.Item("NEWYRTOTAL").Value = SCPricePerSqFoot(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                dTemp)
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 2 Then _
               .Fields.Item("NEWYRTOTAL").Value = SCPercentage(adoConn, .Fields.Item("UNITNUMBER").Value, _
                                                .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                adoFROM.Fields.Item("DEPT").Value, dTemp)
            If CInt(adoFROM.Fields.Item("C_METHOD_ID").Value) = 3 Then _
               .Fields.Item("NEWYRTOTAL").Value = SCAnnual(adoConn, .Fields.Item("NEWVALUE").Value, _
                                                .Fields.Item("FREQ").Value, dTemp)
         End If
         If szDemandCharges = "INSURANCE" Then
            .Fields.Item("NEWYRTOTAL").Value = InsuranceUpdate(CInt(adoFROM.Fields.Item("C_METHOD_ID").Value), adoConn, _
                                                .Fields.Item("UNITNUMBER").Value, _
                                                .Fields.Item("FREQ").Value, .Fields.Item("NEWVALUE").Value, _
                                                adoFROM.Fields.Item("DEPT").Value, dTemp)
         End If
         .Fields.Item("NEWDUEEACHPERIOD").Value = dTemp
         .Update
         adoFROM.MoveNext
         iCount = iCount + 1
      Wend
   End With

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Function InsuranceUpdate(ByVal C_METHOD As Byte, ByVal adoConn As ADODB.Connection, _
                                 ByVal szUnit As String, ByVal bFreq As Byte, ByVal dNewChargingValue As Double, _
                                 ByVal szRCDept As String, ByRef retDuePeriod As Double) As Double
   Dim Area As String, iPartOfYear As Integer
   Dim SQLStr1 As String
   Dim TotalInsurance As Double
   Dim adoRST As New ADODB.Recordset

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly
   iPartOfYear = CInt(adoRST.Fields.Item("PARTOFYEAR").Value)
   adoRST.Close

   If C_METHOD = 3 Then                                                                'ANNUAL CHARGES
      InsuranceUpdate = CDbl(dNewChargingValue)
      retDuePeriod = CDbl(dNewChargingValue) / iPartOfYear
   End If

   If C_METHOD = 2 Then                                                                'PERCENTAGE
      SQLStr1 = "SELECT Amount " & _
                "FROM GlobalInsurance, Units " & _
                "WHERE Units.PropertyID = GlobalInsurance.PropertyID " & _
                  "AND Units.UnitNumber = '" & szUnit & "' " & _
                  "AND GlobalInsurance.DemandType = " & CInt(szRCDept) & ";"
      adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

      If Not adoRST.EOF Then
         TotalInsurance = CDbl(adoRST.Fields.Item("Amount").Value)
      Else
         TotalInsurance = 0
      End If
      adoRST.Close

      InsuranceUpdate = TotalInsurance * (CDbl(dNewChargingValue) / 100)
      retDuePeriod = TotalInsurance * (CDbl(dNewChargingValue) / 100) / iPartOfYear
   End If
End Function

Private Function RCAnnual(ByVal adoConn As ADODB.Connection, ByVal szValue As String, ByVal bFreq As Byte, ByRef retDuePeriod As Double) As Double
   Dim SQLStr1  As String
   Dim adoRST As New ADODB.Recordset

   RCAnnual = CDbl(szValue)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
'Debug.Print SQLStr1
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = RCAnnual / CInt(adoRST.Fields.Item("PARTOFYEAR").Value)

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Function SCAnnual(ByVal adoConn As ADODB.Connection, ByVal szValue As String, ByVal bFreq As Byte, ByRef retDuePeriod As Double) As Double
   Dim SQLStr1  As String
   Dim adoRST As New ADODB.Recordset

   SCAnnual = CDbl(szValue)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"

   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = SCAnnual / CInt(adoRST.Fields.Item("PARTOFYEAR").Value)

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Function RCPercentage(ByVal adoConn As ADODB.Connection, ByVal szUnit As String, ByVal bFreq As Byte, ByVal dNewChargingValue As Double, ByVal szRCDept As String, ByRef retDuePeriod As Double) As Double
   Dim TotalRentCharge As String, SQLStr1 As String
   Dim adoRST As New ADODB.Recordset

   TotalRentCharge = GetGlobalTotalRC(adoConn, szUnit, szRCDept)

   RCPercentage = CDbl(TotalRentCharge) * (dNewChargingValue / 100)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = RCPercentage / CInt(adoRST.Fields.Item("PARTOFYEAR").Value)

   adoRST.Close
   Set adoRST = Nothing

   Exit Function

ErrorHander:
   adoRST.Close
   Set adoRST = Nothing
End Function

Private Function SCPercentage(ByVal adoConn As ADODB.Connection, ByVal szUnit As String, ByVal bFreq As Byte, ByVal dNewChargingValue As Double, ByVal szRCDept As String, ByRef retDuePeriod As Double) As Double
   Dim TotalServiceCharge As String, SQLStr1 As String
   Dim adoRST As New ADODB.Recordset

   TotalServiceCharge = GetGlobalTotalSC(adoConn, szUnit, szRCDept)

   SCPercentage = CDbl(TotalServiceCharge) * (dNewChargingValue / 100)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
   adoRST.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = SCPercentage / CInt(adoRST.Fields.Item("PARTOFYEAR").Value)

   adoRST.Close
   Set adoRST = Nothing

   Exit Function
ErrorHander:
   adoRST.Close
   Set adoRST = Nothing
End Function

Private Function RCPricePerSqFoot(ByVal adoConn As ADODB.Connection, ByVal szUnit As String, ByVal bFreq As Byte, ByVal dNewChargingValue As Double, ByRef retDuePeriod As Double) As Double
   Dim Area As String, SQLStr1 As String
   Dim rstSQL As New ADODB.Recordset

   Area = GetUnitTA_ADODB(adoConn, szUnit)

   RCPricePerSqFoot = Area * CDbl(dNewChargingValue)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
   rstSQL.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = RCPricePerSqFoot / CInt(rstSQL.Fields.Item("PARTOFYEAR").Value)

   rstSQL.Close
   Set rstSQL = Nothing
End Function

Private Function SCPricePerSqFoot(ByVal adoConn As ADODB.Connection, ByVal szUnit As String, ByVal bFreq As Byte, ByVal dNewChargingValue As Double, ByRef retDuePeriod As Double) As Double
   Dim Area As String, SQLStr1 As String
   Dim rstSQL As New ADODB.Recordset

   Area = GetUnitTA_ADODB(adoConn, szUnit)

   SCPricePerSqFoot = Area * CDbl(dNewChargingValue)

   SQLStr1 = "SELECT PARTOFYEAR " & _
             "FROM FREQUENCIES " & _
             "WHERE ID = " & bFreq & ";"
   rstSQL.Open SQLStr1, adoConn, adOpenStatic, adLockReadOnly

   retDuePeriod = SCPricePerSqFoot / CInt(rstSQL.Fields.Item("PARTOFYEAR").Value)

   rstSQL.Close
   Set rstSQL = Nothing
End Function

Private Function DemandCategory(adoConn As ADODB.Connection, szDemandType As String) As String
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
'Fixed by anol 02 Nov 2015,Cbye was changed to cint due to consistency datacheck
   szSQL = "SELECT SECONDARYCODE.VALUE " & _
           "FROM DEMANDTYPES, SECONDARYCODE " & _
           "WHERE DEMANDTYPES.ID = " & CInt(szDemandType) & " AND " & _
               "DEMANDTYPES.CATEGORYCODE = CINT(SECONDARYCODE.CODE) AND " & _
               "SECONDARYCODE.PRIMARYCODE = 'DCTG';"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

   DemandCategory = adoRST.Fields.Item("VALUE").Value

   adoRST.Close
   Set adoRST = Nothing
End Function

Private Sub cmdPreview_Click()
    Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport
   CreatePreview
   bPreviewShowed = True
   Me.Hide
'   ShowReport App.Path & szReportPath & "\GLUpdate.rpt"
    Set Report = reportApp.OpenReport(App.Path & szReportPath & "\GLUpdate.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   Me.Show
End Sub

Private Sub cmdPreview_LostFocus()
   

   FocusControl cmdFinish
End Sub

Private Sub flxChargingMethod_Click()
   Dim iSel As Integer
   Dim i    As Integer
   Dim K As Integer
   Call SelectOnly1RowFlxGrid(flxChargingMethod, flxChargingMethod.row, 0)
End Sub

Private Sub flxClients_Click()
   Dim iSel As Integer
   Dim i    As Integer
   Dim K As Integer
   Call SelectOnly1RowFlxGrid(flxClients, flxClients.row, 0)
End Sub

Private Sub flxDemandTypes_Click()
   Dim iSel As Integer
   Dim i    As Integer
   Dim K As Integer
   Call SelectOnly1RowFlxGrid(flxDemandTypes, flxDemandTypes.row, 0)
End Sub

Private Sub flxProperty_Click()
   Dim iSel As Integer
   Dim i    As Integer
   Dim K As Integer
   Call SelectOnly1RowFlxGrid(flxProperty, flxProperty.row, 0)
End Sub

Private Sub Form_Load()
   Dim szHeader As String
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.Width = 6435
   Me.Height = 5500
   Label1(10).Caption = "0"
   Me.BackColor = MODULEBACKCOLOR
   Frame1(0).BackColor = Me.BackColor
   Frame1(1).BackColor = Me.BackColor
   Frame1(2).BackColor = Me.BackColor
   Frame1(3).BackColor = Me.BackColor
   Frame1(4).BackColor = Me.BackColor
   Frame1(5).BackColor = Me.BackColor
   Frame1(6).BackColor = Me.BackColor
   Frame1(8).BackColor = Me.BackColor
   Frame1(7).BackColor = Me.BackColor
   chkAll.BackColor = Me.BackColor
   Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT;"
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly
   szHeader$ = "|< CLIENT ID|< CLIENT NAME"
   flxClients.FormatString = szHeader$
   flxClients.ColWidth(0) = 200
   flxClients.ColWidth(1) = 1500
   flxClients.ColWidth(2) = 4000
   flxClients.Rows = 2
   flxClients.Cols = 3
   r = 1
   While Not adoRST.EOF
      flxClients.TextMatrix(r, 1) = adoRST.Fields.Item("CLIENTID").Value
      flxClients.TextMatrix(r, 2) = adoRST.Fields.Item("CLIENTNAME").Value
      flxClients.AddItem ""
      r = r + 1
      adoRST.MoveNext
   Wend

   adoRST.Close
   adoConn.Close
   Set adoRST = Nothing
   Set adoConn = Nothing
   configureFlxGLU
   bPreviewShowed = False
   Call WheelHook(Me.hWnd)
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
'
'  ' Scroll was not handled by any controls, so treat as a general message send to the form
'  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame1(Index).MousePointer = vbArrow
End Sub

Private Sub lstChargingMethod_LostFocus()
   cmdNext.SetFocus
End Sub

Private Sub lstClients_LostFocus()
   cmdNext.SetFocus
End Sub

Private Sub lstDemandTypes_LostFocus()
   cmdNext.SetFocus
End Sub

Private Sub lstProperties_LostFocus()
   cmdNext.SetFocus
End Sub

Private Sub optIncPCT_Click(Index As Integer)
   Label1(10).Caption = Index
   If optIncPCT(5).Value Then
      txtValue.text = ""
      txtValue.Enabled = False
   Else
      txtValue.text = ""
      txtValue.Enabled = True
   End If
End Sub

Private Sub optIncPCT_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.stbStatusBar.Panels(1).text = optIncPCT(Index).ToolTipText
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Picture1.MousePointer = vbArrow
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtValue, KeyAscii, 8
   If KeyAscii = 13 Then
        cmdNext.SetFocus
   End If
End Sub

'Private Sub txtValue_LostFocus()
'   cmdNext.SetFocus
'End Sub

Private Sub configureFlxGLU()
   Dim szHeader As String

   flxGLU.Cols = 5
   flxGLU.Clear
   szHeader$ = "|<ID|<Unit Name|<Lease Name|<Current yearly value|<New yearly value"
   flxGLU.FormatString = szHeader$
   flxGLU.ColWidth(0) = 300
   flxGLU.ColWidth(1) = 0          'ID
   flxGLU.ColWidth(2) = 2000       'Unit Name
   flxGLU.ColWidth(3) = 1700       'Lease Name
   flxGLU.ColWidth(4) = 900        'Current yearly value
   flxGLU.ColWidth(5) = 900        'New yearly value
    flxGLU.ColWidth(6) = 0        'C_value
   flxGLU.Rows = 2

   Label1(13).Left = 1200
   Label1(14).Left = 2800
   Label1(15).Left = 4100
   Label1(16).Left = 5000

   flxGLU.RowHeight(0) = 0
End Sub

Private Sub LoadLeaseList()
   Dim szSQL As String, iRow As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset
   ReDim baChangesMade(flxGLU.Rows) As Boolean
    chkAll.Value = 0
'   connect to database
   adoConn.Open getConnectionString
   
   szDemandCharges = DemandCategory(adoConn, flxDemandTypes.TextMatrix(flxDemandTypes.row, 1))

   If szDemandCharges = "RENT" Then
      szSQL = "SELECT LeaseDetails.LeaseID, LeaseDetails.SAGEACCOUNTNUMBER AS SAN, LeaseDetails.CompanyName, " & _
                  "RC.BRFrequency AS FREQ, RC.BRDemandType AS DMDTYPE, " & _
                  "RC.spare1 AS C_METHOD, RC.spare2 AS C_VALUE, " & _
                  "RC.BRTotal, RC.BRAmount AS DUE_PERIOD, " & _
                  "RC.RentChargeDept AS DEPT, UNITS.UnitNumber as UnitNumber, Units.UnitName, " & _
                  "RC.RentCharges AS CHILDID, Property.PropertyName " & _
              "FROM LeaseDetails, Property, Units, LRentCharges AS RC "
      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LeaseDetails.LeaseID = RC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "RC.spare1 = '" & flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) & "' AND " & _
                     "RC.BRDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(RC.spare3);"
      Else
         szSQL = szSQL & _
                 "WHERE LeaseDetails.LeaseID = RC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "RC.BRDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(RC.spare3);"
      End If
   End If

   If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
      szSQL = "SELECT LeaseDetails.LeaseID, LeaseDetails.SAGEACCOUNTNUMBER AS SAN, LeaseDetails.CompanyName, " & _
                  "SC.SCFrequency AS FREQ, SC.SCDemandType AS DMDTYPE, " & _
                  "SC.ChargingMethod AS C_METHOD, SC.CMFigure AS C_VALUE, " & _
                  "SC.SCTotal, SC.SCAmount AS DUE_PERIOD, " & _
                  "SC.ServiceChargeDept AS DEPT, UNITS.UnitNumber as UnitNumber, Units.UnitName, " & _
                  "SC.ServiceCharge AS CHILDID, Property.PropertyName " & _
              "FROM LeaseDetails, Property, Units, LServiceCharges AS SC "

      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LeaseDetails.LeaseID = SC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "SC.ChargingMethod = " & CInt(flxChargingMethod.TextMatrix(flxChargingMethod.row, 1)) & " AND " & _
                     "SC.SCDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(SC.spare3);"
      Else
         szSQL = szSQL & _
                 "WHERE LeaseDetails.LeaseID = SC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "SC.SCDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(SC.spare3);"
      End If
   End If

   If szDemandCharges = "INSURANCE" Then
      szSQL = "SELECT LeaseDetails.LeaseID, LeaseDetails.SAGEACCOUNTNUMBER AS SAN, " & _
                  "IC.InsuranceFrequency AS FREQ, IC.InsuranceDemandType AS DMDTYPE, LeaseDetails.CompanyName, " & _
                  "IC.ChargingType AS C_METHOD, IC.ChargingFigure AS C_VALUE, " & _
                  "IC.TotalYearlyInsurance, IC.InsuranceEachPeriod AS DUE_PERIOD, " & _
                  "IC.InsuranceDemandType AS DEPT, UNITS.UnitNumber as UnitNumber, Units.UnitName, " & _
                  "IC.InsCharges AS CHILDID, Property.PropertyName " & _
              "FROM LeaseDetails, Property, Units, LInsuranceCharges AS IC "
      If flxChargingMethod.TextMatrix(flxChargingMethod.row, 1) <> "ALL" Then
         szSQL = szSQL & _
                 "WHERE LeaseDetails.STATUS = TRUE AND " & _
                     "LeaseDetails.LeaseID = IC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "IC.ChargingType = " & CInt(flxChargingMethod.TextMatrix(flxChargingMethod.row, 1)) & " AND " & _
                     "IC.InsuranceDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(IC.spare3);"
      Else
         szSQL = szSQL & _
                 "WHERE LeaseDetails.STATUS = TRUE AND " & _
                     "LeaseDetails.LeaseID = IC.LeaseID AND " & _
                     "Property.ClientID = '" & flxClients.TextMatrix(flxClients.row, 1) & "' AND " & _
                     "Property.PropertyID = '" & flxProperty.TextMatrix(flxProperty.row, 1) & "' AND " & _
                     "Property.PropertyID = UNITS.PropertyID AND " & _
                     "UNITS.UnitNumber = LeaseDetails.UnitNumber AND " & _
                     "IC.InsuranceDemandType = " & CInt(flxDemandTypes.TextMatrix(flxDemandTypes.row, 1)) & " AND " & _
                     "ISNULL(IC.spare3);"
      End If
   End If
  
   adoRST.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'Debug.Print szSQL
   
   With adoRST
      iRow = 1
      flxGLU.Clear
      flxGLU.Rows = 2
      While Not .EOF
         Label1(17).Caption = .Fields.Item("PropertyName").Value
         If Not .EOF Then flxGLU.AddItem ""
         flxGLU.TextMatrix(iRow, 1) = .Fields.Item("LeaseID").Value
         flxGLU.TextMatrix(iRow, 2) = .Fields.Item("UnitName").Value
         flxGLU.TextMatrix(iRow, 3) = .Fields.Item("CompanyName").Value
         flxGLU.TextMatrix(iRow, 6) = .Fields.Item("C_Value").Value
         If szDemandCharges = "RENT" Then
            flxGLU.TextMatrix(iRow, 4) = Format(.Fields.Item("BRTotal").Value, "0.00")
            If Label1(10).Caption = 5 Then
               flxGLU.TextMatrix(iRow, 5) = CCur(.Fields.Item("BRTotal").Value)
            End If
         Else
            If szDemandCharges = "SERVICE CHARGE" Or szDemandCharges = "OTHER" Then
               flxGLU.TextMatrix(iRow, 4) = Format(.Fields.Item("SCTotal").Value, "0.00")
               'fixed by anol 02 Nov 2015 added val
               If Val(Label1(10).Caption) = 5 Then
                  flxGLU.TextMatrix(iRow, 5) = Format(.Fields.Item("SCTotal").Value, "0.00")
               End If
            Else
               If szDemandCharges = "INSURANCE" Then
                  flxGLU.TextMatrix(iRow, 4) = Format(.Fields.Item("TotalYearlyInsurance").Value, "0.00")
                  If Label1(10).Caption = 5 Then
                     flxGLU.TextMatrix(iRow, 5) = CCur(.Fields.Item("TotalYearlyInsurance").Value)
                  End If
               End If
            End If
         End If
         iRow = iRow + 1
      .MoveNext
      Wend
      .Close
   End With
   
   Set adoRST = Nothing
   flxGLU.row = 0
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxGLU_Scroll()
   If txtGLU.Visible Then      'The grid is in edting mode
      txtGLU.text = szUndoText
      txtGLU.Enabled = True
      txtGLU.Visible = False
   End If
End Sub

Private Sub txtGLU_Click()
   SelTxtInCtrl txtGLU
End Sub

Private Sub txtGLU_GotFocus()
   SelTxtInCtrl txtGLU
   iCurRow = flxGLU.row
End Sub

Private Sub txtGLU_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then flxGLU.SetFocus
End Sub

Private Sub txtGLU_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtGLU, KeyAscii
End Sub

Private Sub txtGLU_LostFocus()
   txtGLU.text = Val(txtGLU.text)

   flxGLU.TextMatrix(iCurRow, 5) = txtGLU.text

   txtGLU.text = "0.00"

   txtGLU.Visible = False
   txtGLU.Enabled = True
End Sub

Private Sub flxGLU_Click()
   Dim i As Integer
   Dim iCurRowHeight As Integer
   If flxGLU.TextMatrix(flxGLU.row, 1) <> "" Then
      If Label1(10).Caption = 5 Then
         iflxGLUCol = 5
         flxGLU.col = iflxGLUCol

         szUndoText = flxGLU.TextMatrix(flxGLU.row, iflxGLUCol)

         txtGLU.Top = flxGLU.CellTop + flxGLU.Top
         txtGLU.Left = flxGLU.CellLeft + flxGLU.Left
         txtGLU.Width = flxGLU.ColWidth(iflxGLUCol) - 15
         txtGLU.Height = flxGLU.RowHeight(flxGLU.row) - 15
         txtGLU.text = flxGLU.TextMatrix(flxGLU.row, iflxGLUCol)
         txtGLU.Visible = True
         txtGLU.SetFocus
      Else
         Dim j As Integer
         j = Select1RowFlxGrid(flxGLU, flxGLU.row, 0)
         If flxGLU.TextMatrix(flxGLU.row, 0) = "X" Then
            'Increase by percentage (%)
            If Label1(10).Caption = "0" Then
               flxGLU.TextMatrix(flxGLU.row, 5) = Val(flxGLU.TextMatrix(flxGLU.row, 4)) * (Val(txtValue.text) + 100) / 100
            End If
   '        Increase by value
            If Label1(10).Caption = "1" Then _
               flxGLU.TextMatrix(flxGLU.row, 5) = Val(flxGLU.TextMatrix(flxGLU.row, 4)) + (Val(txtValue.text))
   '        Decrease by percentage (%)
            If Label1(10).Caption = "2" Then _
               flxGLU.TextMatrix(flxGLU.row, 5) = Val(flxGLU.TextMatrix(flxGLU.row, 4)) * (100 - Val(txtValue.text)) / 100
   '        Decrease by value
            If Label1(10).Caption = "3" Then _
               flxGLU.TextMatrix(flxGLU.row, 5) = Val(flxGLU.TextMatrix(flxGLU.row, 4)) - (Val(txtValue.text))
   '        Change to a fixed amount
            If Label1(10).Caption = "4" Then
                If flxGLU.TextMatrix(flxGLU.row, 4) = "" Or flxGLU.TextMatrix(flxGLU.row, 6) = "" Then
                Else
                    flxGLU.TextMatrix(flxGLU.row, 5) = Format((flxGLU.TextMatrix(flxGLU.row, 4) / flxGLU.TextMatrix(flxGLU.row, 6)) * Val(txtValue.text), "0.00")
                End If
                lblTotalSelectedLease.Caption = str(Val(lblTotalSelectedLease.Caption) + 1)
            End If
         Else
            flxGLU.TextMatrix(flxGLU.row, 5) = ""
            lblTotalSelectedLease.Caption = str(Val(lblTotalSelectedLease.Caption) - 1)
         End If
      End If
   End If
End Sub
Private Function Select1RowFlxGrid(conFlxGrid As Control, iSelRow As Integer, iXCol As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iXCol) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iXCol) = ""
      conFlxGrid.row = iSelRow
      For iRow = 1 To iXCol
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iRow
      Select1RowFlxGrid = -1
   Else
      conFlxGrid.TextMatrix(iSelRow, iXCol) = "X"

      conFlxGrid.row = iSelRow
      For iRow = 1 To conFlxGrid.Cols - 1
         conFlxGrid.col = iRow
         conFlxGrid.CellBackColor = RGB(255, 255, 255)
      Next iRow
      Select1RowFlxGrid = 1
   End If
End Function
Private Sub chkAll_Click()
   Dim i As Integer

   For i = 1 To flxGLU.Rows - 1
      flxGLU.TextMatrix(i, 0) = ""
      flxGLU.TextMatrix(i, 5) = ""
   Next i
   lblTotalSelectedLease.Caption = "0"

   If Not flxGLU.TextMatrix(1, 1) = "" Then
      If chkAll.Value Then
         For i = 1 To flxGLU.Rows - 1
            If Not flxGLU.TextMatrix(i, 1) = "" Then
               flxGLU.TextMatrix(i, 0) = "X"
               lblTotalSelectedLease.Caption = str(Val(lblTotalSelectedLease.Caption) + 1)
               'Increase by percentage (%)
               If Label1(10).Caption = "0" Then
                  'flxGLU.TextMatrix(i, 5) = CCur(flxGLU.TextMatrix(i, 4)) * (Val(Format(txtValue.text, "0.00")) + 100) / 100
                   flxGLU.TextMatrix(i, 5) = Val(flxGLU.TextMatrix(i, 4)) * (Val(txtValue.text) + 100) / 100
               End If
      '        Increase by value
               If Label1(10).Caption = "1" Then
'                  flxGLU.TextMatrix(i, 5) = CCur(flxGLU.TextMatrix(i, 4)) + (Val(Format(txtValue.text, "0.00")))
                   flxGLU.TextMatrix(i, 5) = Val(flxGLU.TextMatrix(i, 4)) + (Val(txtValue.text))
                   End If
      '        Decrease by percentage (%)
               If Label1(10).Caption = "2" Then
'                  flxGLU.TextMatrix(i, 5) = CCur(flxGLU.TextMatrix(i, 4)) * (100 - Val(Format(txtValue.text, "0.00"))) / 100
                    flxGLU.TextMatrix(i, 5) = Val(flxGLU.TextMatrix(i, 4)) * (100 - Val(txtValue.text)) / 100
                    End If
      '        Decrease by value
               If Label1(10).Caption = "3" Then
'                  flxGLU.TextMatrix(i, 5) = CCur(flxGLU.TextMatrix(i, 4)) - (Val(Format(txtValue.text, "0.00")))
                    flxGLU.TextMatrix(i, 5) = Val(flxGLU.TextMatrix(i, 4)) - (Val(txtValue.text))
                    End If
      '        Change to a fixed amount
               If Label1(10).Caption = "4" Then
'                  flxGLU.TextMatrix(i, 5) = CCur(Format(txtValue.text, "0.00"))
'                     flxGLU.TextMatrix(i, 5) = CCur(txtValue.text)
'                       flxGLU.TextMatrix(i, 5) = Val(txtValue.text)
'modified by anol 20170609
                        If flxGLU.TextMatrix(i, 4) = "" Or flxGLU.TextMatrix(i, 6) = "" Then
                            
                        Else
                            flxGLU.TextMatrix(i, 5) = Format((Val(flxGLU.TextMatrix(i, 4)) / flxGLU.TextMatrix(i, 6)) * Val(txtValue.text), "0.00")
                        End If
                       End If
              End If
         Next i
'      Else
'         For i = 1 To flxGLU.Rows - 1
'            flxGLU.TextMatrix(i, 0) = ""
'            flxGLU.TextMatrix(i, 5) = ""
'         Next i
'         lblTotalSelectedLease.Caption = "0"
      End If
   End If
End Sub
