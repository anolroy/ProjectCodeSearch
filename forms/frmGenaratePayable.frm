VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGenaratePayable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genarate Payable"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   18240
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   855
      TabIndex        =   35
      Top             =   7785
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
         TabIndex        =   43
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   4335
         Left            =   90
         TabIndex        =   42
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
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   5265
         TabIndex        =   41
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   2
         Left            =   5175
         TabIndex        =   40
         Top             =   180
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   39
         Top             =   195
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   38
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   315
         TabIndex        =   37
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   180
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape2 
         Height          =   5010
         Left            =   0
         Top             =   0
         Width           =   6585
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
   Begin VB.Frame Frame1 
      Height          =   5865
      Left            =   45
      TabIndex        =   21
      Top             =   1620
      Width           =   18150
      Begin VB.CommandButton cmdDeleteLine 
         Caption         =   "Clear"
         Height          =   420
         Left            =   14265
         TabIndex        =   54
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   52
         Top             =   5355
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   16515
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   5310
         Width           =   1215
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   5355
         MaxLength       =   20
         TabIndex        =   46
         Top             =   945
         Width           =   3465
      End
      Begin VB.TextBox txtPayableAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   14490
         MaxLength       =   10
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   405
         Width           =   1215
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   420
         Left            =   16155
         TabIndex        =   4
         Top             =   1080
         Width           =   1530
      End
      Begin VB.CommandButton cmdPayableTypes 
         Caption         =   ". ."
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
         Left            =   3375
         TabIndex        =   3
         Top             =   945
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtPayableTypes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   945
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.TextBox txtPayableDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "01/01/2000"
         Top             =   360
         Width           =   1170
      End
      Begin VB.CommandButton cmdFundListForCreatePI 
         Caption         =   ". ."
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
         Left            =   11385
         TabIndex        =   2
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtFundForPI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   8775
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   23
         Top             =   360
         Width           =   2565
      End
      Begin VB.CommandButton cmdProperty 
         Caption         =   ". ."
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
         Left            =   7020
         TabIndex        =   1
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtProperty 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   5355
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   22
         Top             =   360
         Width           =   1620
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayFees 
         Height          =   2940
         Left            =   90
         TabIndex        =   29
         Top             =   2205
         Width           =   17745
         _ExtentX        =   31300
         _ExtentY        =   5186
         _Version        =   393216
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
      Begin VB.Line Line2 
         X1              =   0
         X2              =   18090
         Y1              =   1710
         Y2              =   1710
      End
      Begin MSForms.Label Label6 
         Height          =   210
         Left            =   135
         TabIndex        =   53
         Top             =   5355
         Visible         =   0   'False
         Width           =   915
         VariousPropertyBits=   276824091
         Caption         =   "Description"
         Size            =   "1614;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   11
         Left            =   15030
         TabIndex        =   50
         Top             =   5355
         Width           =   1110
         VariousPropertyBits=   276824091
         Caption         =   "Total Amount"
         Size            =   "1958;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   49
         Top             =   1890
         Width           =   195
         VariousPropertyBits=   276824091
         Caption         =   "SL"
         Size            =   "344;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   2
         Left            =   855
         TabIndex        =   48
         Top             =   1890
         Width           =   1950
         VariousPropertyBits=   276824091
         Caption         =   "Client/Landlord Account"
         Size            =   "3440;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label14 
         Height          =   210
         Left            =   4455
         TabIndex        =   47
         Top             =   945
         Width           =   795
         VariousPropertyBits=   276824091
         Caption         =   "Reference"
         Size            =   "1402;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label11 
         Height          =   420
         Left            =   13050
         TabIndex        =   45
         Top             =   405
         Width           =   1440
         VariousPropertyBits=   276824091
         Caption         =   "Payable Amount"
         Size            =   "2540;741"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   10
         Left            =   15930
         TabIndex        =   34
         Top             =   1890
         Width           =   405
         VariousPropertyBits=   276824091
         Caption         =   "Total"
         Size            =   "714;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   6
         Left            =   13320
         TabIndex        =   33
         Top             =   1890
         Width           =   795
         VariousPropertyBits=   276824091
         Caption         =   "Reference"
         Size            =   "1402;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   5
         Left            =   7515
         TabIndex        =   32
         Top             =   1890
         Width           =   915
         VariousPropertyBits=   276824091
         Caption         =   "Description"
         Size            =   "1614;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   4
         Left            =   6030
         TabIndex        =   31
         Top             =   1890
         Width           =   900
         VariousPropertyBits=   276824091
         Caption         =   "Percentage"
         Size            =   "1587;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   420
         Index           =   3
         Left            =   3060
         TabIndex        =   30
         Top             =   1890
         Width           =   2175
         VariousPropertyBits=   276824091
         Caption         =   "Client/Landlord  Name"
         Size            =   "3836;741"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   210
         Left            =   225
         TabIndex        =   28
         Top             =   945
         Visible         =   0   'False
         Width           =   1050
         VariousPropertyBits=   276824091
         Caption         =   "Payable Type"
         Size            =   "1852;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label12 
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   26
         Top             =   360
         Width           =   990
         VariousPropertyBits=   276824091
         Caption         =   "Invoice Date"
         Size            =   "1746;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label13 
         Height          =   210
         Left            =   8190
         TabIndex        =   25
         Top             =   405
         Width           =   420
         VariousPropertyBits=   276824091
         Caption         =   "Fund"
         Size            =   "741;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   210
         Left            =   4500
         TabIndex        =   24
         Top             =   360
         Width           =   690
         VariousPropertyBits=   276824091
         Caption         =   "Property"
         Size            =   "1217;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   45
      TabIndex        =   18
      Top             =   7425
      Width           =   18150
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   510
         Left            =   15975
         TabIndex        =   6
         Top             =   360
         Width           =   1665
      End
      Begin VB.CommandButton cmdCreatePI 
         Caption         =   "Generate Rent Payable"
         Height          =   510
         Left            =   13545
         TabIndex        =   5
         Top             =   360
         Width           =   2250
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1590
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   18150
      Begin VB.TextBox txtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   6075
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   405
         Width           =   2970
      End
      Begin VB.TextBox txtClientAccount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   405
         Width           =   2250
      End
      Begin VB.TextBox txtAvailableFund1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   900
         Width           =   1260
      End
      Begin VB.TextBox txtStatementBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   405
         Width           =   1260
      End
      Begin VB.TextBox txtStatementDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   6075
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "01/01/2000"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtStatementNumber2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   2250
      End
      Begin MSForms.Label Label2 
         Height          =   210
         Left            =   4770
         TabIndex        =   20
         Top             =   405
         Width           =   975
         VariousPropertyBits=   276824091
         Caption         =   "Client Name"
         Size            =   "1720;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   210
         Left            =   225
         TabIndex        =   19
         Top             =   405
         Width           =   1185
         VariousPropertyBits=   276824091
         Caption         =   "Client Account"
         Size            =   "2090;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   210
         Left            =   9855
         TabIndex        =   17
         Top             =   900
         Width           =   1260
         VariousPropertyBits=   276824091
         Caption         =   "Available Funds"
         Size            =   "2222;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label8 
         Height          =   210
         Left            =   225
         TabIndex        =   16
         Top             =   900
         Width           =   1500
         VariousPropertyBits=   276824091
         Caption         =   "Statement number "
         Size            =   "2646;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label9 
         Height          =   210
         Left            =   4770
         TabIndex        =   15
         Top             =   900
         Width           =   1230
         VariousPropertyBits=   276824091
         Caption         =   "Statement Date"
         Size            =   "2170;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label10 
         Height          =   210
         Left            =   9855
         TabIndex        =   14
         Top             =   405
         Width           =   1485
         VariousPropertyBits=   276824091
         Caption         =   "Statement Balance"
         Size            =   "2619;370"
         FontName        =   "Myriad Web"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   45
      X2              =   18090
      Y1              =   3375
      Y2              =   3375
   End
End
Attribute VB_Name = "frmGenaratePayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTextBox As String
Public szCurrentStatementID As String
Public szClientID As String
Dim strPreviousDate As String
Dim statementTodate As String
Dim iRow As Integer
Dim strListofFunds As String
Dim strListOfPayableTypeID As String
Dim boolConsolidatedStatement As Boolean
Dim whichFieldToCheck As String
Public strRef As String

'Private Sub GeneratePreview(szStatmentID As String)
'    Dim adoconn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoconn.Open getConnectionString
'    Dim dblLasClosingBalance As Double
'    Dim szSQL As String
'    'Before writing this table you need to delete this table
'
'    adoconn.Execute "Delete from  RentSummaryStatementPreview"
'    szSQL = "Select StatementClosingBal from RentSummaryStatement where StatementNo=" & GetLastStatementNoByClient & " AND ClientIDLandlordID='" & txtClientAccount.text & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenKeyset, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        dblLasClosingBalance = rsRentSummaryStatement!StatementClosingBal
'    End If
'    Dim X As String
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    rsRentSummaryStatement.Open "Select * from RentSummaryStatementPreview where 1=2", adoconn, adOpenDynamic, adLockOptimistic
'    With rsRentSummaryStatement
'            .AddNew
'            !StatementID = szStatmentID 'we are setting this column atutomatically
'            !StatementNo = GetLastStatementNoByClient + 1
'            !ClientIDLandlordID = txtClientAccount.text
'            !BankCode = szSelectedBankAccount
'            !PreviousStatementDate = IIf(GetLastStatementDateByClient = "", Null, GetLastStatementDateByClient)
'            !StatementDate = Format(txtPayableDate2.text, "dd/mmmm/yyyy")
'            !StatementOpBal = dblLasClosingBalance
'            !Retentions = txtRetention.text 'we need to further analyse detail/add/deduct retension
'            !Clearretentions = False 'Will need to come again
'            !AccrualsAcBalance = GetAccrualsControlBalance 'you need to check
'            !SupplierAcBalance = GetSupplierOSAmount 'GetBalance("Supplier") 'GetBalanceSupplier'wrong'' for consolidated I need not filter it by property but for other I need to filter
'            !ManagingAgentAcBalance = GetAgentBalance 'GetBalance("Agent") 'GetBalanceAgent'wrong' for consolidated I need not filter it by property but for other I need to filter
'            !ClientACBalance = GetClientACBalance
'            !LandLordACBalance = GetLandLordACBalance
'            !ListOffundID = ListOfFundsForDBSave
'            !ListOfPayableTypeID = ListOfFundsForDBSave ' ListOfPayableTypes
'            !ListOfinputProperties = ListOfProperties
'            !TenantDepositsReceived = GetRentDeposit()
'            !Availablefunds = getAvailablefunds(dblLasClosingBalance)
'            !PaymentsonAccount = -GetPaymentsonAccount 'date  filter added
'            'New fields added 2021-02-22
'
'            !ClientPayments = GetClientPayments
'            !LandlordPayments = GetLandLordPayments
'            !ManagingAgentPayments = GetAGENTPaymentsPreview(CLng(szStatmentID))
'
'             'New fields added 2021-01-24
'            !TenantReceipts = GetTenantReceiptsPreview(CLng(szStatmentID))
'
'            !SupplierPayments = GetSupplierPaymentPreview(CLng(szStatmentID)) 'Purchase payment
'            !BankPaymentReceipts = GetBankPaymentReceipts
'            'addded newly by anol 2021-08-19
'            !BankPayment = GetBankPaymentPreview
'            !BankReceipts = GetBankreceiptsPreview
'            !BankACBalancePreview = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient)
'            !ClientLandlordBalance = GetClientACBalance + GetLandLordACBalance
'
'
'            !PayableAmount = 0 'Val(txtRentPayable.text)
'            '!StatementClosingBal = BankAccBalance(adoconn, szSelectedBankAccount, szSelectedClient) - !Availablefunds 'getClosingBalance(dblLasClosingBalance)
'            !StatementClosingBal = !Availablefunds  'getClosingBalance(dblLasClosingBalance)
'            !PINumber = ""
'            !Generated_Date = Format(Now, "dd/mmmm/yyyy")
'            !Printed = False
'            !Emailed = False
'            !Invoiced = False
'            !PostTohistory = False
'            .Update
'    End With
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    adoconn.Close
'    Set adoconn = Nothing
'End Sub
'Private Function GetLastStatementNoByClient() As Integer
'    Dim intmaxStatementNo As Integer
'    Dim adoconn As New ADODB.Connection
'    Dim rsRentSummaryStatement As New ADODB.Recordset
'    adoconn.Open getConnectionString
'    Dim szSQL As String
'    'Rent Summary Statement Opening Balance (=Closing balance of previous statement)
'    'This is by client
'    'Get ID by Client max ID from RentSummaryStatement
'    szSQL = "Select max(StatementNo) as IDbyCL from RentSummaryStatement where ClientIDLandlordID='" & szSelectedClient & "'"
'    rsRentSummaryStatement.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'    If Not rsRentSummaryStatement.EOF Then
'        GetLastStatementNoByClient = IIf(IsNull(rsRentSummaryStatement!IDbyCL), 0, rsRentSummaryStatement!IDbyCL)
'    End If
'    rsRentSummaryStatement.Close
'    Set rsRentSummaryStatement = Nothing
'    adoconn.Close
'    Set adoconn = Nothing
'End Function
Private Function Feestogenerate() As Boolean
        Dim lSlNumber As Long
        Dim adoconn As New ADODB.Connection
        Dim adoPIHeader As New ADODB.Recordset
        Dim adoPISplit As New ADODB.Recordset
        Dim szSQL As String
        Dim szSQLManagingAgent  As String
        Dim szMYID As String
        Dim szFundID As String
        Dim szSelectedPayableTypeID As String
        Dim szSQL1 As String
        Dim rsfixedMethod As New ADODB.Recordset
        Dim rsfixedMethodDetails As New ADODB.Recordset
        Dim j As Integer
        Dim percnetageOramount As Double
        Dim dblGrandTotal As Double
        Dim dtNextDue As Date
        Dim dtFDD As Date
        Dim dblFeqID As Integer
        Dim dblTotalAmount As Double
        Dim lngMgtFeeSL As Long
        Dim iCountPI As Integer
        Dim iClientCount As Integer
        Dim iPropertyCount As Integer
        Dim strLastChargeDate  As String
        Dim strFundName As String
        Dim strStopDate  As String
        Dim dblCapAmount As Double
        Dim rsManagingAgent As New ADODB.Recordset
        Dim rsGlobalData As New ADODB.Recordset
        Dim szManagingAgent() As String
        Dim iManagingAgentCount As Integer
        Dim bControlACForPayable As Boolean
        Dim FinalControlACForPayable As String
        Dim szTemp
        Dim dblNoOfDaysToSendMFB4Due As Integer
        Dim dtNDDInitial As Date
        Dim strFromDate As String
        Dim strToDate As String
        Dim rsFromandToDate As New ADODB.Recordset
        Dim szSQLFrom As String
        Dim dtNDD As Date

        Dim iCount As Long
        Dim iCount1 As Long
        Dim rCount As Long
        Dim dblFundId As Integer
        Dim dblDemandTypeId As Integer
        Dim i As Integer
        Dim lT_ID As Long

        Dim strSelectedDemandType As String
        Dim szSQL5 As String
        Dim rstSet As New ADODB.Recordset
        Dim szPropertySelectionALL As String
        Dim szPropertySelection1 As String
        Dim szSelectedClient As String


 '      ************************************Write tblPurInv **************************************

        adoconn.Open getConnectionString
        adoconn.Execute "Delete from tblPurInvPreview"
        adoconn.Execute "Delete from tblPurInvSRecPreview"
        adoconn.Close
'      For iClientCount = 1 To flxClients.Rows - 1
'            If flxClients.TextMatrix(iClientCount, 0) = "X" Then
'                        szSelectedClient = flxClients.TextMatrix(iClientCount, 1)
'                        If szSelectedClient = "" Then Exit Function
'
'            For iPropertyCount = 1 To flxProperties.Rows - 1
'                    If iPropertyCount >= flxProperties.Rows Then Exit For
'                    If flxProperties.TextMatrix(iPropertyCount, 0) = "X" And szSelectedClient = flxProperties.TextMatrix(iPropertyCount, 3) Then
'                            szPropertySelection1 = flxProperties.TextMatrix(iPropertyCount, 1)
            adoconn.Open getConnectionString
            Dim rsCharge As New ADODB.Recordset
            szTemp = ""
            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And (F.FundID)=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "'  And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "'"
              rsManagingAgent.Open szSQLManagingAgent, adoconn, adOpenDynamic, adLockOptimistic
              szTemp = SQL2String(rsManagingAgent, 0)
              rsManagingAgent.Close
              If Len(szTemp) > 0 Then
                    szManagingAgent = Split(szTemp, ",")
              Else
                    adoconn.Close
                    GoTo EndOfAgreement
              End If


             szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
             rstSet.Open szSQL5, adoconn, adOpenStatic, adLockReadOnly
             lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
             rstSet.Close
             Set rstSet = Nothing
             adoconn.Close

            For iManagingAgentCount = 0 To UBound(szManagingAgent) 'this for shall end after the creation of the PI preview

            adoconn.Open getConnectionString
            'For each managing agent I am creating PI
            lSlNumber = SlNumber("PI", "tblPurInv", adoconn)

            szSQL = "SELECT agr.EachPeriod,agr.Capamount, agr.StopDate, CPA.agreementStartDate,CPA.agreementEndDate,agr.CHARGE_METHOD," & _
                "cpa.agreementEndDate as agreementEndD,agr.LastChargeDate,agr.TotalAmount,agr.Amount,agr.Fund,F.FundName, " & _
                "agr.NtDueDate,agr.FDD,agr.Frequency as FrequencyID,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as FrequencyName,ManagingAgentID,agr.DEMAND_TYPE  " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And (F.FundID)=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "'  And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "' AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
        Debug.Print szPropertySelection1
        Debug.Print Trim(szManagingAgent(iManagingAgentCount))
        Debug.Print iPropertyCount
            rsCharge.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
            If rsCharge.EOF Then
                rsCharge.Close
                Set rsCharge = Nothing

                GoTo EndOfAgreement
            End If


            i = 1
                'Dim rsGlobalData As New ADODB.Recordset
                Dim VAT_ID As String
                Dim VAT_CODE As String
                Dim VAT_RATE As Double

                 rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' AND vatOptionEnabled=true", adoconn, adOpenStatic, adLockReadOnly
                 If Not rsGlobalData.EOF Then
                          VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                          VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                          VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                 Else
                          VAT_ID = -1
                          VAT_RATE = 0
                          VAT_CODE = ""
                 End If
                 rsGlobalData.Close
                 Set rsGlobalData = Nothing

                 Dim rsGlobalData1 As New ADODB.Recordset
                 Dim bolVatOptionEnabled As Boolean
                 Dim bolOptedTotax As String
                 Dim strManagingAgentID As String
                 rsGlobalData1.Open "Select vatOptionEnabled from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' ", adoconn, adOpenStatic, adLockReadOnly

                If Not rsGlobalData1.EOF Then
                        bolVatOptionEnabled = rsGlobalData1("vatOptionEnabled").Value
                        strManagingAgentID = rsCharge("ManagingAgentID").Value
                End If
                rsGlobalData1.Close





            szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
            adoPIHeader.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
            adoPIHeader.Close
             j = 1
                 rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoconn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData1.EOF Then
                        bolOptedTotax = rsGlobalData1("optedTotax").Value
                Else
                        bolOptedTotax = False
                End If
                rsGlobalData1.Close



            While Not rsCharge.EOF
                  dblTotalAmount = rsCharge.Fields.Item("TotalAmount").Value
                  dblCapAmount = rsCharge.Fields.Item("CapAmount").Value
                  dblFundId = rsCharge.Fields.Item("Fund").Value
                  dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value

                  strFundName = rsCharge.Fields.Item("FundName").Value
                  If Not IsNull(rsCharge.Fields.Item("NtDueDate").Value) Then
                      dtNextDue = rsCharge.Fields.Item("NtDueDate").Value
                      dtNDDInitial = rsCharge.Fields.Item("NtDueDate").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FDD").Value) Then
                      dtFDD = rsCharge.Fields.Item("FDD").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FrequencyID").Value) Then
                    If rsCharge.Fields.Item("FrequencyID").Value <> "" Then
                            dblFeqID = rsCharge.Fields.Item("FrequencyID").Value
                      End If
                  End If
                szMYID = UniqueID()
                strLastChargeDate = IIf(IsNull(rsCharge("LastChargeDate").Value), "", rsCharge("LastChargeDate").Value)
                If strLastChargeDate = "" Then
                       rsCharge.Close

                       GoTo EndOfAgreement
                End If
                rsGlobalData.Open "Select NoOfDaysToSendMFB4Due from globaldata where PropertyID='" & szPropertySelection1 & "'", adoconn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData.EOF Then
                    dblNoOfDaysToSendMFB4Due = IIf(IsNull(rsGlobalData!NoOfDaysToSendMFB4Due), 0, rsGlobalData!NoOfDaysToSendMFB4Due)
                End If
                rsGlobalData.Close

                If DateDiff("d", Date, rsCharge("NtDueDate").Value) > dblNoOfDaysToSendMFB4Due Then
                        GoTo EndOfChargeType
                End If


                                'validations
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then   'when working on the fixed procedure only 1 line of setup is done then (fixed basis)
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then

                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtPayableDate2.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing

                                                        GoTo EndOfChargeType
                                                End If

                            '      ************************************Write tblPurInvSRec **************************************

                                                    dblTotalAmount = IIf(IsNull(rsCharge("EachPeriod").Value), 0, rsCharge("EachPeriod").Value) 'rsfixedMethodDetails.Fields.Item("Amt").Value
        '                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                    dblFundId = rsCharge.Fields.Item("Fund").Value

                                                    If dblCapAmount > 0 Then
                                                           'make a condition in the split as well so that amount doesnot exeed cap amount
                                                           If dblTotalAmount > dblCapAmount Then
                                                                dblTotalAmount = dblCapAmount
                                                           End If
                                                    End If

'                                                                 txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
'                                                                dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
'                                                                dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                strFromDate = Format(dtNextDue, "dd/MM/yyyy")
'                                                                strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy")

                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                   ' adoPISplit.Close
                                                    adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                    'Add New Records. At least there is only one split line
                                                       With adoPISplit
                                                           .AddNew
                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                           .Fields.Item("ParentID").Value = szMYID
                                                           .Fields.Item("TRAN_ID").Value = j

                                                          'If chkAssignProperty.Value = 0 Then
                                                                .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                           ' Else
                                                             '    .Fields.Item("TRANS").Value = ""
                                                            'End If
                                                           .Fields.Item("UNIT_ID").Value = ""
                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                           .Fields.Item("RecoverablePt").Value = 0

                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then

                                                                     .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value

                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled means global data and bolOptedTotax is supplier table
                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                   strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                If Not rsGlobalData.EOF Then
                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                Else
                                                                         VAT_ID = -1
                                                                         VAT_RATE = 0
                                                                         VAT_CODE = ""
                                                                End If
                                                                rsGlobalData.Close

                                                                 'done modification on 15-10-2021
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value


                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then

                                                                 .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value

                                                           End If

                                                           .Update
                                                       End With
                                                    adoPISplit.Close
                                                   If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then

                                                     End If
                                                    dblGrandTotal = dblGrandTotal + dblTotalAmount


                                End If 'end of rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX"
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then ' receipt basis

                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then

                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtPayableDate2.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing

                                                        GoTo EndOfChargeType

                                                End If


                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
                                            "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
                                            "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtPayableDate2.text, "dd MMM yyyy") & "# " & _
                                            "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                            szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & ""

                                                                                           'by anol 20211024
                                                    rsfixedMethod.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for reciept type . I have not written for the credit yet need to understand the principle

                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)

                            '      ************************************Write tblPurInvSRec **************************************

                                     szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
                                     "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtPayableDate2.text, "dd MMM yyyy") & "# " & _
                                     "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                     szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & ""

                                                        'need to consider the selected property in where clause
                                                        'modified on 20211103
                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtPayableDate2.text, "dd MMM yyyy") & "# " & _
                                                        "AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "# and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                                        szPropertySelection1 & "' and R.ISMGTFEE=false AND Rs.FundID=" & dblFundId & " "


                                                        rsFromandToDate.Open szSQLFrom, adoconn, adOpenStatic, adLockReadOnly
                                                        If Not rsFromandToDate.EOF Then
                                                                strFromDate = Format(rsFromandToDate("DateFromMin").Value, "dd/MM/yyyy")
                                                                strToDate = Format(rsFromandToDate("DateTOMAX").Value, "dd/MM/yyyy")
                                                        Else
                                                                strFromDate = ""
                                                                strToDate = ""
                                                        End If
                                                        rsFromandToDate.Close

                                                         'Need to take only allocated transactions
                                                        If rsfixedMethodDetails.State = 1 Then
                                                            rsfixedMethodDetails.Close
                                                        End If
                                                                 rsfixedMethodDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Amt").Value), 0, rsfixedMethodDetails.Fields.Item("Amt").Value)
                                                                                 If dblTotalAmount <= 0 Then
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 'dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * (percnetageOramount / 100)
                                                                                        dblTotalAmount = Round(dblTotalAmount, 2)
                                                                                 End If

                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'Else
                                                                                          '    .Fields.Item("TRANS").Value = ""
                                                                                         'End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"

                                                                                        If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                                                dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then 'bolVatOptionEnabled=global data
                                                                                                'Modified by anol 2021-10-15
                                                                                                 dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                                 .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled=global data
                                                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                                                   strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                                                If Not rsGlobalData.EOF Then
                                                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                                                Else
                                                                                                         VAT_ID = -1
                                                                                                         VAT_RATE = 0
                                                                                                         VAT_CODE = ""
                                                                                                End If
                                                                                                rsGlobalData.Close
                                                                                                'done modification on 15-10-2021
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                           End If
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close
                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then

                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then

                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtPayableDate2.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing

                                                        GoTo EndOfChargeType

                                                End If
                                                 szSQL1 = "Select sum(SWITCH(TYPE =1,RS.Amount,TYPE =2,-RS.Amount,TYPE =24,RS.Amount)) as Amt from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                              "R.TransactionID=RS.RptHeader AND RDate<=#" & Format(txtPayableDate2.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  and Type in (1,2,24) " & _
                                                              "AND U.UnitNumber=R.UnitID AND U.PropertyID='" & szPropertySelection1 & "' and ISMGTFEE=false AND Rs.FundID=" & dblFundId & ""
                                                              'need to consider the selected property in where clause
                                                            '  rsfixedMethod.Close
                                                    rsfixedMethod.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for recipt type . I have not written for the credit yet need to understand the principle

                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)

                            '      ************************************Write tblPurInvSRec **************************************

                                                                 szSQL = "Select sum(SWITCH(TYPE =3,RS.Amount,TYPE =4,-RS.Amount,TYPE =23,RS.Amount)) as Amt,rs.FundID from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                                         "R.TransactionID=RS.RptHeader and Type in (3,4,23) AND RDate<=#" & Format(txtPayableDate2.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  AND  " & _
                                                                         "U.UnitNumber=R.UnitID and ISMGTFEE=false AND U.PropertyID='" & szPropertySelection1 & "' AND RS.FundID=" & dblFundId & " group by RS.FundID"
                                                                 'need to consider the selected property in where clause
                                                                 If rsfixedMethodDetails.State = 1 Then
                                                                     rsfixedMethodDetails.Close
                                                                 End If
                                                                 rsfixedMethodDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = rsfixedMethodDetails.Fields.Item("Amt").Value
                                                                                 If dblTotalAmount < 0 Then
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * percnetageOramount / 100
                                                                                 End If

                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         'Else
                                                                                          '    .Fields.Item("TRANS").Value = ""
                                                                                         'End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"

                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                        .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close

                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then

                                szSQL = "SELECT * FROM tblPurInvPreview"

                                    dblTotalAmount = Round(dblTotalAmount, 2)
                                    If dblTotalAmount = 0 Then

                                           GoTo EndOfChargeType
                                    End If
                                    With adoPIHeader
                                            .Open szSQL, adoconn, adOpenDynamic, adLockPessimistic
                                            .AddNew
                                            .Fields.Item("MY_ID").Value = szMYID
                                            .Fields.Item("SlNumber").Value = lSlNumber
                                            .Fields.Item("SUPP_AC").Value = Trim(szManagingAgent(iManagingAgentCount))
                                            .Fields.Item("TRAN_DATE").Value = Format(txtPayableDate2.text, "DD MMMM YYYY")
                                            .Fields.Item("TransactionType").Value = 6
                                            .Fields.Item("INV_NO").Value = szPropertySelection1 + "-" + "MFee" + "-" + CStr(lngMgtFeeSL)
                                             lngMgtFeeSL = lngMgtFeeSL + 1
                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                            .Fields.Item("History").Value = False
                                            .Fields.Item("TrfPayment").Value = False
                                            .Fields.Item("PropertyID").Value = ""
                                            .Fields.Item("CL_ID").Value = szSelectedClient
                                            .Fields.Item("NLPost").Value = False
                                            .Fields.Item("DueDate").Value = Format(dtNDDInitial, "DD MMMM YYYY")
                                            .Fields.Item("PostingDate").Value = Format(txtPayableDate2.text, "DD MMMM YYYY")
                                            .Fields.Item("ReportFromDatePreview").Value = IIf(strFromDate = "", Null, strFromDate)
                                            .Fields.Item("ReportToDatePreview").Value = IIf(strToDate = "", Null, strToDate)
                                            .Update
                                            iCountPI = iCountPI + 1
                                            lSlNumber = lSlNumber + 1
                                   End With
                                   adoPIHeader.Close
                                   If iCountPI > 0 Then
                                        Feestogenerate = True
                                        adoconn.Close
                                        Set adoconn = Nothing
                                        Exit Function
                                   End If


EndOfChargeType:
            rsCharge.MoveNext
            j = j + 1

      Wend
                        rsCharge.Close
                        Set rsCharge = Nothing
                        adoconn.Close
                        Set adoconn = Nothing
EndOfOneManagingAgentforOneAgreement:
           Next iManagingAgentCount
         'end if for 'X' in grid seletion
EndOfAgreement:
          'End If
'
'          Next iPropertyCount
'          End If 'end if only for selected client
'  Next iClientCount
            
       
End Function
'Private Function FeeshasBeenPaid() As Boolean
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    Dim rstblPurInv As New ADODB.Recordset
'    rstblPurInv.Open "Select * from tblPurInv V,tlbPayment P  where V.MY_ID=P.PI AND isManagementFee=true and OSAmount>0 and TRAN_DATE<= #" & _
'    txtStatementDate1.text & " # AND TRAN_DATE>#" & txtLastStatementDate1.text & "#", adoConn, adOpenStatic, adLockReadOnly
'    If Not rstblPurInv.EOF Then
'            FeeshasBeenPaid = True
'            rstblPurInv.Close
'            adoConn.Close
'            Exit Function
'    End If
'    rstblPurInv.Close
'    adoConn.Close
'    Set adoConn = Nothing
'End Function

Private Function GetSupplierOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    'If boolConsolidatedStatement = 1 Then
            'whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
'    If boolConsolidatedStatement = 1 Then
            'whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'    Else
'            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Supplier')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetSupplierOSAmount = GetSupplierOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetClientOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    'If boolConsolidatedStatement = 1 Then
            'whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Client')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
'    If boolConsolidatedStatement = 1 Then
            'whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'    Else
'            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Client')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientOSAmount = GetClientOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetAgentOSAmount() As Double   'This function return result as minus'This is getting supplier balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    Dim whereProperty As String
    ' for consolidated I need not filter it by property but for other I need to filter
    'If boolConsolidatedStatement = 1 Then
            'whereProperty = "(P.UNITID IN (" & ListOfProperties & ") OR isnull(P.UNITID)  OR P.UNITID='' ) AND "
'    Else
'            whereProperty = "P.UNITID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select  SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,24) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Agent')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentOSAmount = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    'I have modified the code according to added field in paymentsplit and splite into 2 SQL modified 20211024
'    If boolConsolidatedStatement = 1 Then
            'whereProperty = "(S.PropertyID IN (" & ListOfProperties & ") OR isnull(S.PropertyID)  OR S.PropertyID='' ) AND "
'    Else
'            whereProperty = "S.PropertyID in (" & ListOfProperties & ") AND "
'    End If
    szSQL = "Select   SUM(S.Amount) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND   " & whereProperty & " P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(7,8,9) AND S.FundID=F.FundID  AND P.ClientID ='" & szClientID & "'  and SP.TYPE in ('Agent')   AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetAgentOSAmount = GetAgentOSAmount - IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Sub cmdCreatePI_Click()
        'Fund saving technic
    'you you appostrophe to save fund list while you save it to database fields you may have problem then you can put small a as a seperator
      Dim bControlACForPayable As Boolean
      Dim FinalControlACForPayable As String
      Dim iCount As Long
      Dim iCount1 As Long
      Dim iSplitID As Long
      Dim lSlNumber As Long
      Dim adoconn As New ADODB.Connection
      Dim adoPIHeader As New ADODB.Recordset
      Dim adoPISplit As New ADODB.Recordset
      Dim szSQL As String
      Dim szMYID As String
      Dim szSelectedPayableTypeID As String
      Dim colPInumber As String
      Dim rsStatementSplit As New ADODB.Recordset
      If txtPayableDate2.text = "" Then
             MsgBox "Please enter invoice date", vbInformation, "Warning "
             FocusControl txtPayableDate2
             Exit Sub
      End If
      If Val(txtAvailableFund1.text) < Val(txtTotalAmount.text) Then
                 MsgBox "The total amount of rent payable entered cannot be greater than Available Funds.", vbYesNo, "Warning"
                 Exit Sub
      End If
      If Val(txtAvailableFund1.text) <> Val(txtTotalAmount.text) Then
            If MsgBox("The total amount of rent payable entered is less than the Available Funds. Do you wish to generate the amount of rent payable entered?", vbYesNo, "Warning") = vbNo Then
                 Exit Sub
            End If
      End If
      If szCurrentStatementID = "" Then
            MsgBox "Please Select a statement", vbInformation, "Warning"
            Frame4.Visible = False
            Exit Sub
      End If
      If txtFundForPI.text = "" Then
            MsgBox "Please Select a fund.", vbInformation, "Warning"
            Exit Sub
      End If
      
      If GetSupplierOSAmount > 0 Then
        If MsgBox("You have outstanding supplier balances to pay. Do you wish to pay them before Generate Rent Payable?", vbYesNo, "Supplier Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szClientID
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
    End If
    If GetClientOSAmount > 0 Then
        If MsgBox("You have outstanding Client balances to pay. Do you wish to pay them before generating rent payable?", vbYesNo, "Client Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szClientID
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
    End If
    If GetAgentOSAmount > 0 Then
        If MsgBox("You have outstanding Managing Agent balances to pay. Do you wish to pay them before generating rent payable?", vbYesNo, "Managing Agent Os Balance") = vbYes Then
                LoadForm frmPurchaseExpense
                frmPurchaseExpense.tabPurExp.Tab = 1
                frmPurchaseExpense.txtClientIDPurPay.text = szClientID
                frmPurchaseExpense.txtBankCode.text = ""
                frmPurchaseExpense.txtBankAc.text = ""
                Exit Sub
        Else
            'proceed
        End If
    End If
'    If Feestogenerate Then
'            If MsgBox("You have management fees to generate. Do you wish to generate  " & _
'                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
'                 Exit Sub
'            Else
'                'proceed
'            End If
'    End If
    'check if management fee has been paid or not
'    If FeeshasBeenPaid Then
'        If MsgBox("You have management fees to pay. Do you wish to generate  " & _
'                "them before producing your client statement?", vbYesNo, "Please confirm") = vbYes Then
'                 Exit Sub
'            Else
'                'proceed
'            End If
'    End If
      
      For iCount = 1 To flxPayFees.Rows - 1
            If flxPayFees.TextMatrix(iCount, 2) <> "" Then
                   iCount1 = iCount1 + 1
            End If
      Next
      If iCount1 < 1 Then
            MsgBox "Please apply the amount of rent you wish to pay.", vbInformation, "Warning"
            FocusControl cmdApply
            Exit Sub
      End If
      adoconn.Open getConnectionString
      If FinalControlACForPayable = "" Then
            FinalControlACForPayable = GetNominalCodeForControlAccount(adoconn, "Rent & Other Amounts Payable (P&L)", txtClientAccount.text)
      End If
      'if still control acount is empty generate a warning message and exit this sub procedure
      If FinalControlACForPayable = "" Then
          MsgBox "A Control Account is not set for this Payable type", vbInformation, "Warning"
          Exit Sub
      End If
      
      
 '      ************************************Write tblPurInv **************************************
      'szMYID = UniqueID()
      adoconn.Close
      Set adoconn = Nothing
      
      
      'Exit Sub
      'start a loop for each line in grid
      If boolConsolidatedStatement = True Then
              iSplitID = 1
                        For iCount = 1 To flxPayFees.Rows - 1
                              If flxPayFees.TextMatrix(iCount, 2) <> "" Then
                                        szMYID = UniqueID()
                                        adoconn.Open getConnectionString
                                        szSQL = "SELECT * FROM tblPurInv"
                                        adoconn.BeginTrans
                                        lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
                                         
                                        With adoPIHeader
                                                  .Open szSQL, adoconn, adOpenDynamic, adLockPessimistic
                                                  .AddNew
                                                  .Fields.Item("MY_ID").Value = szMYID
                                                  .Fields.Item("CreatedBy").Value = User
                                                  .Fields.Item("CreatedDate").Value = Now
                                                  .Fields.Item("SlNumber").Value = lSlNumber
                                                  .Fields.Item("SUPP_AC").Value = flxPayFees.TextMatrix(iCount, 3) 'txtClientAccount.text
                                                  .Fields.Item("TRAN_DATE").Value = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                  .Fields.Item("TransactionType").Value = 6
                                                  .Fields.Item("INV_NO").Value = txtReference.text ' szCurrentStatementID 'txtInv(0).text
                                                  .Fields.Item("TOTAL_AMOUNT").Value = flxPayFees.TextMatrix(iCount, 15) ' CCur(txtRentPayable.text)
                                  '                .Fields.Item("TTP").Value = "PURCHASE INVOICE" 'CByte(TransactionTakePlace("TTP", "PURCHASE INVOICE", adoconn))
                                                  .Fields.Item("History").Value = False
                                                  .Fields.Item("TrfPayment").Value = False
                                                  .Fields.Item("PropertyID").Value = ""
                                                  .Fields.Item("CL_ID").Value = txtClientAccount.text
                                                  .Fields.Item("NLPost").Value = False
                                                  .Fields.Item("DueDate").Value = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                  .Fields.Item("PostingDate").Value = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                  .Fields.Item("isRentPayable").Value = True
                                                  .Update
                                                  'Exit Sub
                                        End With
                                       
                                  '      adoconn.Close
                                  '      Set adoconn = Nothing
                                    
                                  '      ************************************Write tblPurInvSRec **************************************
                                  'adoconn.Open getConnectionString
                                        szSQL = "SELECT * FROM tblPurInvSRec"
                                     adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                  
                                      'Add New Records. At least there is only one split line
                                           With adoPISplit
                                              .AddNew
                                              .Fields.Item("MY_ID").Value = UniqueID()
                                              .Fields.Item("ParentID").Value = szMYID
                                              .Fields.Item("TRAN_ID").Value = "1"
                                              .Fields.Item("TRANS").Value = flxPayFees.TextMatrix(iCount, 9)  ' If you select One property then you can write a value here
                                              .Fields.Item("UNIT_ID").Value = ""
                                              .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                              .Fields.Item("DEPT_ID").Value = flxPayFees.TextMatrix(iCount, 10)
                                             ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                              .Fields.Item("RecoverablePt").Value = 0
                                              .Fields.Item("description").Value = flxPayFees.TextMatrix(iCount, 7)
                                              .Fields.Item("NET_AMOUNT").Value = flxPayFees.TextMatrix(iCount, 15)
                                              .Fields.Item("TAX_CODE").Value = flxPayFees.TextMatrix(iCount, 13)
                                              .Fields.Item("VAT").Value = Val(flxPayFees.TextMatrix(iCount, 14))
                                              .Fields.Item("TOTAL_AMOUNT").Value = flxPayFees.TextMatrix(iCount, 15)
                                              .Update
                                           End With
                                        adoPIHeader.Close
                                        'Exit Sub
                                         '*************************************Write tlbPayment **************************************
                                        Dim lT_ID As Long
                                        szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
                                        adoPIHeader.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                        lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
                                        adoPIHeader.Close
                                        
                                        
                                         
                                         szSQL = "SELECT * FROM tlbPayment where 1=2"
                                       
                                         With adoPIHeader
                                                .Open szSQL, adoconn, adOpenDynamic, adLockOptimistic 'Add New Mode
                                                .AddNew
                                                !TransactionID = lT_ID
                                                !Type = 6  'PP - Purchase Invoice, look in the tlbTransactionType
                                                !SageAccountNumber = flxPayFees.TextMatrix(iCount, 3) ' txtClientAccount.text
                                                !Pi = szMYID
                                                !PDate = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                !dDate = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                !ref = strRef 'txtReference.text '"PI from statement:" & szCurrentStatementID
                                                !Details = flxPayFees.TextMatrix(iCount, 7) '"PI from statement:SS" & szCurrentStatementID
                                                 
                                                !ExtRef = flxPayFees.TextMatrix(iCount, 7)
                                                !amount = flxPayFees.TextMatrix(iCount, 15)
                                                !OSAmount = !amount
                                                !PaymentView = True
                                               
                                                '!unitid = txtProperty.text
                                                !SlNumber = lSlNumber
                                                !fundID = flxPayFees.TextMatrix(iCount, 10)
                                  '              !AdjTag = IIf(bAdjustment, "Y", "N")
                                                !Recoverable = 0
                                                !postingDate = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                                !ClientID = txtClientAccount.text
                                                .Update
                                                .Close
                                       End With
                                       
                                       'Saving into statement Split
                                       Dim rsRentSummaryStatementMAXID As New ADODB.Recordset
                                       Dim maxID As Long
                                       rsRentSummaryStatementMAXID.Open "Select (Max(ID)) AS DID from RentSummaryStatementDetails", adoconn, adOpenStatic, adLockReadOnly
                                        If Not rsRentSummaryStatementMAXID.EOF Then
                                                    maxID = IIf(IsNull(rsRentSummaryStatementMAXID("DID").Value), 0, rsRentSummaryStatementMAXID("DID").Value) + 1
                                        End If
                                        rsRentSummaryStatementMAXID.Close
                        
                                        szSQL = "SELECT * FROM RentSummaryStatementdetails"
                                        rsStatementSplit.Open szSQL, adoconn, adOpenDynamic, adLockPessimistic
                                        With rsStatementSplit
                                            rsStatementSplit.AddNew
                                            !statementID = szCurrentStatementID
                                            !PINumber = lT_ID
                                            !amount = flxPayFees.TextMatrix(iCount, 15)
                                            !OSAmount = !amount
                                            !SageAccountNumber = flxPayFees.TextMatrix(iCount, 3)
                                            !ClientID = txtClientAccount.text
                                            !Id = maxID
                                            !splitID = iSplitID
                                            !PercentageLL = IIf(flxPayFees.TextMatrix(iCount, 6) = "N/A", 100, flxPayFees.TextMatrix(iCount, 6))
                                            rsStatementSplit.Update
                                        End With
                                        iSplitID = iSplitID + 1
                                        rsStatementSplit.Close
                                        
                                     '*************************************Write tlbPaymentSplit **************************************
                                     adoPISplit.Close
                                     Set adoPISplit = Nothing
                                        szSQL = "SELECT * FROM tlbPaymentSplit;"
                                     adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockPessimistic
                                  'Add New Records. At least there is one split line.
                                           With adoPISplit
                                              .AddNew
                                              .Fields.Item("TransactionID").Value = UniqueID()
                                              .Fields.Item("PayHeader").Value = lT_ID
                                              .Fields.Item("FundID").Value = flxPayFees.TextMatrix(iCount, 10)
                                              .Fields.Item("Amount").Value = flxPayFees.TextMatrix(iCount, 15)
                                              .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
                                              .Fields.Item("SplitID").Value = "1"
                                              .Fields.Item("DueDate").Value = Format(txtPayableDate2.text, "DD/MMMM/YYYY")
                                              .Fields.Item("Description").Value = flxPayFees.TextMatrix(iCount, 7)
                                              '.Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                              .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                              .Fields.Item("TRANS").Value = "" 'Put property ID  Later
                                              .Fields.Item("UNIT_ID").Value = "" '
                                              .Fields.Item("ScheduleID").Value = Null
                                  '            .Fields.Item("ScheduleID").Value = IIf(flxPI.TextMatrix(iRow, 20) = "", Null, _
                                  '                                                   flxPI.TextMatrix(iRow, 20))
                                  '            .Fields.Item("RecoverablePt").Value = flxPI.TextMatrix(iRow, 22)
                                  '            .Fields.Item("AllocTranID").Value = flxPI.TextMatrix(iRow, 23)
                                  
                                              .Update
                                           End With
                                        
                                     
                                     'Update statement with latest PI number
                                     
                                     colPInumber = colPInumber + " " + "PI" & lSlNumber & ""
                                     'Now Post the transaction to the NLPosting
                                     
                                     Dim szTran2Fix As String
                                     Dim postResult As Boolean
                                     If PI_Check(adoconn, szTran2Fix) = False Then
                                           adoconn.RollbackTrans
                                           adoconn.Close
                                           MsgBox "An error occurred while saving, transaction rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked"
                                           Exit Sub
                                      Else
                                           postResult = Export_PInPC_2_NL_ForClientLandLord(adoconn)
                                           If postResult = False Then 'this part modified by anol 2021-01-13
                                                      adoconn.RollbackTrans
                                                      adoconn.Close
                                                      MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
                                           Else
                                                      adoconn.CommitTrans
                                                      szSQL = "UPDATE tblPurInv AS P, tblPurInvSRec AS S " & _
                                                          "SET P.NLPost = TRUE " & _
                                                          "WHERE  P.MY_ID = S.ParentID AND NOT P.NLPost AND " & _
                                                              "(P.TransactionType = 6 OR P.TransactionType = 7);"
                                                      adoconn.Execute szSQL
                                           End If
                                           adoconn.Close
                                     End If
                              End If 'check for empty lind end if
                     Next ' For 2 lines in grid you save 2 PI
          Else 'else for consolidated false
          End If
     
     Dim SQLforInsert As String
     Dim rsRentSummaryStatement As New ADODB.Recordset
     adoconn.Open getConnectionString
     adoconn.Execute "Update RentSummaryStatement set PITransactionID='" & lT_ID & "',PINumber='" & colPInumber & "',Invoiced=true " & _
                                            "where StatementID=" & szCurrentStatementID & ""
     adoconn.Execute "Update RentSummaryStatement set PayableAmount= " & Val(txtTotalAmount.text) _
                                            & ", StatementClosingBal=round(StatementClosingBal,2)-" & Val(txtTotalAmount.text) & ",AvailableFunds=AvailableFunds-" & Val(txtTotalAmount.text) & "," & _
                                            " ClientACBalance=" & GetClientACBalance & ",LandlordACBalance=" & GetLandLordACBalance(szClientID) & " where StatementID=" & szCurrentStatementID & ""
                                            ' StatementClosingBal=StatementClosingBal-" & Val(txtTotalAmount.text) & ",
                                            ',StatementClosingBal=StatementClosingBal-" & Val(txtTotalAmount.text) & ",AvailableFunds=AvailableFunds-" & Val(txtTotalAmount.text)
    
    Call WorkOnMgtfeedueSupplierOS(adoconn, szCurrentStatementID)
''    adoconn.Execute "DELETE FROM ClientStatementPurchasesSnapshot where StatementID=" & szCurrentStatementID & ""
''    rsRentSummaryStatement.Open "Select InclSupplierOS ,InclMngtFeesDue,ListOfFundId,ClientIDLandlordID  From RentSummaryStatement where StatementID=" & szCurrentStatementID & "", adoconn, adOpenStatic, adLockReadOnly
''    Dim szWhere As String
''    If Not rsRentSummaryStatement.EOF Then
''            If IsNull(rsRentSummaryStatement("InclMngtFeesDue").Value) Then   'null means false
''                   ' if  Or rsRentSummaryStatement("InclSupplierOS").Value = False
''                'Enter data into snapshot table where there is I am not consedering osamount<amount because I want save all mgt fee (means fully paid or partially paid) and type management fee
''
''            ElseIf rsRentSummaryStatement("InclMngtFeesDue").Value = True Then
''                    szWhere = " AND C.osamount>0" 'if there is any allocation against PI this PI shall come in the snapshot table
''            End If
''            'Inserting mananagement fee into snapshot table (conditional) this is the only place for report I am inserting management fee?? ans: No
''            ' take only PI where there is no allocation
''           SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''                                 "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,C.OSAmount,'1' " & _
''                                 "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''                                  "S.parentID and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ") and C.ClientID='" & rsRentSummaryStatement("ClientIDLandlordID").Value & _
''                                  "' and P.TransactionType=6 AND P.isManagementFee=true " & szWhere
''
''            'this one inserting allocated and unallocated both management fee
''            adoconn.Execute "Insert into ClientStatementPurchasesSnapshot (StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
''                            "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''                            SQLforInsert
''
''
''            If IsNull(rsRentSummaryStatement("InclSupplierOS").Value) Then   'null means false
''
''                'Enter data into snapshot table where there is osamount<amount (means fully paid or partially paid) and type management fee
''            ElseIf rsRentSummaryStatement("InclSupplierOS").Value = True Then
''                   szWhere = " AND amount>osamount and osamount<>0" 'Fully unpaid
''            End If
''            '********* 'Inserting mananagement fee into snapshot table (conditional)*****************
''                    'Type:6
''                'This records do not have allocation record and they are fully unallocated
''                'Select PI SPlit lines those have the selected fund and osAmount=amount in tlbpursplit table
''                            SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''                                "NOMINAL_CODE,'PI'& P.slnumber,S.NET_AMOUNT,VAT, 0,0,C.OSAmount,0 " & _
''                                "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''                       "S.parentID and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ") and C.ClientID='" & _
''                       rsRentSummaryStatement("ClientIDLandlordID").Value & "' and P.TransactionType=6 AND P.isManagementFee=false AND P.isRentPayable=false " & szWhere
''
''
''                 adoconn.Execute "Insert into ClientStatementPurchasesSnapshot (StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE,PaymentRef," & _
''                           "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''                           SQLforInsert
''
''                 'Type:7
''                   SQLforInsert = "select " & szCurrentStatementID & " as StatementID,P.TransactionType as Type, P.MY_ID,S.TRAN_ID,C.TransactionID,C.ClientID,C.UnitID,C.PDate,C.SageAccountNumber," & _
''                                "NOMINAL_CODE,-S.NET_AMOUNT,-VAT, 0,0,-(C.OSAmount),0 " & _
''                                "from tblPurInv P,tblPurInvSRec S,tlbPayment C, Fund F Where C.PI=P.MY_ID and P.MY_ID= " & _
''                       "S.parentID and osamount<amount and F.FundID=S.dept_ID and F.FundID in (" & rsRentSummaryStatement("ListOfFundId").Value & ")  and C.ClientID='" & rsRentSummaryStatement("ClientIDLandlordID").Value & "' and P.TransactionType=7" & szWhere
''
''
''                 adoconn.Execute "Insert into ClientStatementPurchasesSnapshot(StatementID,Type,MY_ID,SplitID,TransactionID,ClientID,PropertyID,TranDate,SupplierID,NOMINAL_CODE," & _
''                           "NetAmount,VATAmount,PaymentAmount,CreditAmount,osAmount,isManagementFee)" & _
''                           SQLforInsert
''                 '****************end********************************************************
''
''        End If
''        rsRentSummaryStatement.Close
     
     adoconn.Close
    MsgBox "Rent Payable successfully generated", vbInformation, "Saved"
   Unload Me 'Because all value is still not clearing in the form. so reload will clear everything
   Call frmRentPayable.loadflxPayFees("") 'This shall refress mother grid with updates PI info and invoiced info
End Sub
Private Function GetLandLordACBalance(szSelectedClient As String) As Double    'This function return result as minus This is getting LLORD balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND P.ClientID ='" & _
            szSelectedClient & "'  and SP.TYPE in ('LLORD') AND  P.PDate <=#" & Format(txtStatementDate2.text, "dd/mmm/yyyy") & "#"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetLandLordACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function GetClientACBalance() As Double   'This function return result as minus This is getting CLIENT balance
    Dim rsPayment As New ADODB.Recordset
    Dim szSQL As String
    Dim adoconn As New ADODB.Connection
    'F.CategoryCode = 1 Fund category 1 Means rent
    'Implement switch here in SQL
    'Bank code does not exits in PI,so do not put it in where clause
    adoconn.Open getConnectionString
    szSQL = "Select  SUM(SWITCH(P.TYPE =24,S.Amount,P.TYPE =6,S.Amount,P.TYPE =7,-S.Amount,P.TYPE =8,-S.Amount,P.TYPE =9,-S.Amount)) as AMT" & _
            " from tlbPayment P,tlbPaymentSplit S,Fund F,Supplier SP where " & _
            " SP.SupplierID=P.SageaccountNumber AND  P.TransactionID=S.PayHeader AND P.TYPE " & _
            "IN(6,7,8,9,24) AND S.FundID=F.FundID AND P.ClientID ='" & txtClientAccount & "'  and SP.TYPE in ('CLIENT')"
    
    rsPayment.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not rsPayment.EOF Then
        GetClientACBalance = IIf(IsNull(rsPayment.Fields.Item("AMT").Value), 0, rsPayment.Fields.Item("AMT").Value)
    End If
    rsPayment.Close
    adoconn.Close
    Set adoconn = Nothing
End Function
Private Function PI_Check(adoconn As ADODB.Connection, ByRef szTran2Fix As String) As Boolean
       Dim szSQL      As String
       Dim adoRst     As New ADODB.Recordset
      
        
    
       
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
      
        szSQL = "SELECT P.SlNumber, P.TRAN_DATE, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID,  " & _
        "tblPurInvSRec.TRAN_ID, tblPurInvSRec.NOMINAL_CODE, tblPurInvSRec.DESCRIPTION " & _
        "FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID " & _
        "WHERE (((P.TransactionType)=6 Or (P.TransactionType)=7) AND ((P.NLPost)=False))  AND P.TOTAL_AMOUNT<>0 AND  tran_ID is null; "
    'Debug.Print szSQL
    'This means there is PI without splits
          adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
            'if you found this problem exists then there is some transaction in tblPurInv which are inconsitence
            'Now by comparing with the NLposting this transaction will be zerorized
            'To prevent happening this again there is check I have implemented  in PI
             adoRst.MoveNext
          Wend
    
          adoRst.Close
          szSQL = "SELECT  P.SlNumber FROM tblPurinv AS P, (SELECT ParentID, Sum(ROUND(TOTAL_AMOUNT, 2)) AS T From tblPurInvSRec Group by ParentID ) AS Q  " & _
            "WHERE P.MY_ID = Q.ParentID AND round(P.TOTAL_AMOUNT,2) <> round(Q.T,2);"
           adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
             adoRst.MoveNext
          Wend
          
          adoRst.Close
          Set adoRst = Nothing
          'This part shall check for the dupplicate serial number in the tblPurInv
          szSQL = "SELECT slnumber,transactiontype, COUNT(*) From tblPurInv GROUP BY slnumber, transactiontype HAVING COUNT(*) > 1"
          adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
             adoRst.MoveNext
          Wend
          
          adoRst.Close
          Set adoRst = Nothing
     
       If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)
    
       If Len(szTran2Fix) > 0 Then
            PI_Check = False
       Else
            PI_Check = True
            'MsgBox "HI"
       End If
      
End Function
Private Sub cmdDeleteLine_Click()
        flxPayFees.Clear
        Call configflxPayFees
'    If flxPayFees.row = 0 Then
'        MsgBox "Please select a line to remove", vbInformation, "Warning"
'        Exit Sub
'    End If
'    'MsgBox("Do you want to add new Set of payment Dates?", vbQuestion + vbYesNo, "Payment Dates") = vbNo
'    If MsgBox("Are you sure you want to delete line no  " & flxPayFees.row & "", vbYesNo, "") = vbNo Then
'        Exit Sub
'    End If
'    txtPayableAmount.text = txtPayableAmount.text + flxPayFees.TextMatrix(flxPayFees.row, 14)
'    txtPayableAmount.text = Format(txtPayableAmount.text, "0.00")
'    flxPayFees.RemoveItem flxPayFees.row
End Sub

Private Sub flxClientList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call flxClientList_Click
    End If
End Sub

Private Sub flxPayFees_Click()
      txtDescription.text = flxPayFees.TextMatrix(flxPayFees.row, 6)
End Sub

Private Sub Form_Load()
    Dim adoconn As New ADODB.Connection
    If szCurrentStatementID = "" Then Exit Sub
    iRow = 1
    adoconn.Open getConnectionString
    Dim rsStatement As New ADODB.Recordset
    txtStatementNumber2 = "SS" & szCurrentStatementID
    rsStatement.Open "Select * from RentSummaryStatement R,Client C where R.ClientIDLandlordID=C.ClientID AND StatementID=" & _
    Replace(szCurrentStatementID, "SS", "") & "", adoconn, adOpenStatic, adLockReadOnly
    If Not rsStatement.EOF Then
           txtClientAccount.text = rsStatement("ClientID").Value
           txtClientName.text = rsStatement("ClientName").Value
           txtStatementDate2.text = Format(rsStatement("StatementDate").Value, "dd/MM/yyyy")
           txtPayableDate2.text = txtStatementDate2.text
           txtAvailableFund1.text = Format(rsStatement("Availablefunds").Value, "0.00")
           txtPayableAmount.text = txtAvailableFund1.text
'           txtBalancePayable.text = txtAvailableFund1.text
           strPreviousDate = Format(rsStatement("PreviousStatementDate").Value, "dd/MM/yyyy")
           txtReference.text = "SS" & szCurrentStatementID
           strListofFunds = rsStatement("ListOfFundID").Value
           txtStatementBalance.text = Round(rsStatement("StatementClosingBal").Value, 2)
'           strListOfPayableTypeID = rsStatement("ListOfPayableTypeID").Value
    End If
    rsStatement.Close
    Set rsStatement = Nothing
    
        Dim adoRst As New ADODB.Recordset
        'adoconn.Open getConnectionString
        Dim rsConsolidatedStatement As New ADODB.Recordset
        'adoConn.Open getConnectionString
        If txtClientAccount.text <> "" Then
            rsConsolidatedStatement.Open "Select * from client where clientID ='" & txtClientAccount.text & "'", adoconn, adOpenStatic, adLockReadOnly
            If Not rsConsolidatedStatement.EOF Then
                If IsNull(rsConsolidatedStatement("ConsolidatedStatement").Value) Then
                    MsgBox "The Consolidated Client Statement option has not been set for this client. Please set this option on the client record", vbInformation, "Warning"
                End If
                
                boolConsolidatedStatement = IIf(IsNull(rsConsolidatedStatement("ConsolidatedStatement").Value), 0, rsConsolidatedStatement("ConsolidatedStatement"))
            End If
            rsConsolidatedStatement.Close
        End If
        
        
    Call configflxPayFees
    adoconn.Close
    Set adoconn = Nothing
'    txtPayableDate2.text = Format(Now, "dd/MM/yyyy")
    txtPayableDate2.SelStart = 0
    txtPayableDate2.SelLength = Len(txtPayableDate2.text)
    flxPayFees.Clear
    Call configflxPayFees
    Call LoadOnePropertyForThisClient
    If boolConsolidatedStatement = True Then
        cmdProperty.Enabled = False
        txtProperty.Enabled = False
        txtProperty.text = "Consolidated"
    Else
        cmdProperty.Enabled = True
        txtProperty.Enabled = True
    End If
End Sub
Private Sub LoadOnePropertyForThisClient()
    Dim adoconn As New ADODB.Connection
    Dim rsProperty As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsProperty.Open "Select * from Property where clientID='" & txtClientAccount.text & "'", adoconn, adOpenStatic, adLockReadOnly
    If rsProperty.RecordCount = 1 Then
            txtProperty.text = rsProperty("PropertyID").Value
            txtProperty.Tag = rsProperty("PropertyID").Value
    End If
    rsProperty.Close
    Set rsProperty = Nothing
    adoconn.Close
    Set adoconn = Nothing
    
End Sub
Private Sub cmdApply_Click()
'    flxPayFees.ColWidth(2) = 0 'SL
'    flxPayFees.ColWidth(3) = Label12(1).Left - Label6.Left 'Client/ Landlord Account
'    flxPayFees.ColWidth(4) = Label12(2).Left - Label12(1).Left 'Client/Landlord Account Name
'    flxPayFees.ColWidth(5) = Label12(3).Left - Label12(2).Left 'Percentage
'    flxPayFees.ColWidth(6) = Label12(4).Left - Label12(3).Left 'Description
'    flxPayFees.ColWidth(7) = Label12(5).Left - Label12(4).Left 'Reference
'    flxPayFees.ColWidth(8) = Label12(6).Left - Label12(5).Left 'Net Amount
'    flxPayFees.ColWidth(9) = Label12(7).Left - Label12(6).Left 'TAX Code
'    flxPayFees.ColWidth(10) = Label12(8).Left - Label12(7).Left 'VAT
'    flxPayFees.ColWidth(11) = 1200 'Total
'On Error GoTo ERR
        Dim FinalControlACForPayable As String
        Dim iRow As Integer
        If boolConsolidatedStatement = False Then
            If txtProperty.text = "" Then
                 MsgBox "Please enter a Propery", vbInformation, "Warning"
                 FocusControl cmdProperty
                 Exit Sub
            End If
        End If
        If txtPayableDate2.text = "" Then
            MsgBox "Please enter payable dates", vbInformation, "Warning "
            FocusControl txtPayableDate2
             Exit Sub
             
        End If
        If txtFundForPI.text = "" Then
             MsgBox "Please select a Fund", vbInformation, "Warning"
             FocusControl cmdFundListForCreatePI
             Exit Sub
        End If
'        If txtPayableTypes.text = "" Then
'             MsgBox "Please enter a Payable type", vbInformation, "Warning"
'             FocusControl cmdPayableTypes
'             Exit Sub
'        End If
        
        If Val(txtAvailableFund1.text) - Val(txtPayableAmount.text) < 0 Then
            MsgBox "The payable amount entered cannot be greater than available funds", vbInformation, "Warning"
            txtPayableAmount.text = txtAvailableFund1.text
            FocusControl txtPayableAmount
            Exit Sub
        End If
'        If Val(txtAvailableFund1.text) - Val(txtPayableAmount.text) < 0 Then
'            MsgBox "The payable amount entered cannot be greater than available funds.", vbInformation, "Warning"
'            FocusControl txtPayableAmount
'            Exit Sub
'        End If
        If Val(txtAvailableFund1.text) <> Val(txtPayableAmount.text) Then
            MsgBox "The payable amount entered cannot be less than available funds", vbInformation, "Warning"
                txtPayableAmount.text = txtAvailableFund1.text
                 Exit Sub
            
       End If
        If Val(txtPayableAmount.text) <= 0 Then
            MsgBox "Payable amount value should be greater than zero", vbInformation, "Warning"
            FocusControl txtPayableAmount
            Exit Sub
        End If
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        If FinalControlACForPayable = "" Then
            FinalControlACForPayable = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", txtClientAccount.text)
        End If
         adoconn.Close
        Set adoconn = Nothing
        'if still control acount is empty generate a warning message and exit this sub procedure
        If FinalControlACForPayable = "" Then
            MsgBox "Control Account is not set for this Payable type", vbInformation, "Warning"
            Exit Sub
        End If
        iRow = 1
'        For iRow = 1 To flxPayFees.Rows - 1
'            If flxPayFees.TextMatrix(iRow, 10) = txtPayableTypes.Tag Then
'                MsgBox "This payable type is already in the list", vbInformation, "Warning"
'                Exit Sub
'            End If
'        Next
  
        For iRow = 1 To flxPayFees.Rows - 1
            If flxPayFees.TextMatrix(iRow, 3) = txtClientAccount.text Then
                MsgBox "This payable is already in use", vbInformation, "Warning"
                Exit Sub
            End If
        Next
       'tlbPayable is saving Payable information for a property. You can have multiple lines in a rent Payable for a property
        Dim szSQL As String
        Dim rsPayable As New ADODB.Recordset
        Dim dblAmout As Double
        adoconn.Open getConnectionString
        Call configflxPayFees
        If boolConsolidatedStatement = False Then ' for each property write one PI with propertyID
                    szSQL = "SELECT S.SupplierName as CID, T.ID as PID,P.PAYABLE_ID,P.CPA_ID, T.PayType , PAYABLE_TYPE, F.FundID,F.FundName,clientLandlordID, " & _
                                "PAY_START_DATE, PAY_END_DATE,   P.ONDD,P.PAYABLE_BASIS_,PAY_NtDueDate,Percentage,StopDate,PAY_END_DATE,c.PropertyID as PropertyID" & _
                                "FROM tlbPayable AS P, ClientProAgr AS C,  PayableTypes AS T, FUND as F,Supplier S " & _
                                "WHERE  F.FundID=P.PAY_Fund AND P.CPA_ID = C.CPA_ID And S.SupplierID=P.ClientLandlordID and C.ClientID = '" & txtClientAccount.text & "' And " & _
                                "T.ID = P.PAYABLE_TYPE  And C.PropertyID = '" & txtProperty.text & "';" 'Removed 2021/11/09 AND T.ID=" & txtPayableTypes.Tag & "
            'AND T.ID=" & txtPayableTypes.Tag & "
            'Here  PAYABLE_TYPE field is related to the PayableType ID
                rsPayable.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                iRow = 1
                While Not rsPayable.EOF
                    flxPayFees.AddItem ""
                    flxPayFees.TextMatrix(iRow, 1) = ""
                    flxPayFees.TextMatrix(iRow, 2) = iRow
                    flxPayFees.TextMatrix(iRow, 3) = rsPayable("ClientLandlordID").Value  ''Client/ Landlord Account
                    flxPayFees.TextMatrix(iRow, 4) = rsPayable("CID").Value  'Client/Landlord Account Name
                    flxPayFees.TextMatrix(iRow, 5) = rsPayable("PropertyID").Value  'c.PropertyID
                    dblAmout = Format(Val(IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) * txtPayableAmount.text) / 100, "0.00")
                    flxPayFees.TextMatrix(iRow, 6) = IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) 'strPercentage 'Percentage
                    flxPayFees.TextMatrix(iRow, 7) = rsPayable("PayType").Value & " of  £" & dblAmout & " for Statement " & _
                                                     txtStatementNumber2.text & " From " & strPreviousDate & " To " & txtStatementDate2.text & ""   'Description"
                    flxPayFees.TextMatrix(iRow, 8) = txtReference.text  'Reference
                    flxPayFees.TextMatrix(iRow, 9) = txtProperty.text
                    flxPayFees.TextMatrix(iRow, 10) = rsPayable("fundID").Value  'txtFundForPI.Tag
                    flxPayFees.TextMatrix(iRow, 11) = rsPayable("PID").Value 'txtPayableTypes.Tag
                    flxPayFees.TextMatrix(iRow, 12) = 0 'Format(Val(txtPayableAmount.text), "0.00") 'Net Amount
                    flxPayFees.TextMatrix(iRow, 13) = "" 'TAX Code
                    flxPayFees.TextMatrix(iRow, 14) = "" 'VAT
                    flxPayFees.TextMatrix(iRow, 15) = Format(Val(IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) * txtPayableAmount.text) / 100, "0.00") 'Total
                    iRow = iRow + 1
                    rsPayable.MoveNext
                Wend
                rsPayable.Close
                Set rsPayable = Nothing
         Else 'consolidated statement is true  Write property field with no property when
            Dim rsClientLandlordId As New ADODB.Recordset
            
             szSQL = "SELECT Distinct clientLandlordID " & _
                                "FROM tlbPayable AS P, ClientProAgr AS C,  PayableTypes AS T, FUND as F,Supplier S " & _
                                "WHERE   F.FundID=P.PAY_Fund AND P.CPA_ID = C.CPA_ID And S.SupplierID=P.ClientLandlordID and C.ClientID = '" & txtClientAccount.text & "' And " & _
                                "T.ID = P.PAYABLE_TYPE  ;"
              rsClientLandlordId.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
              iRow = 1
            While Not rsClientLandlordId.EOF
                                szSQL = "SELECT S.SupplierName as CID, T.ID as PID,P.PAYABLE_ID,P.CPA_ID, T.PayType , PAYABLE_TYPE, F.FundID,F.FundName,clientLandlordID, " & _
                                "PAY_START_DATE, PAY_END_DATE,   P.ONDD,P.PAYABLE_BASIS_,PAY_NtDueDate,Percentage,StopDate,PAY_END_DATE,c.PropertyID as PropertyID " & _
                                "FROM tlbPayable AS P, ClientProAgr AS C,  PayableTypes AS T, FUND as F,Supplier S " & _
                                "WHERE  F.FundID=P.PAY_Fund AND P.CPA_ID = C.CPA_ID And S.SupplierID=P.ClientLandlordID and C.ClientID = '" & txtClientAccount.text & "' And " & _
                                "T.ID = P.PAYABLE_TYPE  AND P.ClientLandlordID = '" & rsClientLandlordId("clientLandlordID").Value & "';"  'Removed 2021/11/09 AND T.ID=" & txtPayableTypes.Tag & "
                                'Property selection is disabled in this case so I am removing And property filter 'C.PropertyID = '" & txtProperty.text & "'
                                'also show all property in one line (tough job need temporary steps breakdown)
                                'write one PI for
                                'when multiple fund id in the payable exist what I am going to when I comile a PI which fund ID shall I use
            'AND T.ID=" & txtPayableTypes.Tag & "
            'Here  PAYABLE_TYPE field is related to the PayableType ID
                                Dim strline2 As String
                                Dim strline3 As String
                                Dim strline4 As String
                                Dim strline5 As String
                                Dim strline6 As String
                                Dim strline7 As String
                                Dim strline8 As String
                                Dim strline9 As String
                                Dim strline10 As String
                                Dim strline11 As String
                                Dim strline12 As String
                                Dim strline13 As String
                                Dim strline14 As String
                                Dim strline15 As String
            
                            rsPayable.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                            'iRow = 1
                            If Not rsPayable.EOF Then
                                flxPayFees.AddItem ""
                                flxPayFees.TextMatrix(iRow, 1) = ""
                                flxPayFees.TextMatrix(iRow, 2) = iRow
                                strline2 = iRow
                                flxPayFees.TextMatrix(iRow, 3) = rsPayable("ClientLandlordID").Value  ''Client/ Landlord Account
                                strline3 = rsPayable("ClientLandlordID").Value
                                flxPayFees.TextMatrix(iRow, 4) = rsPayable("CID").Value  'Client/Landlord Account Name
                                strline4 = rsPayable("CID").Value
                                'flxPayFees.TextMatrix(iRow, 5) = rsPayable("PropertyID").Value  'c.PropertyID
                                strline5 = rsPayable("PropertyID").Value
                                dblAmout = Format(Val(IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", rsPayable.Fields.Item("Percentage").Value) * txtPayableAmount.text) / 100, "0.00")
                                flxPayFees.TextMatrix(iRow, 6) = IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) 'strPercentage 'Percentage
                                If flxPayFees.TextMatrix(iRow, 6) = 100 Then
                                    flxPayFees.TextMatrix(iRow, 6) = "N/A"
                                End If
                                strline6 = IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) 'strPercentage 'Percentage
                                flxPayFees.TextMatrix(iRow, 7) = rsPayable("PayType").Value & " of  £" & dblAmout & " for Statement " & _
                                                                 txtStatementNumber2.text & " From " & strPreviousDate & " To " & txtStatementDate2.text & ""   'Description"
                                strline7 = rsPayable("PayType").Value & " of  £" & dblAmout & " for Statement " & _
                                                                 txtStatementNumber2.text & " From " & strPreviousDate & " To " & txtStatementDate2.text & ""   'Description"
                                flxPayFees.TextMatrix(iRow, 8) = txtReference.text  'Reference
                                strline8 = txtReference.text  'Reference
                                flxPayFees.TextMatrix(iRow, 9) = "" 'txtProperty.text No property here
                                flxPayFees.TextMatrix(iRow, 10) = rsPayable("fundID").Value  'txtFundForPI.Tag
                                strline9 = rsPayable("fundID").Value
                                flxPayFees.TextMatrix(iRow, 11) = rsPayable("PID").Value 'txtPayableTypes.Tag
                                strline10 = rsPayable("PID").Value
                                flxPayFees.TextMatrix(iRow, 12) = 0 'Format(Val(txtPayableAmount.text), "0.00") 'Net Amount
                                flxPayFees.TextMatrix(iRow, 13) = "" 'TAX Code
                                flxPayFees.TextMatrix(iRow, 14) = "" 'VAT
                                flxPayFees.TextMatrix(iRow, 15) = Format(Val(IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) * txtPayableAmount.text) / 100, "0.00") 'Total
            '                    strline15 = Format(Val(IIf(IsNull(rsPayable.Fields.Item("Percentage").Value), "0.00", Format(rsPayable.Fields.Item("Percentage").Value, "0.00")) * txtPayableAmount.text) / 100, "0.00") 'Total
                                 'flxPayFees.TextMatrix(iRow, 15) = txtPayableAmount.text
                                iRow = iRow + 1
                                rsPayable.MoveNext
                            End If ' Show only one  record for on landlord
                            rsPayable.Close
                            Set rsPayable = Nothing
                rsClientLandlordId.MoveNext 'loop for each landlord
                Wend
                rsClientLandlordId.Close
         End If
    adoconn.Close
'        txtBalancePayable.text = Format(Val(txtBalancePayable.text) - Val(txtPayableAmount.text), "0.00")
    Call calculateTotals
    Exit Sub
Err:
End Sub
Private Sub calculateTotals()
    On Error GoTo Err
      txtTotalAmount.text = "0.00"
    Dim iRow As Integer
    For iRow = 1 To flxPayFees.Rows - 1
            txtTotalAmount.text = txtTotalAmount.text + Val(flxPayFees.TextMatrix(iRow, 15))
    Next
    txtTotalAmount.text = Format(txtTotalAmount.text, "0.00")
    If Round(txtTotalAmount.text - txtPayableAmount.text, 2) = 0.01 Then
        flxPayFees.TextMatrix(1, 15) = Val(flxPayFees.TextMatrix(1, 15)) - 0.01
         txtTotalAmount.text = Format(txtTotalAmount.text - 0.01, "0.00")
     End If
    Exit Sub
Err:
End Sub
Private Sub cmdGridUnitLookup_Click()
        Frame3.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
End Sub

Private Sub cmdPayableTypes_Click()
        Frame1.Enabled = False
        Frame2.Enabled = False
        sTextBox = "PAYABLETYPES"
        Call loadpayableTypes
        Frame3.Top = 2000
        Frame3.Left = 5000
        Frame3.Visible = True
        FocusControl txtSearchClientID
End Sub

Private Sub cmdproperty_Click()
        Frame1.Enabled = False
        Frame2.Enabled = False
        sTextBox = "Property"
        Call LoadProperties
        Frame3.Top = 2070
        Frame3.Left = 2520
        Frame3.Visible = True
        FocusControl txtSearchClientID
End Sub
Private Sub LoadFunds()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoconn As New ADODB.Connection
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
   Dim propertyvar As String
   Dim rsCheck As New ADODB.Recordset
   adoconn.Open getConnectionString
''   szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
''

'drop down was not showing opening balance. So we needed to load the funds from the assigned fund.
'Code written by anol 2023-04-20
    Dim startPos As Integer
    Dim endPos As Integer
    Dim extractedSubstring As String
   rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoconn, adOpenStatic, adLockReadOnly
   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
   Else
        iSel = 1
        rsCheck.Open "Select ListOfinputProperties From RentSummaryStatement where StatementID=" & StrDigitVal(txtStatementNumber2.text) & "", adoconn, adOpenStatic, adLockReadOnly
        If Not rsCheck.EOF Then
            propertyvar = rsCheck("ListOfinputProperties").Value
            
            'Here I am taking first Property ID so that I can select from fund matrix. if I do not take one property multiple fund comes form assigned tables  20-08-2023
            startPos = InStr(propertyvar, "'") + 1 ' Find the position after the first single quote
            endPos = InStr(startPos, propertyvar, "'") ' Find the position of the second single quote
            If startPos > 0 And endPos > startPos Then
                   propertyvar = Mid(propertyvar, startPos, endPos - startPos)
            End If
        End If
        rsCheck.Close
         szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND M.PropertyID = '" & propertyvar & "' and isDeleted=false"
'        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
'                szPropertySelecti & "' and ClientID='" & txtClientID & "' and isDeleted=false"
   End If
''   rsFundMatrix.Close
 Rem by anol 2023-04-20
'     szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund where FundID in(" & strListofFunds & ")"
     
     
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
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
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

Private Sub flxClientList_Click()
    Frame3.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    If sTextBox = "Property" Then
            txtProperty.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
            txtProperty.text = flxClientList.TextMatrix(flxClientList.row, 1)
            FocusControl cmdFundListForCreatePI
    ElseIf sTextBox = "Fund" Then
            txtFundForPI.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
            txtFundForPI.text = flxClientList.TextMatrix(flxClientList.row, 2)
            FocusControl txtReference
    ElseIf sTextBox = "PAYABLETYPES" Then
            txtPayableTypes.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
            txtPayableTypes.text = flxClientList.TextMatrix(flxClientList.row, 2)
            FocusControl txtReference
          
    End If
    
End Sub

'Private Sub flxClientList_Click()
'    Frame5.Visible = False
'    If sTextBox = "Fund" Then
'        txtFundForPI.Tag = flxClientList.TextMatrix(flxClientList.row, 1)
'        txtFundForPI.text = flxClientList.TextMatrix(flxClientList.row, 2)
'        FocusControl cmdCreatePI
'    End If
'End Sub
Private Sub cmdFundListForCreatePI_Click()
    Frame1.Enabled = False
    Frame2.Enabled = False
    sTextBox = "Fund"
    Call LoadFunds
    Frame3.Top = 2070
    Frame3.Left = 4000
    Frame3.Visible = True
    FocusControl txtSearchClientID
    
End Sub
Private Sub LoadflxClientList()
        Call ConfigflxClientList
        Dim adoconn As New ADODB.Connection
        Dim rstRec As New ADODB.Recordset
        Dim szSQL As String
        
        Dim rRow As Integer
        adoconn.Open getConnectionString
        
        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
        rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

        rRow = 1
        While Not rstRec.EOF
                flxClientList.TextMatrix(rRow, 0) = ""
                flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundID").Value
                flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundCode").Value
                flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("FundName").Value
                flxClientList.RowHeight(rRow) = 280
                rstRec.MoveNext
                If Not rstRec.EOF Then flxClientList.AddItem ""
                rRow = rRow + 1
        Wend

        rstRec.Close
        adoconn.Close
        Set rstRec = Nothing
        Set adoconn = Nothing
End Sub
Private Sub cmdFrameFundClose_Click()
    Frame3.Visible = False
End Sub
Private Sub ConfigflxClientList()
        flxClientList.Clear
        Dim szHeader As String
        szHeader$ = "|<FundID|<FundCode|<FundName"
        flxClientList.FormatString = szHeader$
    
        flxClientList.Cols = 4
        flxClientList.Rows = 2
        flxClientList.RowHeight(0) = 0
        flxClientList.ColWidth(0) = 250   'Selection Row put plus or minus sign
        flxClientList.ColWidth(1) = 0 'FundID
        flxClientList.ColWidth(2) = 2000 ' FundCode
        flxClientList.ColWidth(3) = 2000  ' FundName
        flxClientList.ColAlignment(0) = vbLeftJustify
        flxClientList.ColAlignment(1) = vbLeftJustify
        flxClientList.ColAlignment(2) = vbLeftJustify
        flxClientList.ColAlignment(3) = vbLeftJustify
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub loadpayableTypes()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
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
   
   

   adoconn.Open getConnectionString
   If txtProperty.text = "N/A" Then
        szSQL = "SELECT ID, PayType FROM PayableTypes ;"
   Else
        szSQL = "SELECT ID, PayType FROM PayableTypes where ClientID='" & txtClientAccount.text & "'  "
   End If
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadProperties()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
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
   
   lblClientID(0).Caption = " Property ID"
   lblClientID(1).Caption = " Property Name"
   lblClientID(2).Caption = " "
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""

   adoconn.Open getConnectionString
   'Property will only   come from the selected input
   szSQL = "SELECT  PropertyId,PropertyName From Property P   where PropertyID in ( Select PropertyID from  RentSummaryStatement " & _
            "R where StatementID =" & szCurrentStatementID & " )AND ClientID='" & txtClientAccount.text & "';"
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        rRow = 1
        'You cannot add N/A beause on loading this main grid you have a property fileter
'        flxClientList.AddItem ""
'        flxClientList.TextMatrix(rRow, 1) = "N/A"
'        flxClientList.TextMatrix(rRow, 2) = "N/A"
'        rRow = 2
        While Not rstRec.EOF
            flxClientList.row = 1
            flxClientList.RowSel = 1
            flxClientList.ColSel = 1
            flxClientList.TextMatrix(rRow, 0) = ""
            flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("PropertyId").Value
            flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("PropertyName").Value
            'flxClientList.TextMatrix(rRow, 3) = rstRec.Fields.Item("SType").Value
            flxClientList.RowHeight(rRow) = 280
            rstRec.MoveNext
            If Not rstRec.EOF Then flxClientList.AddItem ""
            rRow = rRow + 1
         Wend
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub configflxPayFees()
    flxPayFees.Clear
    flxPayFees.Cols = 12
    flxPayFees.Rows = 2
    flxPayFees.RowHeight(0) = 0
    flxPayFees.ColWidth(0) = 0
    flxPayFees.ColWidth(1) = 200 'Selection Column
    flxPayFees.ColWidth(2) = 400 'Serial NL
    flxPayFees.ColWidth(3) = Label12(3).Left - Label12(2).Left 'Client/ Landlord Account
    flxPayFees.ColWidth(4) = 3000 ' Label12(4).Left - Label12(3).Left 'Client/Landlord Account Name
    flxPayFees.ColWidth(5) = 0 '1600 'Label12(4).Left - Label12(3).Left 'PropertyID
    flxPayFees.ColWidth(6) = 1600 'Label12(5).Left - Label12(4).Left 'Percentage
    flxPayFees.ColWidth(7) = Label12(6).Left - Label12(5).Left  'Description
    flxPayFees.ColWidth(8) = 2000 'Label12(6).Left - Label12(5).Left '0 'Label12(7).Left - Label12(6).Left 'Reference
    flxPayFees.ColWidth(9) = 0  'Property ID
    flxPayFees.ColWidth(10) = 0  'Fund ID
    flxPayFees.ColWidth(11) = 0 'Payable Type ID
    flxPayFees.ColWidth(12) = 0 ' Label12(8).Left - Label12(7).Left 'Net Amount
    flxPayFees.ColWidth(13) = 0 'Label12(9).Left - Label12(8).Left 'TAX Code
    flxPayFees.ColWidth(14) = 0 ' Label12(10).Left - Label12(9).Left 'VAT
    flxPayFees.ColWidth(15) = 2400 'Total
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MousePointer = vbArrow
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MousePointer = vbArrow
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       MousePointer = vbArrow
End Sub

Private Sub TextBox1_Change()
     Dim i As Integer
    If Len(TextBox1.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
        flxClientList.RowHeight(i) = 240
        If InStr(1, UCase(flxClientList.TextMatrix(i, 3)), UCase(TextBox1.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
        End If
        If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
        End If
   Next i
End Sub

Private Sub TextBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        FocusControl flxClientList
    End If
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl flxClientList
    End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdCreatePI
        txtDescription.text = flxPayFees.TextMatrix(flxPayFees.row, 6)
    End If
End Sub
Private Sub txtPayableDate2_Change()
   TextBoxChangeDate txtPayableDate2
   End Sub
Private Sub txtPayableDate2_GotFocus()
   SelTxtInCtrl txtPayableDate2
End Sub

Private Sub txtPayableDate2_LostFocus()
    TextBoxFormatDate txtPayableDate2
End Sub
Private Sub txtPayableAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdApply
    End If
    DigitTextKeyPress txtPayableAmount, KeyAscii
End Sub

Private Sub txtPayableAmount_LostFocus()
    txtPayableAmount.text = Format(txtPayableAmount.text, "0.00")
End Sub

Private Sub txtPayableDate2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdProperty.Enabled Then
                FocusControl cmdProperty
         Else
                FocusControl cmdFundListForCreatePI
         End If
    End If
    DigitTextKeyPress txtPayableDate2, KeyAscii
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdApply
        txtPayableAmount.SelStart = 0
        txtPayableAmount.SelLength = Len(txtPayableAmount.text)
    End If
    
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
End Sub

Private Sub txtSearchClientName_Change()
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
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If TextBox1.Visible = False And KeyCode = 13 Then
        FocusControl flxClientList
    End If
End Sub

