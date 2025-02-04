VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDemandTypes_ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Types"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemandTypes_.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCommands 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      TabIndex        =   36
      Top             =   3600
      Width           =   9255
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000009&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   495
         Index           =   0
         Left            =   8100
         TabIndex        =   44
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3480
         TabIndex        =   43
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2460
         TabIndex        =   42
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   1440
         TabIndex        =   41
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   495
         Left            =   420
         TabIndex        =   40
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel Changes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6855
         TabIndex        =   39
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5850
         TabIndex        =   38
         Top             =   2760
         Width           =   930
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   4845
         TabIndex        =   37
         Top             =   2760
         Width           =   930
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypes 
         Height          =   2265
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   3995
         _Version        =   393216
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
            Name            =   "Calibri"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Height          =   210
         Index           =   20
         Left            =   420
         TabIndex        =   49
         Top             =   0
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Type"
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
         Left            =   1140
         TabIndex        =   48
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
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
         Index           =   22
         Left            =   3540
         TabIndex        =   47
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
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
         Height          =   210
         Index           =   23
         Left            =   5940
         TabIndex        =   46
         Top             =   0
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   735
         Index           =   2
         Left            =   4785
         Top             =   2640
         Width           =   3075
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   735
         Index           =   0
         Left            =   360
         Top             =   2640
         Width           =   4125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   735
         Index           =   1
         Left            =   360
         Top             =   2640
         Width           =   4125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   735
         Index           =   3
         Left            =   4785
         Top             =   2640
         Width           =   3075
      End
   End
   Begin VB.Frame fraDemandType 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   80
      Width           =   9135
      Begin VB.CommandButton cmdAddEmailTemplate 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8695
         Picture         =   "frmDemandTypes_.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3200
         Width           =   275
      End
      Begin VB.TextBox txtPrefix 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7995
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdAddReport 
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8695
         Picture         =   "frmDemandTypes_.frx":A06B4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2855
         Width           =   275
      End
      Begin VB.TextBox txtDemandTemplate 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   6435
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2855
         Width           =   2270
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000016&
         Caption         =   "Payment Dates:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   60
         TabIndex        =   22
         Top             =   2160
         Width           =   5205
         Begin MSForms.ComboBox cboDemandTypePayDates 
            Height          =   315
            Left            =   2385
            TabIndex        =   11
            Top             =   540
            Width           =   2745
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "4842;556"
            TextColumn      =   2
            ColumnCount     =   2
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1058"
         End
         Begin MSForms.OptionButton optPreset 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   495
            Width           =   2175
            VariousPropertyBits=   746588179
            BackColor       =   12632256
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3836;661"
            Value           =   "0"
            Caption         =   "Use Preset Payment Dates"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optAuto 
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   195
            Width           =   2535
            VariousPropertyBits=   746588179
            BackColor       =   12632256
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4471;661"
            Value           =   "1"
            Caption         =   "Use Automatic Payment Dates"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txt4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5635
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Width           =   3335
      End
      Begin VB.TextBox txt3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5635
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   3335
      End
      Begin VB.TextBox txt2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5635
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   3335
      End
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5635
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   3335
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6360
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2835
         MaxLength       =   40
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
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
         Index           =   6
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   720
      End
      Begin MSForms.ComboBox cboProperty 
         Height          =   315
         Left            =   2835
         TabIndex        =   1
         Top             =   0
         Width           =   6165
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "10874;556"
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   3
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;7761"
      End
      Begin MSForms.TextBox txtEmailTemplate 
         Height          =   285
         Left            =   6435
         TabIndex        =   34
         Top             =   3200
         Width           =   2265
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "4004;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Tmpt:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   5400
         TabIndex        =   33
         Top             =   3200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prefix:"
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
         Left            =   7440
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin MSForms.ComboBox cboGroup 
         Height          =   285
         Left            =   6795
         TabIndex        =   13
         Top             =   2175
         Width           =   2175
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3836;503"
         TextColumn      =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   5400
         TabIndex        =   31
         Top             =   2175
         Width           =   705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Report:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   5400
         TabIndex        =   30
         Top             =   2855
         Width           =   750
      End
      Begin MSForms.CheckBox chkGroup 
         Height          =   255
         Left            =   6435
         TabIndex        =   12
         Top             =   2175
         Width           =   255
         VariousPropertyBits=   746588179
         BackColor       =   16764879
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "450;450"
         Value           =   "0"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5400
         TabIndex        =   29
         Top             =   2520
         Width           =   600
      End
      Begin MSForms.ComboBox cboDemandTypeCategory 
         Height          =   315
         Left            =   2835
         TabIndex        =   8
         Top             =   1800
         Width           =   2775
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4895;556"
         TextColumn      =   1
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;5000"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Category:"
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
         Index           =   5
         Left            =   60
         TabIndex        =   28
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5640
         TabIndex        =   27
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Demand Type:"
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
         Index           =   1
         Left            =   60
         TabIndex        =   26
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code for Total Amount:"
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
         Index           =   3
         Left            =   60
         TabIndex        =   25
         Top             =   1440
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code for VAT Amount:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   24
         Top             =   1080
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code for Amount:"
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
         Index           =   0
         Left            =   60
         TabIndex        =   23
         Top             =   720
         Width           =   2145
      End
      Begin MSForms.ComboBox cboDemandTypeNCTotal 
         Height          =   315
         Left            =   2835
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4895;556"
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1234;5000"
      End
      Begin MSForms.ComboBox cboDemandTypeNCvat 
         Height          =   315
         Left            =   2835
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4895;556"
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1234;5000"
      End
      Begin MSForms.ComboBox cboDemandTypeNCAmt 
         Height          =   315
         Left            =   2835
         TabIndex        =   5
         Top             =   720
         Width           =   2775
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4895;556"
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1234;5000"
      End
      Begin MSForms.ComboBox cboBank 
         Height          =   315
         Left            =   6435
         TabIndex        =   14
         Top             =   2520
         Width           =   2535
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4471;556"
         TextColumn      =   2
         ColumnCount     =   4
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;7408;1587;1411"
      End
   End
End
Attribute VB_Name = "frmDemandTypes_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
           As Long) As Long
Const SW_SHOW = 5

Dim szDemandId          As String
Dim szDemandType        As String
Dim bAddNew             As Boolean
Dim iSelRow             As Integer
Dim bSortingCol1        As Boolean
Dim bSortingCol2        As Boolean
Dim szExistingProperty  As String

Public Sub GetRecord()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, " & _
               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
               "CategoryCode, PaymentDates, Spare1, DTGroup, DemandReportName, " & _
               "EmailInvoiceTemplate " & _
            "FROM DemandTypes " & _
            "WHERE ID = " & szDemandId & ""
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   txtType.text = adoRst!Type
   txtID.text = adoRst!ID
   If IsNull(adoRst!NominalCodeforAmount) = False Then cboDemandTypeNCAmt.text = adoRst!NominalCodeforAmount
   If IsNull(adoRst!NominalNameforAmount) = False Then txt1.text = adoRst!NominalNameforAmount
   If IsNull(adoRst!NominalCodeForVAT) = False Then cboDemandTypeNCvat.text = adoRst!NominalCodeForVAT
   If IsNull(adoRst!NominalNameforVAT) = False Then txt2.text = adoRst!NominalNameforVAT
   If IsNull(adoRst!NominalCodeForTotal) = False Then cboDemandTypeNCTotal.text = adoRst!NominalCodeForTotal
   If IsNull(adoRst!NominalNameforTotal) = False Then txt3.text = adoRst!NominalNameforTotal
   txtPrefix.text = IIf(IsNull(adoRst!prefix), "", adoRst!prefix)
   

   If IsNull(adoRst!CategoryCode) = False Then cboDemandTypeCategory.Value = adoRst!CategoryCode
   If adoRst!PaymentDates = 255 Then
      optAuto.Value = True
   Else
      optPreset.Value = True
      cboDemandTypePayDates.Value = CInt(adoRst!PaymentDates)
   End If
   If Val(IIf(IsNull(adoRst!DTGroup), 0, adoRst!DTGroup)) > 0 Then
      cboGroup.Value = IIf(IsNull(adoRst!DTGroup), 0, adoRst!DTGroup)
      chkGroup.Value = True
   Else
      chkGroup.Value = False
   End If
   If IsNull(adoRst!DemandReportName) Then
      txtDemandTemplate.text = ""
   Else
      txtDemandTemplate.text = adoRst!DemandReportName
   End If
   If Not IsNull(adoRst!spare1) Then cboBank.Value = adoRst!spare1
   txtEmailTemplate.text = IIf(IsNull(adoRst!EmailInvoiceTemplate), "", adoRst!EmailInvoiceTemplate)

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cboDemandTypeCategory_Change()
   On Error GoTo ErorrHandler

   txt4.text = cboDemandTypeCategory.Column(1)

   Exit Sub
ErorrHandler:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboDemandTypeNCAmt_Change()
   On Error GoTo ErorrHandler

   txt1.text = cboDemandTypeNCAmt.Column(1)

   Exit Sub
ErorrHandler:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboDemandTypeNCTotal_Change()
   On Error GoTo ErorrHandler

   txt3.text = cboDemandTypeNCTotal.Column(1)

   Exit Sub
ErorrHandler:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboDemandTypeNCvat_Change()
   On Error GoTo ErorrHandler

   txt2.text = cboDemandTypeNCvat.Column(1)

   Exit Sub
ErorrHandler:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Function SecCodeValue(szPrimaryCode As String, szCode As String) As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString

   szSQL = "SELECT Value " & _
            "FROM SecondaryCode " & _
            "WHERE PrimaryCode = '" & szPrimaryCode & "' AND Code = '" & szCode & "'"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   SecCodeValue = adoRst!Value
   adoRst.Close
   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub cboDemandTypePayDates_Click()
   optPreset.Value = 1
End Sub

Private Sub cboGroup_GotFocus()
   If cboDemandTypeCategory.text = "" Then
      MsgBox "Please select the Demand Category first.", vbCritical + vbOKOnly, "Demand Type"
      cboDemandTypeCategory.SetFocus
   End If
End Sub

Private Sub cboProperty_Change()
   If cboProperty.ListIndex = -1 Then
      If cboProperty.text <> "" Then cboProperty.text = ""
      Exit Sub
   End If

   If szExistingProperty <> "" And cboProperty.text <> "" And cmdSave.Enabled Then
      If cboProperty.Column(0) = "ALL" Then
         MsgBox "Do not set this demand type to all properties.", vbCritical + vbOKOnly, "Demand Type"
         cboProperty.SetFocus
         cboProperty.ListIndex = -1
      End If
   End If
End Sub

Private Sub cboProperty_GotFocus()
   If cboProperty.ListIndex = -1 Then Exit Sub

   If cmdSave.Enabled Then
      If MsgBox("Do you want to change the Property?", _
                 vbQuestion + vbYesNo, _
                "Demand Type - Allocation") = vbYes Then
         cboProperty.SetFocus
         If cboProperty.ListIndex >= 0 Then
            szExistingProperty = cboProperty.Column(0)
         Else
            szExistingProperty = ""
         End If
         Exit Sub
      End If
      cmdSave.SetFocus
   End If
End Sub

Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If cboProperty.ListIndex = -1 Then
      If cboProperty.text <> "" Then cboProperty.text = ""
      KeyAscii = 0
   End If
End Sub

Private Sub cboProperty_LostFocus()
   If cboProperty.ListIndex = -1 Then Exit Sub

   If szExistingProperty <> "" And cboProperty.text <> "" And cmdSave.Enabled Then
      If cboProperty.Column(0) = "ALL" Then
         MsgBox "Do not set this demand type to all properties.", vbCritical + vbOKOnly, "Demand Type"
         cboProperty.SetFocus
         cboProperty.ListIndex = -1
      End If

      If szExistingProperty <> cboProperty.Column(0) Then
         Dim adoConn As New ADODB.Connection
         Dim adoRst  As New ADODB.Recordset
         Dim szSQL   As String

         adoConn.Open getConnectionString

         If cboDemandTypeCategory.Value = 1 Then                              'Rent Charge
            szSQL = "SELECT PropertyName " & _
                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
                          "FROM   DemandTypes AS T, LRentCharges AS R, " & _
                               "LeaseDetails AS L, Units AS U, Property AS P " & _
                          "WHERE  T.ID = R.BRDemandType AND R.LeaseID = L.LeaseID AND " & _
                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
                    "WHERE Q.PropertyID <> '" & cboProperty.Column(0) & "';"
'Debug.Print szSQL
         End If
         If cboDemandTypeCategory.Value = 2 Then                              'Service Charge
            szSQL = "SELECT PropertyName " & _
                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
                          "FROM   DemandTypes AS T, LServiceCharges AS S, " & _
                               "LeaseDetails AS L, Units AS U, Property AS P " & _
                          "WHERE  T.ID = S.SCDemandType AND S.LeaseID = L.LeaseID AND " & _
                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
                    "WHERE Q.PropertyID <> '" & cboProperty.Column(0) & "';"
'Debug.Print szSQL
         End If
         If cboDemandTypeCategory.Value = 3 Then                              'Insurance Charge
            szSQL = "SELECT PropertyName " & _
                    "FROM (SELECT DISTINCT U.PropertyID, P.PropertyName " & _
                          "FROM   DemandTypes AS T, LInsuranceCharges AS I, " & _
                               "LeaseDetails AS L, Units AS U, Property AS P " & _
                          "WHERE  T.ID = I.InsuranceDemandType AND I.LeaseID = L.LeaseID AND " & _
                               "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                               "T.ID = " & flxDemandTypes.TextMatrix(flxDemandTypes.row, 1) & ") AS Q " & _
                    "WHERE Q.PropertyID <> '" & cboProperty.Column(0) & "';"
'Debug.Print szSQL
         End If
         
            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            
            If Not adoRst.EOF Then
               szSQL = ""
               While Not adoRst.EOF
                  szSQL = adoRst.Fields.Item(0).Value
                  adoRst.MoveNext
                  If Not adoRst.EOF Then szSQL = szSQL & ", "
               Wend
               
               MsgBox "This demand type is being used in leases by " & szSQL & "." & Chr(13) & _
                      "Please reschedule the demand type in the lease first.", vbCritical + vbOKOnly, "Demand Types"
               cboProperty.ListIndex = -1
            End If
            adoRst.Close
         

         Set adoRst = Nothing
         adoConn.Close
         Set adoConn = Nothing
      End If
   End If
End Sub

Private Sub chkGroup_Click()
   If chkGroup.Value Then
      cboGroup.Enabled = True
      If cmdSave.Enabled Or cmdSaveNew.Enabled Then cboGroup.SetFocus
   Else
      cboGroup.ListIndex = -1
      cboGroup.Enabled = False
   End If
End Sub

Private Sub cmdAdd_Click()
   Call AddNewDemandType
   flxDemandTypes.Enabled = False
   txtType.SetFocus
   szDemandId = ""
   cboProperty.ListIndex = -1
   cboProperty.SetFocus
End Sub

Public Sub AddNewDemandType()
   Call EmptyBoxes

   cmdDelete.Enabled = False
   cmdEdit.Enabled = False
   cmdSaveNew.Enabled = True
   cmdSaveNew.ZOrder 0
   cmdCancelNew.Enabled = True
   cmdAdd.Enabled = False

'   Call EnableBoxes
   fraDemandType.Enabled = True
   chkGroup.Enabled = True
End Sub

Private Sub cmdAddEmailTemplate_Click()
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
   ofn.lpstrInitialDir = CurDir
   ofn.lpstrTitle = "Select a Report file"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   txtEmailTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
End Sub

Private Sub cmdAddReport_Click()
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
   ofn.lpstrInitialDir = CurDir
   ofn.lpstrTitle = "Select a Report file"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   txtDemandTemplate.text = JustifyFilePath(ofn.lpstrFileTitle)
End Sub

Private Sub cmdCancel_Click()
   Call GetRecord
'   Call DisableBoxes
   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   szExistingProperty = ""
End Sub

Private Sub cmdCancelNew_Click()
   Call EmptyBoxes
'   Call DisableBoxes
   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False
End Sub

Private Sub cmdDelete_Click()
   Dim Response As Byte
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   
   If szDemandId = "" Then
       MsgBox "You must select a demand type to delete!", vbOKOnly + vbCritical, "No demand type selected"
       Exit Sub
   End If
   Response = MsgBox("Are you sure you want to delete demand type: " & szDemandType, vbYesNo + vbQuestion, "Delete")
   If Response = vbNo Then Exit Sub

   'delete record.
   adoConn.Open getConnectionString
   
   szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, " & _
               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
               "CategoryCode, PaymentDates " & _
            "FROM DemandTypes WHERE ID = " & szDemandId & ""
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
   
   adoRst.Delete
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadFlxDemandTypes(adoConn)
   
   adoConn.Close
   Set adoConn = Nothing

   Call EmptyBoxes
End Sub

Private Sub cmdEdit_Click()
   If szDemandId = "" Then
      MsgBox "You must select a demand type to edit", vbOKOnly + vbCritical, "No demand type selected"
      Exit Sub
   End If

'   If Left(lblProperty.Caption, 3) = "All Properties" Then
'      MsgBox "Please select property in the previous screen.", vbCritical + vbOKOnly, "Demand Type"
'      Exit Sub
'   End If

'   lblClient.Caption = frmDCTypesPre.cboClientList.Column(1)
'   lblProperty.Caption = frmDCTypesPre.cboPropertyList.Column(1)

'   Call EnableBoxes
   fraDemandType.Enabled = True
   flxDemandTypes.Enabled = False

   cmdEdit.Enabled = False
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   chkGroup.Enabled = True
   txtType.SetFocus
End Sub

Private Function IsBankEdit() As Boolean
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String ', szaDemandID() As String

   adoConn.Open getConnectionString

   szSQL = "SELECT DR.DemandId " & _
           "FROM DemandRecords as DR, DemandSplitRecords as DSR " & _
           "WHERE DSR.TypeOfDemand = " & szDemandId & " AND " & _
               "(DR.IsPrinted = FALSE OR " & _
               "DR.UPDATE_SAGE = FALSE) AND " & _
               "DR.DemandHistory = FALSE AND " & _
               "DR.DemandId = DSR.DemandId;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      IsBankEdit = True
   Else
      IsBankEdit = False
   End If

   adoRst.Close
   adoConn.Close
   Set adoRst = Nothing
   Set adoConn = Nothing
End Function

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
End Sub
'
'Private Sub cmdGSCancel_Click(Index As Integer)
'   If MsgBox("Are you sure to cancel the changes?", vbQuestion + vbYesNo, "Cancel Saving") = vbNo Then Exit Sub
'   ButtonMode DefaultMode, Index
'End Sub

Private Sub cmdSave_Click()
   On Error Resume Next
   
   If txtPrefix.text = "" Then
      MsgBox "You must enter Demand Prefix", vbOKOnly + vbCritical, "Demand Type"
      txtPrefix.SetFocus
      Exit Sub
   End If
   If cboDemandTypeNCAmt.text = "" Then
      MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
      cboDemandTypeNCAmt.SetFocus
      Exit Sub
   End If
   If cboDemandTypeNCAmt.text <> "" And txt1.text = "" Then
      MsgBox "You must select a correct Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
      cboDemandTypeNCAmt.text = ""
      cboDemandTypeNCAmt.SetFocus
      Exit Sub
   End If

   If szDemandId <> 4 Then
      If cboDemandTypeNCvat.text = "" Then
         MsgBox "You must select a Nominal Account for VAT Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         cboDemandTypeNCvat.SetFocus
         Exit Sub
      End If
   End If

   If cboDemandTypeNCTotal.text = "" Then
      MsgBox "You must select a Nominal Account for Total Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
      cboDemandTypeNCTotal.SetFocus
      Exit Sub
   End If

   If cboDemandTypeCategory.text = "" Then
      MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
      cboDemandTypeCategory.SetFocus
      Exit Sub
   End If

   If cboDemandTypePayDates.text = "" And optPreset.Value Then
      MsgBox "You must select a demand payment date.", vbOKOnly + vbCritical, "Payment Date"
      cboDemandTypePayDates.SetFocus
      Exit Sub
   End If

   If cboBank.text = "" Then
      If cboBank.ListCount = 1 Then
         cboBank.ListIndex = 0
      Else
         MsgBox "You must select a bank details.", vbOKOnly + vbCritical, "Bank Details"
         cboBank.SetFocus
         Exit Sub
      End If
   End If

   If txtDemandTemplate.text = "" Then
      MsgBox "You must enter a demand template file name.", vbOKOnly + vbCritical, "Demand Template"
      cmdAddReport.SetFocus
      Exit Sub
   End If

   If txtEmailTemplate.text = "" Then
      MsgBox "You must enter a demand email template.", vbOKOnly + vbCritical, "Demand Email Template"
      cmdAddEmailTemplate.SetFocus
      Exit Sub
   End If

   If cboProperty.text = "" Or cboProperty.ListIndex = -1 Then
      MsgBox "You must select a property.", vbOKOnly + vbCritical, "Demand Type property"
      cboProperty.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoConn.Open getConnectionString
   'UPDATE RECORD
   szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, " & _
               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
               "CategoryCode, PaymentDates, DTGroup, DemandReportName, Spare1, " & _
               "PropertyID, EmailInvoiceTemplate " & _
            "FROM DemandTypes WHERE ID =" & szDemandId & ""
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   adoRst!Type = txtType.text
   adoRst!prefix = txtPrefix.text
   adoRst!NominalCodeforAmount = cboDemandTypeNCAmt.text
   adoRst!NominalNameforAmount = txt1.text
   adoRst!NominalCodeForVAT = cboDemandTypeNCvat.text
   adoRst!NominalNameforVAT = txt2.text
   adoRst!NominalCodeForTotal = cboDemandTypeNCTotal.text
   adoRst!NominalNameforTotal = txt3.text

   adoRst!CategoryCode = cboDemandTypeCategory.Value
   If optAuto.Value Then
      adoRst!PaymentDates = CByte(255)
   Else
      adoRst!PaymentDates = CByte(cboDemandTypePayDates.Value)
   End If
   If chkGroup.Value Then adoRst!DTGroup = cboGroup.Value
   adoRst!DemandReportName = txtDemandTemplate.text
   adoRst!spare1 = cboBank.Value

'   If flxDemandTypes.TextMatrix(iSelRow, 5) = "ALL" And cboProperty.Column(0) <> "ALL" Then
      adoRst!PropertyID = cboProperty.Column(0)
'   End If

   adoRst!EmailInvoiceTemplate = txtEmailTemplate.text

   adoRst.Update
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadFlxDemandTypes(adoConn)

   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Your changes have been saved."

'   Call DisableBoxes
   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   iSelRow = 0
   szExistingProperty = ""
End Sub

Private Sub cmdSaveNew_Click()
   Dim a As Integer
   
   If txtType.text = "" Then
      MsgBox "You must enter Demand Type", vbOKOnly + vbCritical, "Demand Type"
      txtType.SetFocus
      Exit Sub
   End If
   If txtPrefix.text = "" Then
      MsgBox "You must enter Demand Prefix", vbOKOnly + vbCritical, "Demand Type"
      txtPrefix.SetFocus
      Exit Sub
   End If
   If cboDemandTypeNCAmt.text = "" Then
      MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
      cboDemandTypeNCAmt.SetFocus
      Exit Sub
   End If
   If cboDemandTypeNCvat.text = "" Then
       MsgBox "You must select a Nominal Account for VAT Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
       cboDemandTypeNCvat.SetFocus
       Exit Sub
   End If
   If cboDemandTypeNCTotal.text = "" Then
      MsgBox "You must select a Nominal Account for Total Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
      cboDemandTypeNCTotal.SetFocus
      Exit Sub
   End If
   If cboDemandTypeCategory.text = "" Then
      MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
      cboDemandTypeCategory.SetFocus
      Exit Sub
   End If
   If cboDemandTypePayDates.text = "" And optPreset.Value Then
      MsgBox "You must select a demand payment date.", vbOKOnly + vbCritical, "Payment Date"
      cboDemandTypePayDates.SetFocus
      Exit Sub
   End If
   If cboBank.text = "" Then
      If cboBank.ListCount = 1 Then
         cboBank.ListIndex = 0
      Else
         MsgBox "You must select a bank details.", vbOKOnly + vbCritical, "Bank Details"
         cboBank.SetFocus
         Exit Sub
      End If
   End If
   If txtDemandTemplate.text = "" Then
      MsgBox "You must enter a demand template file name.", vbOKOnly + vbCritical, "Demand Template"
      cmdAddReport.SetFocus
      Exit Sub
   End If
   If txtEmailTemplate.text = "" Then
      MsgBox "You must enter a demand email template file name.", vbOKOnly + vbCritical, "Demand Email Template"
      txtEmailTemplate.SetFocus
      Exit Sub
   End If
   If chkGroup.Value = 1 And cboGroup.text = "" Then
      MsgBox "You must select a group id the demand type.", vbOKOnly + vbCritical, "Group"
      cboGroup.SetFocus
      Exit Sub
   End If

   Dim szSQL As String
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   adoConn.Open getConnectionString

   szSQL = "SELECT MAX(ID) FROM DemandTypes"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   a = 0
   a = CInt(IIf(IsNull(adoRst.Fields.Item(0).Value), 0, adoRst.Fields.Item(0).Value))
   
   adoRst.Close

   txtID.text = a + 1

   szSQL = "SELECT ID, Type, Prefix, NominalCodeforAmount, InvCrd, " & _
               "NominalNameforAmount, NominalCodeforVAT, NominalNameforVAT, " & _
               "NominalCodeforTotal, NominalNameforTotal, TransactionType, " & _
               "CategoryCode, PaymentDates, DTGroup, DemandReportName, " & _
               "Spare1, PropertyID, EmailInvoiceTemplate " & _
            "FROM DemandTypes"
   adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

   adoRst.AddNew
   adoRst!ID = txtID.text
'  ---------------------------------------------------
'  'InvCrd' field is required in the table. This field is no longer in use.
'  Thats why I am saving a charecter. I have changed the field as not required (03/09/2009).
   adoRst!InvCrd = "X"
   adoRst!Type = txtType.text
   adoRst!prefix = txtPrefix.text
   adoRst!NominalCodeforAmount = cboDemandTypeNCAmt.text
   adoRst!NominalNameforAmount = txt1.text
   adoRst!NominalCodeForVAT = cboDemandTypeNCvat.text
   adoRst!NominalNameforVAT = txt2.text
   adoRst!NominalCodeForTotal = cboDemandTypeNCTotal.text
   adoRst!NominalNameforTotal = txt3.text
   adoRst!prefix = "NULL"
   adoRst!CategoryCode = cboDemandTypeCategory.Value
   If optAuto.Value Then
      adoRst!PaymentDates = CByte(255)
   Else
      adoRst!PaymentDates = CByte(cboDemandTypePayDates.Value)
   End If
   If chkGroup.Value = 1 Then adoRst!DTGroup = cboGroup.Value
   adoRst!DemandReportName = txtDemandTemplate.text
   adoRst!spare1 = cboBank.Value
   adoRst!PropertyID = cboProperty.Column(0)
   adoRst!EmailInvoiceTemplate = txtEmailTemplate.text

   adoRst.Update
   
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadFlxDemandTypes(adoConn)
   
   adoConn.Close
   Set adoConn = Nothing

   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False

   fraDemandType.Enabled = False
   flxDemandTypes.Enabled = True

   cmdAdd.SetFocus

   ShowMsgInTaskBar "Your new demand type details have been saved."
End Sub

Private Sub flxDemandTypes_RowColChange()
   With flxDemandTypes
      szDemandId = .TextMatrix(.row, 1)
      szDemandType = .TextMatrix(.row, 2)
      iSelRow = .row
      cboProperty.text = .TextMatrix(.row, 6)
      Me.Caption = "Demand Types - " & cboProperty.text
   End With

   EmptyBoxes
   Call GetRecord
End Sub

Private Sub Form_Load()
   Me.Width = 9435
   Me.Height = 7530
   Me.Top = 0
   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   fraCommands.BackColor = MODULEBACKCOLOR
   fraDemandType.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR
   fraDemandType.Top = 120
   fraDemandType.Left = 120
MsgBox "1"
   chkGroup.Enabled = False

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
MsgBox "2"
   LoadNCinCombo adoConn         'all nominal code and name are collecting in all combos from sage
MsgBox "3"
   LoadFlxDemandTypes adoConn
MsgBox "4"
   LoadPaymentDates adoConn
MsgBox "5"
   LoadDemandCategory adoConn  'all category
MsgBox "6"
   LoadGroup adoConn
MsgBox "7"
   LoadBankDetails adoConn      'all clints' bank details
MsgBox "8"
   LoadProperty adoConn
MsgBox "9"
   adoConn.Close
   Set adoConn = Nothing
MsgBox "10"
   Call WheelHook(Me.hWnd)
MsgBox "11"
End Sub

Private Sub LoadProperty(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer, i As Integer, j As Integer

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProPostCode " & _
           "FROM Property " & _
           "ORDER BY PropertyID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   ReDim Data(TotalCol - 1, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   For i = 1 To TotalRow
      For j = 0 To TotalCol - 1
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadGroup(adoConn As ADODB.Connection)
   cboGroup.Clear

   Dim szSQL As String, iSt As Integer, iEnd As Integer
   Dim adoRst As New ADODB.Recordset
   Dim i As Integer

   szSQL = "SELECT CODE, VALUE " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'GR';"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      If adoRst.Fields.Item("Code").Value = "ENDRNG" Then
         iEnd = adoRst.Fields.Item("VALUE").Value
      Else
         iSt = adoRst.Fields.Item("VALUE").Value
      End If
      adoRst.MoveNext
   Wend

   For i = iSt To iEnd
      cboGroup.AddItem i
   Next i

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadBankDetails(adoConn As ADODB.Connection)
   cboBank.Clear

   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim adoRst As New ADODB.Recordset

   szSQL = "SELECT My_ID, Bank_AC_Name, BANK_AC_NUM, BANK_SC " & _
               "FROM tlbClientBanks;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount < 1 Then
      MsgBox "There are no client's bank details has been setup.", vbCritical + vbOKOnly, "Bank Details Missing"
      cmdCancelNew_Click
      Exit Sub
   End If

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   Dim Data() As String
   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To adoRst.RecordCount - 1
      For j = 0 To adoRst.Fields.count - 1
         Data(j, i) = adoRst.Fields(j)
      Next j
      adoRst.MoveNext
   Next i

   cboBank.Column() = Data()
   
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadDemandCategory(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer
   
   cboDemandTypeCategory.Clear

   szSQL = "SELECT Code, Value " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'DCTG';"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount < 1 Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If
 
   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To adoRst.RecordCount - 1
       For j = 0 To adoRst.Fields.count - 1
           Data(j, i) = adoRst.Fields(j)
       Next j
       adoRst.MoveNext
   Next i

   cboDemandTypeCategory.Column() = Data()
   
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadPaymentDates(adoConn As ADODB.Connection)
   cboDemandTypePayDates.Clear

   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim Data() As String
   Dim TotalRow, TotalCol As Integer

   ReDim Data(1, 0) As String

   Data(0, 0) = "0"
   Data(1, 0) = "DEFAULT"

   szSQL = "SELECT NameOfSet " & _
               "FROM PaymentDates " & _
               "ORDER BY DateSetID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount = 0 Then
      cboDemandTypePayDates.Column() = Data()
      
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   TotalRow = adoRst.RecordCount

   ReDim Data(1, TotalRow) As String
   Dim i As Integer

   adoRst.MoveFirst
   Data(0, 0) = "0"
   Data(1, 0) = "DEFAULT"
   For i = 1 To adoRst.RecordCount
      Data(0, i) = i
      Data(1, i) = adoRst.Fields(0).Value
      adoRst.MoveNext
   Next i

   cboDemandTypePayDates.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadChargeType(adoConn As ADODB.Connection)
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim TotalRow, TotalCol As Integer
   Dim Data() As String

   szSQL = "SELECT ID, FeeType, FeeIC, FeeSagePrefix, FeeNCAmt, FeeNNAmt, " & _
                  "FeeNCVat, FeeNNVat, FeeNCTotal, FeeNNTotal, TransactionType, " & _
                  "CategoryCode, PaymentDates, RecoverableExp " & _
                "FROM ChargeTypes ORDER BY ID"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount < 1 Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadNCinCombo(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, TotalRow As Integer
   Dim Data() As String, i As Integer

   szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   TotalRow = adoRst.RecordCount
   ReDim Data(2, TotalRow) As String

   i = 0
   While Not adoRst.EOF
      Data(0, i) = adoRst.Fields.Item("Code").Value
      Data(1, i) = adoRst.Fields.Item("Name").Value
      i = i + 1
      adoRst.MoveNext
   Wend

   cboDemandTypeNCAmt.Column() = Data()
   cboDemandTypeNCvat.Column() = Data()
   cboDemandTypeNCTotal.Column() = Data()

   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub ConfigFlxDemandTypes()
   With flxDemandTypes
      .Cols = 7
      .RowHeight(0) = 0

      .ColWidth(0) = Label1(20).Left - .Left                      '
      .ColWidth(1) = Label1(21).Left - Label1(20).Left            'ID
      .ColWidth(2) = Label1(22).Left - Label1(21).Left            'Demand Type
      .ColWidth(3) = 0                                            'Client ID
      .ColWidth(4) = Label1(23).Left - Label1(22).Left            'Client Name
      .ColWidth(5) = 0                                            'Property ID
      .ColWidth(6) = .Left + .Width - Label1(23).Left - 350       'Property Name
   End With
End Sub

Public Sub LoadFlxDemandTypes(adoConn As ADODB.Connection)
   Dim Data() As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim szSQL As String, szHeader As String
   Dim adoRst As New ADODB.Recordset

'   If frmDCTypesPre.cboPropertyList.Column(0) <> "ALL" Then
'      szSQL = "SELECT D.ID, D.Type, D.PropertyID, " & _
'                     "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
'                     "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
'                     "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
'              "FROM (DemandTypes AS D LEFT JOIN Property AS P ON " & _
'                    "D.PropertyID = P.PropertyID) LEFT JOIN Client AS C ON P.ClientID = C.ClientID " & _
'              "WHERE D.PropertyID = '" & frmDCTypesPre.cboPropertyList.Column(0) & "' OR " & _
'                    "D.PropertyID = 'ALL' " & _
'              "ORDER BY D.ID;"
'   Else
   szSQL = "SELECT D.ID, D.Type, D.PropertyID, " & _
                  "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                  "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID, " & _
                  "IIF(ISNULL(P.ClientID), '', C.ClientName) AS ClientName " & _
           "FROM (DemandTypes AS D LEFT JOIN Property AS P ON " & _
                 "D.PropertyID = P.PropertyID) LEFT JOIN Client AS C ON P.ClientID = C.ClientID " & _
           "ORDER BY D.ID;"
'   End If
'Debug.Print szSQL

   flxDemandTypes.Clear
   flxDemandTypes.Rows = 2
   szHeader$ = "|<ID|<Type|<ClientID|<ClientName|<PropertyID|<PropertyName"
   flxDemandTypes.FormatString = szHeader$

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockOptimistic

   With adoRst.Fields
      For i = 0 To adoRst.RecordCount - 1
         flxDemandTypes.TextMatrix(i + 1, 1) = IIf(IsNull(.Item("ID")), "", .Item("ID"))
         flxDemandTypes.TextMatrix(i + 1, 2) = IIf(IsNull(.Item("Type")), "", .Item("Type"))
         flxDemandTypes.TextMatrix(i + 1, 3) = IIf(IsNull(.Item("ClientID")), "", .Item("ClientID"))
         flxDemandTypes.TextMatrix(i + 1, 4) = IIf(IsNull(.Item("ClientName")), "", .Item("ClientName"))
         flxDemandTypes.TextMatrix(i + 1, 5) = IIf(IsNull(.Item("PropertyID")), "", .Item("PropertyID"))
         flxDemandTypes.TextMatrix(i + 1, 6) = IIf(IsNull(.Item("PropertyName")), "", .Item("PropertyName"))

         adoRst.MoveNext
         If Not adoRst.EOF Then flxDemandTypes.AddItem ""
      Next i
   End With
   ConfigFlxDemandTypes
   flxDemandTypes.row = 0
   iSelRow = 0

NoRes:
   adoRst.Close
End Sub

Public Sub EmptyBoxes()
   cboDemandTypeNCAmt.text = ""
   cboDemandTypeNCvat.text = ""
   cboDemandTypeNCTotal.text = ""
   cboDemandTypeCategory.text = ""
   cboBank.text = ""
   cboGroup.text = ""
   txtType.text = ""
   txtID.text = ""
   txtPrefix.text = ""
   txtDemandTemplate.text = ""
   txtEmailTemplate.text = ""

   chkGroup.Value = False
   txt1.text = ""
   txt2.text = ""
   txt3.text = ""
   txt4.text = ""
End Sub

Public Sub EnableBoxes()
   cmdAdd.Enabled = False

   cboDemandTypeNCAmt.Enabled = True
   cboDemandTypeNCvat.Enabled = True
   cboDemandTypeNCTotal.Enabled = True
   cboDemandTypeCategory.Enabled = True
   cboBank.Enabled = True

   cmdAddReport.Enabled = True
   cmdAddEmailTemplate.Enabled = True
   chkGroup.Enabled = True
End Sub

Public Sub DisableBoxes()
   cmdAdd.Enabled = True

   cboDemandTypeNCAmt.Enabled = False
   cboDemandTypeNCvat.Enabled = False
   cboDemandTypeNCTotal.Enabled = False
   cboDemandTypeCategory.Enabled = False
   cboBank.Enabled = False

   cmdAddReport.Enabled = False
   cmdAddEmailTemplate.Enabled = False
   chkGroup.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub fraCommands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub fraDemandType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Label1_Click(Index As Integer)
   If Index = 20 Then                                       ' Sort ID
      SortingGrid flxDemandTypes, 1, bSortingCol1, "Integer"
      bSortingCol1 = IIf(bSortingCol1, False, True)
      Label1(20).FontBold = True
      Label1(23).FontBold = False
   End If
   If Index = 23 Then                                       ' Sort Property
      SortingGrid flxDemandTypes, 5, bSortingCol2
      bSortingCol2 = IIf(bSortingCol2, False, True)
      Label1(20).FontBold = False
      Label1(23).FontBold = True
   End If
   flxDemandTypes.row = 0
End Sub

Private Sub optAuto_Click()
   cboDemandTypePayDates.Enabled = Not optAuto.Value
End Sub

Private Sub optPreset_Click()
   cboDemandTypePayDates.Enabled = optPreset.Value
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
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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
