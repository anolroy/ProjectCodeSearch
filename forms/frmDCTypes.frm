VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDCTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmDCTypes"
   ClientHeight    =   12735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15090
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDCTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12735
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDemandType 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton cmdAddEmailTemplate 
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
         Left            =   8695
         Picture         =   "frmDCTypes.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3960
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
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdAddReport 
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
         Left            =   8695
         Picture         =   "frmDCTypes.frx":9F734
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3615
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
         TabIndex        =   25
         Top             =   3615
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
         Height          =   1335
         Left            =   60
         TabIndex        =   47
         Top             =   2880
         Width           =   5325
         Begin MSForms.ComboBox cboDemandTypePayDates 
            Height          =   315
            Left            =   2385
            TabIndex        =   17
            Top             =   780
            Width           =   2865
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "5054;556"
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
            TabIndex        =   16
            Top             =   735
            Width           =   2415
            VariousPropertyBits=   746588179
            BackColor       =   12632256
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4260;661"
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
            TabIndex        =   15
            Top             =   315
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
         TabIndex        =   46
         Top             =   2520
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
         TabIndex        =   45
         Top             =   2160
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
         TabIndex        =   44
         Top             =   1800
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
         TabIndex        =   43
         Top             =   1440
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
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtType 
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
         Left            =   2835
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   4965
         TabIndex        =   26
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5970
         TabIndex        =   23
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel Changes"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6975
         TabIndex        =   24
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   495
         Left            =   540
         TabIndex        =   3
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2580
         TabIndex        =   5
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3600
         TabIndex        =   6
         Top             =   4440
         Width           =   930
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000009&
         Cancel          =   -1  'True
         Caption         =   "&Close Screen"
         Height          =   495
         Index           =   0
         Left            =   8040
         TabIndex        =   2
         Top             =   4440
         Width           =   975
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
         Index           =   0
         Left            =   -120
         ScaleHeight     =   45
         ScaleWidth      =   10035
         TabIndex        =   1
         Top             =   780
         Width           =   10095
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
         Index           =   18
         Left            =   4320
         TabIndex        =   102
         Top             =   65
         Width           =   720
      End
      Begin VB.Label lblProperty 
         BackStyle       =   0  'Transparent
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
         Left            =   5160
         TabIndex        =   101
         Top             =   80
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   100
         Top             =   65
         Width           =   495
      End
      Begin VB.Label lblClient 
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   99
         Top             =   80
         Width           =   3495
      End
      Begin MSForms.TextBox txtEmailTemplate 
         Height          =   285
         Left            =   6435
         TabIndex        =   98
         Top             =   3960
         Width           =   2270
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "4004;503"
         SpecialEffect   =   0
         FontName        =   "Calibri"
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
         TabIndex        =   97
         Top             =   3960
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
         TabIndex        =   57
         Top             =   1080
         Width           =   495
      End
      Begin MSForms.ComboBox cboGroup 
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Top             =   2880
         Width           =   735
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1296;503"
         TextColumn      =   1
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "352"
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
         TabIndex        =   56
         Top             =   2880
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
         TabIndex        =   55
         Top             =   3615
         Width           =   750
      End
      Begin MSForms.CheckBox chkGroup 
         Height          =   255
         Left            =   6435
         TabIndex        =   18
         Top             =   2880
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
      Begin MSForms.ComboBox cboBank 
         Height          =   315
         Left            =   6435
         TabIndex        =   20
         Top             =   3240
         Width           =   2535
         VariousPropertyBits=   746604569
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
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;7408;1587;1411"
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
         TabIndex        =   54
         Top             =   3240
         Width           =   600
      End
      Begin MSForms.ComboBox cboDemandTypeCategory 
         Height          =   315
         Left            =   2835
         TabIndex        =   14
         Top             =   2520
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
         TabIndex        =   53
         Top             =   2520
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
         TabIndex        =   52
         Top             =   1080
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
         TabIndex        =   51
         Top             =   1080
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
         TabIndex        =   50
         Top             =   2160
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
         TabIndex        =   49
         Top             =   1800
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
         TabIndex        =   48
         Top             =   1440
         Width           =   2145
      End
      Begin MSForms.ComboBox cboDemandTypeNCTotal 
         Height          =   315
         Left            =   2835
         TabIndex        =   13
         Top             =   2160
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
         TabIndex        =   12
         Top             =   1800
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
         TabIndex        =   11
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   735
         Index           =   0
         Left            =   480
         Top             =   4320
         Width           =   4125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Demand Type:"
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
         Index           =   8
         Left            =   60
         TabIndex        =   27
         Top             =   360
         Width           =   1650
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   735
         Index           =   2
         Left            =   4905
         Top             =   4320
         Width           =   3075
      End
      Begin MSForms.ComboBox cboDemand 
         Height          =   315
         Left            =   2835
         TabIndex        =   7
         Top             =   360
         Width           =   6165
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "10874;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   5
         ListRows        =   20
         cColumnInfo     =   5
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705;4233;0;4937;0"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   735
         Index           =   1
         Left            =   480
         Top             =   4320
         Width           =   4125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   735
         Index           =   3
         Left            =   4905
         Top             =   4320
         Width           =   3075
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   80
         Top             =   20
         Width           =   8945
      End
   End
   Begin VB.Frame fraPayableType 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5175
      Left            =   9990
      TabIndex        =   69
      Top             =   5850
      Width           =   9495
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
         Index           =   2
         Left            =   -240
         ScaleHeight     =   45
         ScaleWidth      =   10035
         TabIndex        =   82
         Top             =   540
         Width           =   10095
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000009&
         Caption         =   "&Close Screen"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   7680
         TabIndex        =   81
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdGSCancel 
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
         Height          =   495
         Index           =   1
         Left            =   4080
         TabIndex        =   80
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSEdit 
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
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   71
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSSave 
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
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   79
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Frame fraMain 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2895
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Width           =   9255
         Begin VB.TextBox txtPayCate 
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
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   1920
            Width           =   4095
         End
         Begin MSForms.Label Label10 
            Height          =   255
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   1095
            BackColor       =   16768960
            VariousPropertyBits=   8388627
            Caption         =   "Payable Type:"
            Size            =   "1931;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPayType 
            Height          =   315
            Left            =   1440
            TabIndex        =   73
            Top             =   0
            Width           =   2985
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "5265;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for Amt:"
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
            Left            =   0
            TabIndex        =   93
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for VAT:"
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
            Left            =   0
            TabIndex        =   92
            Top             =   960
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for Total:"
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
            Left            =   0
            TabIndex        =   91
            Top             =   1440
            Width           =   930
         End
         Begin MSForms.ComboBox cboPayNCAmt 
            Height          =   315
            Left            =   1440
            TabIndex        =   74
            Top             =   480
            Width           =   3015
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5318;556"
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
         Begin MSForms.ComboBox cboPayNCVat 
            Height          =   315
            Left            =   1440
            TabIndex        =   75
            Top             =   960
            Width           =   3015
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5318;556"
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
         Begin MSForms.ComboBox cboPayNCTotal 
            Height          =   315
            Left            =   1440
            TabIndex        =   76
            Top             =   1440
            Width           =   3015
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5318;556"
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
         Begin MSForms.TextBox txtPayNCAmt 
            Height          =   315
            Left            =   4920
            TabIndex        =   90
            Top             =   480
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPayNCVat 
            Height          =   315
            Left            =   4920
            TabIndex        =   89
            Top             =   960
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPayNCTotal 
            Height          =   315
            Left            =   4920
            TabIndex        =   88
            Top             =   1440
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Demand Category:"
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
            Left            =   30
            TabIndex        =   87
            Top             =   1920
            Width           =   1290
         End
         Begin MSForms.ComboBox cboPayableTypeCategory 
            Height          =   315
            Left            =   1440
            TabIndex        =   77
            Top             =   1920
            Width           =   3015
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5318;556"
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
         Begin MSForms.ComboBox cboPayableTypePayDates 
            Height          =   315
            Left            =   1440
            TabIndex        =   78
            Top             =   2400
            Width           =   3015
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5318;556"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Dates:"
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
            TabIndex        =   86
            Top             =   2400
            Width           =   1080
         End
      End
      Begin MSForms.CommandButton cmdAddNew 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   84
         Top             =   4320
         Width           =   1095
         Caption         =   "Add New"
         Size            =   "1931;873"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rent PayableType:"
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
         Left            =   1800
         TabIndex        =   83
         Top             =   120
         Width           =   1275
      End
      Begin MSForms.ComboBox cboPayChrgType 
         Height          =   315
         Left            =   3480
         TabIndex        =   72
         Top             =   120
         Width           =   4005
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7056;556"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;5706"
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   735
         Left            =   120
         Top             =   4200
         Width           =   5295
      End
   End
   Begin VB.Frame fraChargeType 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5175
      Left            =   240
      TabIndex        =   28
      Top             =   6000
      Width           =   9375
      Begin VB.Frame fraMain 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3135
         Index           =   0
         Left            =   135
         TabIndex        =   58
         Top             =   810
         Width           =   9135
         Begin VB.TextBox txtFeeDemandCategory 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   1920
            Width           =   4095
         End
         Begin MSForms.CheckBox chkRecoverableExp 
            Height          =   255
            Left            =   7440
            TabIndex        =   96
            Top             =   2520
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
            Caption         =   "Recoverable Expenses:"
            Height          =   195
            Index           =   6
            Left            =   4800
            TabIndex        =   95
            Top             =   2520
            Width           =   2025
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for Amt:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   68
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for VAT:"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   67
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "N.C. for Total:"
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   66
            Top             =   1440
            Width           =   1695
         End
         Begin MSForms.ComboBox cboFeeNCAmt 
            Height          =   315
            Left            =   1560
            TabIndex        =   38
            Top             =   480
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;556"
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1234;5000"
         End
         Begin MSForms.ComboBox cboFeeNCVat 
            Height          =   315
            Left            =   1560
            TabIndex        =   39
            Top             =   960
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;556"
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1234;5000"
         End
         Begin MSForms.ComboBox cboFeeNCTotal 
            Height          =   315
            Left            =   1560
            TabIndex        =   40
            Top             =   1440
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;556"
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1234;5000"
         End
         Begin MSForms.TextBox txtFeeType 
            Height          =   315
            Left            =   1560
            TabIndex        =   37
            Top             =   0
            Width           =   2865
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "5054;556"
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label9 
            Height          =   195
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   1455
            BackColor       =   16768960
            VariousPropertyBits=   8388627
            Caption         =   "Charge Type:"
            Size            =   "2566;344"
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFeeNCAmt 
            Height          =   315
            Left            =   4800
            TabIndex        =   64
            Top             =   480
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFeeNCVat 
            Height          =   315
            Left            =   4800
            TabIndex        =   63
            Top             =   960
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtFeeNCTotal 
            Height          =   315
            Left            =   4800
            TabIndex        =   62
            Top             =   1440
            Width           =   4095
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "7223;556"
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboChargeTypeCategory 
            Height          =   315
            Left            =   1560
            TabIndex        =   41
            Top             =   1920
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;556"
            TextColumn      =   1
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1058;5000"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Demand Category:"
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   61
            Top             =   1920
            Width           =   1680
         End
         Begin MSForms.ComboBox cboChargeTypePayDates 
            Height          =   315
            Left            =   1560
            TabIndex        =   42
            Top             =   2400
            Width           =   2895
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5106;556"
            TextColumn      =   2
            ColumnCount     =   2
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Calibri"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1058"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Dates:"
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   60
            Top             =   2400
            Width           =   1365
         End
      End
      Begin VB.CommandButton cmdGSCancel 
         Caption         =   "Canc&el"
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
         Height          =   495
         Index           =   0
         Left            =   4080
         TabIndex        =   33
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSEdit 
         Caption         =   "&Edit"
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
         Height          =   495
         Index           =   0
         Left            =   1440
         TabIndex        =   31
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSSave 
         Caption         =   "&Save"
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
         Height          =   495
         Index           =   0
         Left            =   2790
         TabIndex        =   32
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000009&
         Caption         =   "&Close Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   7560
         TabIndex        =   35
         Top             =   4320
         Width           =   1455
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
         Index           =   1
         Left            =   -240
         ScaleHeight     =   45
         ScaleWidth      =   9915
         TabIndex        =   29
         Top             =   540
         Width           =   9975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         Height          =   735
         Left            =   120
         Top             =   4200
         Width           =   5295
      End
      Begin MSForms.CommandButton cmdAddNew 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   4320
         Width           =   1095
         Caption         =   "Add New"
         Size            =   "1931;873"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.ComboBox cboFeeChrgType 
         Height          =   315
         Left            =   3480
         TabIndex        =   36
         Top             =   120
         Width           =   4005
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7056;556"
         BoundColumn     =   2
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;5000"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fee Charge Type:"
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
         Left            =   1800
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDCTypes"
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

Dim szDemandId As String
Dim szDemandType As String
Dim bAddNew As Boolean

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
   txtID.text = adoRst!Id
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
   txtEmailTemplate.text = adoRst!EmailInvoiceTemplate

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cboChargeTypeCategory_Change()
   On Error GoTo NOCODE

   txtFeeDemandCategory.text = cboChargeTypeCategory.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboDemand_Click()
   If cboDemand.text = "" Then
       MsgBox "You must select a Demand Type!", vbOKOnly + vbExclamation, "No Demand Type Selected"
       Exit Sub
   End If

   szDemandId = cboDemand.Column(0)
   szDemandType = cboDemand.Column(1)
   lblProperty.Caption = cboDemand.Column(3)
   lblClient.Caption = cboDemand.Column(4)

   EmptyBoxes
   Call GetRecord
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

Private Sub cboDemandTypePayDates_Change()
   optPreset.Value = 1
End Sub

Private Sub cboFeeChrgType_Change()
   If cboFeeChrgType.ListCount < 1 Then Exit Sub
   If cboFeeChrgType.ListIndex = cboFeeChrgType.ListCount - 1 Then Exit Sub

   ButtonMode DefaultMode, 0

   txtFeeType.text = cboFeeChrgType.Column(1)
   cboFeeNCAmt.text = cboFeeChrgType.Column(4)
   cboFeeNCVat.text = cboFeeChrgType.Column(6)
   cboFeeNCTotal.text = cboFeeChrgType.Column(8)
   cboChargeTypeCategory.Value = cboFeeChrgType.Column(11)
   cboChargeTypePayDates.Value = cboFeeChrgType.Column(12)
   chkRecoverableExp.Value = IIf(cboFeeChrgType.Column(13) = "Y", True, False)
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

Private Sub cboFeeNCAmt_Change()
   On Error GoTo NOCODE

   txtFeeNCAmt.text = cboFeeNCAmt.Column(1)

   Exit Sub
NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboFeeNCTotal_Change()
   On Error GoTo NOCODE

   txtFeeNCTotal.text = cboFeeNCTotal.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboFeeNCVat_Change()
   On Error GoTo NOCODE

   txtFeeNCVat.text = cboFeeNCVat.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboGroup_GotFocus()
   If cboDemandTypeCategory.text = "" Then
      MsgBox "Please select the Demand Category first.", vbCritical + vbOKOnly, "Demand Type"
      cboDemandTypeCategory.SetFocus
   End If
End Sub

Private Sub cboPayableTypeCategory_Change()
   On Error GoTo NOCODE

   txtPayCate.text = cboPayableTypeCategory.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboPayChrgType_Change()
   If cboPayChrgType.ListCount < 1 Then Exit Sub
   If cboPayChrgType.ListIndex = cboPayChrgType.ListCount - 1 Then Exit Sub

   ButtonMode DefaultMode, 1

   txtPayType.text = cboPayChrgType.Column(1)
   cboPayNCAmt.text = cboPayChrgType.Column(4)
   cboPayNCVat.text = cboPayChrgType.Column(6)
   cboPayNCTotal.text = cboPayChrgType.Column(8)
   cboPayableTypeCategory.text = cboPayChrgType.Column(11)
   cboPayableTypePayDates.Value = cboPayChrgType.Column(12)
End Sub

Private Sub cboPayNCAmt_Change()
   On Error GoTo NOCODE

   txtPayNCAmt.text = cboPayNCAmt.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboPayNCTotal_Change()
   On Error GoTo NOCODE

   txtPayNCTotal.text = cboPayNCTotal.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub cboPayNCVat_Change()
   On Error GoTo NOCODE

   txtPayNCVat.text = cboPayNCVat.Column(1)

   Exit Sub

NOCODE:
   MsgBox "Code does not exists", vbCritical + vbOKOnly, "Wrong Code"
End Sub

Private Sub chkGroup_Click()
   If chkGroup.Value Then
      cboGroup.Enabled = True
      If cmdSave.Enabled Or cmdSaveNew.Enabled Then cboGroup.SetFocus
   Else
      cboGroup.Enabled = False
   End If
End Sub

Private Sub cmdAdd_Click()
   If Left(lblProperty.Caption, 3) = "All Properties" Then
      MsgBox "Please select property in the previous screen.", vbCritical + vbOKOnly, "Demand Type"
      Exit Sub
   End If

   Call AddNewDemandType
   txtType.SetFocus
End Sub

Public Sub AddNewDemandType()
   Call EmptyBoxes

   cmdDelete.Enabled = False
   cmdEdit.Enabled = False
   cmdSaveNew.Enabled = True
   cmdSaveNew.ZOrder 0
   cmdCancelNew.Enabled = True

   Call EnableBoxes
   fraDemandType.Enabled = True
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

Private Sub cmdAddNew_Click(Index As Integer)
   If MsgBox("Do you want to add new types?", vbQuestion + vbYesNo, "Add New") = vbNo Then Exit Sub
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   
   fraMain(Index).Enabled = True
   cmdAddNew(Index).Enabled = False
   cmdGSEdit(Index).Enabled = False
   cmdGSSave(Index).Enabled = True
   cmdGSCancel(Index).Enabled = True
   If Index = 0 Then
      txtFeeType.SetFocus
   Else
      txtPayType.SetFocus
   End If
   ButtonMode NewEntryMode, Index
   adoConn.Close
   Set adoConn = Nothing
   bAddNew = True
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
   Call DisableBoxes
'   fraDemandType.Enabled = False
   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
End Sub

Private Sub cmdCancelNew_Click()
   Call EmptyBoxes
   Call DisableBoxes
'   fraDemandType.Enabled = False
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
   
   If cboDemand.text = "" Then
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
   
   Call LoadDemandTypes(adoConn)
   
   adoConn.Close
   Set adoConn = Nothing

   Call EmptyBoxes
End Sub

Private Sub cmdEdit_Click()
   If cboDemand.text = "" Then
      MsgBox "You must select a demand type to edit", vbOKOnly + vbCritical, "No demand type selected"
      Exit Sub
   End If
   
   If Left(lblProperty.Caption, 3) = "All Properties" Then
      MsgBox "Please select property in the previous screen.", vbCritical + vbOKOnly, "Demand Type"
      Exit Sub
   End If

   lblClient.Caption = frmDCTypesPre.cboClientList.Column(1)
   lblProperty.Caption = frmDCTypesPre.cboPropertyList.Column(1)

   Call EnableBoxes
   fraDemandType.Enabled = True

   cmdEdit.Enabled = False
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
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

Private Sub cmdGSCancel_Click(Index As Integer)
   If MsgBox("Are you sure to cancel the changes?", vbQuestion + vbYesNo, "Cancel Saving") = vbNo Then Exit Sub
   ButtonMode DefaultMode, Index
End Sub

Private Sub cmdGSEdit_Click(Index As Integer)
   If Index = 0 Then
      If cboFeeChrgType.text = "" Then
         MsgBox "Please select the Fee Charge Type.", vbCritical + vbOKOnly, "Fee Charge Type"
         Exit Sub
      End If
   Else
      If cboPayChrgType.text = "" Then
         MsgBox "Please select the Rent Payable Type.", vbCritical + vbOKOnly, "Rent Payable Type"
         Exit Sub
      End If
   End If
   Dim szTemp As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   ButtonMode EditMode, Index

   If Index = 0 Then
      txtFeeType.SetFocus
   Else
      txtPayType.SetFocus
   End If
   adoConn.Close
   Set adoConn = Nothing
   bAddNew = False
End Sub

Private Sub cmdGSSave_Click(Index As Integer)
   If Index = 0 Then
      If txtFeeType.text = "" Then
         MsgBox "You must type the Fee/Charge Type.", vbOKOnly + vbCritical, "Fee/Charge Type"
         Exit Sub
      End If
      If cboFeeNCAmt.text = "" Then
         MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboFeeNCVat.text = "" Then
         MsgBox "You must select a Nominal Account for VAT Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboFeeNCTotal.text = "" Then
         MsgBox "You must select a Nominal Account for Total Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboChargeTypeCategory.text = "" Then
          MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
          Exit Sub
      End If
      If cboChargeTypePayDates.text = "" Then
         MsgBox "You must select a charge type payment date.", vbOKOnly + vbCritical, "Charge Type Payment Date"
         Exit Sub
      End If

      Dim adoConn As New ADODB.Connection
      Dim adoRst As New ADODB.Recordset
      Dim szSQL As String
      
      adoConn.Open getConnectionString

      If bAddNew Then
         szSQL = "SELECT * " & _
                  "FROM ChargeTypes"
      Else
         szSQL = "SELECT * " & _
                  "FROM ChargeTypes WHERE ID =" & cboFeeChrgType.Column(0) & ""
      End If
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      If bAddNew Then
         adoRst.AddNew
         adoRst!Id = cboFeeChrgType.ListCount + 1
      End If

      adoRst!FeeType = txtFeeType.text
      adoRst!FeeIC = "I"                           'for nothing i have set the value.. but dont use it
      adoRst!FeeNCAmt = cboFeeNCAmt.text
      adoRst!FeeNNAmt = cboFeeNCAmt.Column(1)
      adoRst!FeeNCVat = cboFeeNCVat.text
      adoRst!FeeNNVat = cboFeeNCVat.Column(1)
      adoRst!FeeNCTotal = cboFeeNCTotal.text
      adoRst!FeeNNTotal = cboFeeNCTotal.Column(1)
      adoRst!FeeSagePrefix = "NULL"
      adoRst!CategoryCode = cboChargeTypeCategory.Value
      adoRst!PaymentDates = CByte(cboChargeTypePayDates.Value)
      adoRst!RecoverableExp = IIf(chkRecoverableExp.Value, "Y", "N")
      adoRst.Update
      adoRst.Close
      Set adoRst = Nothing
      
      LoadChargeType adoConn
      
      adoConn.Close
      Set adoConn = Nothing

      ShowMsgInTaskBar "Your changes have been saved."
      ButtonMode DefaultMode, Index
   End If

   If Index = 1 Then
      If txtPayType.text = "" Then
         MsgBox "You must type the Payable Type.", vbOKOnly + vbCritical, "Payable Type"
         Exit Sub
      End If
      If cboPayNCAmt.text = "" Then
         MsgBox "You must select a Nominal Account for Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboPayNCVat.text = "" Then
         MsgBox "You must select a Nominal Account for VAT Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboPayNCTotal.text = "" Then
         MsgBox "You must select a Nominal Account for Total Amount", vbOKOnly + vbCritical, "No Nominal Account Selected"
         Exit Sub
      End If
      If cboPayableTypeCategory.text = "" Then
          MsgBox "You must select a demand category.", vbOKOnly + vbCritical, "No Demand Category"
          Exit Sub
      End If
      If cboPayableTypePayDates.text = "" Then
         MsgBox "You must select a payable type payment date.", vbOKOnly + vbCritical, "Payable Type Payment Date"
         Exit Sub
      End If

      adoConn.Open getConnectionString

      If bAddNew Then
         szSQL = "SELECT ID, PayType, PayIC, PaySagePrefix, PayNCAmt, " & _
                     "PayNNAmt, PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, " & _
                     "TransactionType, CategoryCode, PaymentDates " & _
                  "FROM PayableTypes"
      Else
         szSQL = "SELECT ID, PayType, PayIC, PaySagePrefix, PayNCAmt, " & _
                     "PayNNAmt, PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, " & _
                     "TransactionType, CategoryCode, PaymentDates " & _
                  "FROM PayableTypes WHERE ID =" & cboPayChrgType.Column(0) & ""
      End If
      adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      If bAddNew Then
         adoRst.AddNew
         adoRst!Id = cboPayChrgType.ListCount + 1
      End If

      adoRst!PayType = txtPayType.text
      adoRst!PayNCAmt = cboPayNCAmt.text
      adoRst!PayIC = "I"                           'for nothing i have set the value.. but dont use it
      adoRst!PayNNAmt = cboPayNCAmt.Column(1)
      adoRst!PayNCVat = cboPayNCVat.text
      adoRst!PayNNVat = cboPayNCVat.Column(1)
      adoRst!PayNCTotal = cboPayNCTotal.text
      adoRst!PayNNTotal = cboPayNCTotal.Column(1)
      adoRst!CategoryCode = cboPayableTypeCategory.Value
      adoRst!PaymentDates = CByte(cboPayableTypePayDates.Value)
      adoRst.Update
      adoRst.Close
      Set adoRst = Nothing

      LoadPayableType adoConn

      adoConn.Close
      Set adoConn = Nothing

      ShowMsgInTaskBar "Your changes have been saved."
      ButtonMode DefaultMode, Index
   End If
End Sub

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
   adoRst!propertyID = frmDCTypesPre.cboPropertyList.Column(0)
   adoRst!EmailInvoiceTemplate = txtEmailTemplate.text

   adoRst.Update
   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "Your changes have been saved."
   
   Call DisableBoxes
   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   cboDemand.SetFocus
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
   adoRst!Id = txtID.text
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
   adoRst!propertyID = frmDCTypesPre.cboPropertyList.Column(0)
   adoRst!EmailInvoiceTemplate = txtEmailTemplate.text

   adoRst.Update
   
   adoRst.Close
   Set adoRst = Nothing
   
   Call LoadDemandTypes(adoConn)
   
   adoConn.Close
   Set adoConn = Nothing

   cmdAdd.Enabled = True
   cmdEdit.Enabled = True
   cmdDelete.Enabled = True
   cmdSaveNew.Enabled = False
   cmdCancelNew.Enabled = False

   ShowMsgInTaskBar "Your new demand type details have been saved."

   Call DisableBoxes

   cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
   Me.Width = 9435
   Me.Height = 5865
   fraDemandType.Top = 120
   fraDemandType.Left = 120
   fraChargeType.Left = 120
   fraChargeType.Top = 120
   fraPayableType.Top = 120
   fraPayableType.Left = 120
   Me.BackColor = MODULEBACKCOLOR
   fraDemandType.BackColor = MODULEBACKCOLOR
   fraChargeType.BackColor = MODULEBACKCOLOR
   fraPayableType.BackColor = MODULEBACKCOLOR
   
   chkGroup.Enabled = False

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadNCinCombo adoConn         'all nominal code and name are collecting in all combos from sage

   If frmDCTypesPre.szMenu = "DEMAND_TYPE" Then
      LoadDemandTypes adoConn
   End If

   If frmDCTypesPre.szMenu = "CHARGE_TYPE" Then
      LoadChargeType adoConn       'all defined charge/fee types are collecting in the combo from db
   End If

   If frmDCTypesPre.szMenu = "PAYABLE_TYPE" Then
      LoadPayableType adoConn       'all defined payable types are collecting in the combo from db
   End If

   LoadPaymentDates adoConn

   LoadDemandCategory adoConn  'all category

   LoadGroup adoConn

   LoadBankDetails adoConn      'all clients' bank details

   adoConn.Close
   Set adoConn = Nothing
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

Private Sub LoadPayableType(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer
   
   cboPayChrgType.Clear

   szSQL = "SELECT ID, PayType, PayIC, PaySagePrefix, PayNCAmt, " & _
                     "PayNNAmt, PayNCVat, PayNNVat, PayNCTotal, PayNNTotal, " & _
                     "TransactionType, CategoryCode, PaymentDates " & _
           "FROM PayableTypes ORDER BY ID"

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

   cboPayChrgType.Column() = Data()

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
   cboChargeTypeCategory.Clear
   cboPayableTypeCategory.Clear

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
   cboChargeTypeCategory.Column() = Data()
   cboPayableTypeCategory.Column() = Data()
   
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadPaymentDates(adoConn As ADODB.Connection)
   cboDemandTypePayDates.Clear
   cboChargeTypePayDates.Clear
   cboPayableTypePayDates.Clear

   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset
   Dim Data() As String
   Dim TotalRow, TotalCol As Integer

   ReDim Data(1, 0) As String

   Data(0, 0) = "0"
   Data(1, 0) = "DEFAULT"

   cboChargeTypePayDates.Column() = Data()

   szSQL = "SELECT NameOfSet " & _
               "FROM PaymentDates " & _
               "ORDER BY DateSetID;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount = 0 Then
      cboDemandTypePayDates.Column() = Data()
      cboPayableTypePayDates.Column() = Data()
      
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
'   cboChargeTypePayDates.Column() = Data()
   cboPayableTypePayDates.Column() = Data()

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

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   ReDim Data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoRst.RecordCount - 1
       For j = 0 To adoRst.Fields.count - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j)), "", adoRst.Fields(j))
       Next j
       adoRst.MoveNext
   Next i
   cboFeeChrgType.Column() = Data()

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
   cboFeeNCAmt.Column() = Data()
   cboFeeNCVat.Column() = Data()
   cboFeeNCTotal.Column() = Data()
   cboPayNCAmt.Column() = Data()
   cboPayNCVat.Column() = Data()
   cboPayNCTotal.Column() = Data()

   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Public Sub LoadDemandTypes(adoConn As ADODB.Connection)
   Dim Data() As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   cboDemand.Clear

   If frmDCTypesPre.cboPropertyList.Column(0) <> "ALL" Then
      szSQL = "SELECT D.ID, D.Type, D.PropertyID, " & _
                     "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                     "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID " & _
              "FROM (DemandTypes AS D LEFT JOIN Property AS P ON " & _
                    "D.PropertyID = P.PropertyID) LEFT JOIN Client AS C ON P.ClientID = C.ClientID " & _
              "WHERE D.PropertyID = '" & frmDCTypesPre.cboPropertyList.Column(0) & "' OR " & _
                    "D.PropertyID = 'ALL' " & _
              "ORDER BY D.ID;"
   Else
      szSQL = "SELECT D.ID, D.Type, D.PropertyID, " & _
                     "IIF(ISNULL(P.PropertyName), 'All Properties', P.PropertyName) AS PropertyName, " & _
                     "IIF(ISNULL(P.ClientID), 'All Clients', P.ClientID) AS ClientID " & _
              "FROM (DemandTypes AS D LEFT JOIN Property AS P ON " & _
                    "D.PropertyID = P.PropertyID) LEFT JOIN Client AS C ON P.ClientID = C.ClientID " & _
              "ORDER BY D.ID;"
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count

   ReDim Data(TotalCol - 1, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboDemand.Column() = Data()

NoRes:
   adoRst.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmDCTypesPre.Show
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
   cboDemand.Enabled = False
   txtType.Enabled = True
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
   cboDemand.Enabled = True
   txtType.Enabled = False
   cboDemandTypeNCAmt.Enabled = False
   cboDemandTypeNCvat.Enabled = False
   cboDemandTypeNCTotal.Enabled = False
   cboDemandTypeCategory.Enabled = False
   cboBank.Enabled = False

   cmdAddReport.Enabled = False
   cmdAddEmailTemplate.Enabled = False
   chkGroup.Enabled = False
End Sub

Private Sub ButtonMode(mode As ComponentMode, iIndex As Integer)
   Select Case mode
      Case ComponentMode.DefaultMode
         fraMain(iIndex).Enabled = False
         cmdAddNew(iIndex).Enabled = True
         cmdGSEdit(iIndex).Enabled = True
         cmdGSSave(iIndex).Enabled = False
         cmdGSCancel(iIndex).Enabled = False

         If iIndex = 0 Then
            cboFeeChrgType.Locked = False
            txtFeeType.text = ""
            cboFeeNCAmt.text = ""
            cboFeeNCVat.text = ""
            cboFeeNCTotal.text = ""
            cboChargeTypeCategory.text = ""
            txtFeeDemandCategory.text = ""
            cboChargeTypePayDates.text = ""
         End If
         If iIndex = 1 Then
            cboPayChrgType.Locked = False
            txtPayType.text = ""
            cboPayNCAmt.text = ""
            cboPayNCVat.text = ""
            cboPayNCTotal.text = ""
            cboPayableTypeCategory.text = ""
            txtPayCate.text = ""
            cboPayableTypePayDates.text = ""
         End If

      Case ComponentMode.EditMode
         fraMain(iIndex).Enabled = True
         cmdAddNew(iIndex).Enabled = False
         cmdGSEdit(iIndex).Enabled = False
         cmdGSSave(iIndex).Enabled = True
         cmdGSCancel(iIndex).Enabled = True
         If iIndex = 0 Then cboFeeChrgType.Locked = True
         If iIndex = 1 Then cboPayChrgType.Locked = True

      Case ComponentMode.NewEntryMode
         fraMain(iIndex).Enabled = True
         cmdAddNew(iIndex).Enabled = False
         cmdGSEdit(iIndex).Enabled = False
         cmdGSSave(iIndex).Enabled = True
         cmdGSCancel(iIndex).Enabled = True

         If iIndex = 0 Then
            cboFeeChrgType.Locked = True
            cboFeeChrgType.text = ""
            txtFeeType.text = ""
            cboFeeNCAmt.text = ""
            cboFeeNCVat.text = ""
            cboFeeNCTotal.text = ""
            cboChargeTypeCategory.text = ""
            txtFeeDemandCategory.text = ""
            cboChargeTypePayDates.text = ""
         End If
         If iIndex = 1 Then
            cboPayChrgType.Locked = True
            cboPayChrgType.text = ""
            txtPayType.text = ""
            cboPayNCAmt.text = ""
            cboPayNCVat.text = ""
            cboPayNCTotal.text = ""
            cboPayableTypeCategory.text = ""
            txtPayCate.text = ""
            cboPayableTypePayDates.text = ""
         End If

      Case ComponentMode.SavedMode
         fraMain(iIndex).Enabled = False
         cmdAddNew(iIndex).Enabled = True
         cmdGSEdit(iIndex).Enabled = True
         cmdGSSave(iIndex).Enabled = False
         cmdGSCancel(iIndex).Enabled = False

         If iIndex = 0 Then cboFeeChrgType.Locked = False
         If iIndex = 1 Then cboPayChrgType.Locked = False
   End Select
End Sub

Private Sub TextBox1_Change()

End Sub
