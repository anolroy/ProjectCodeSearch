VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMaintananceDairy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance - Dairy Entry Detail"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   Icon            =   "frmMaintananceDairy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDmdLeaseList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   43
      Top             =   4590
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   48
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
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
            Index           =   6
            Left            =   3000
            TabIndex        =   52
            Top             =   0
            Width           =   645
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
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
            Index           =   5
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   465
         End
         Begin MSForms.ComboBox ComboBox2 
            Height          =   315
            Left            =   3675
            TabIndex        =   50
            Top             =   0
            Width           =   2295
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4048;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411"
         End
         Begin MSForms.ComboBox ComboBox1 
            Height          =   315
            Left            =   480
            TabIndex        =   49
            Top             =   0
            Width           =   2415
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "4260;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   8
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411"
         End
      End
      Begin VB.TextBox txtDmdTenantSearchUnitName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4080
         TabIndex        =   47
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox txtDmdTenantSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   46
         Top             =   300
         Width           =   2415
      End
      Begin VB.TextBox txtDmdTenantSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdDmdGridUnitLookup 
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
         Left            =   6080
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   20
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   53
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4392
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   59
         Top             =   75
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   58
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   57
         Top             =   75
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   17
         Left            =   45
         Top             =   30
         Width           =   6015
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   56
         Top             =   45
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   55
         Top             =   45
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name"
         Height          =   195
         Index           =   2
         Left            =   4095
         TabIndex        =   54
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.Frame fraReportedBy 
      Caption         =   "Reported By:"
      Height          =   735
      Left            =   90
      TabIndex        =   38
      Top             =   1530
      Width           =   4515
      Begin VB.CommandButton cmdTask_Assigned 
         Caption         =   "..."
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
         Index           =   2
         Left            =   4080
         TabIndex        =   41
         Top             =   280
         Width           =   330
      End
      Begin VB.OptionButton optInternal_Reported 
         Caption         =   "Internal"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   285
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optLessee 
         Caption         =   "Lessee"
         Height          =   255
         Left            =   1065
         TabIndex        =   39
         Top             =   285
         Width           =   855
      End
      Begin MSForms.ComboBox cboReportedBy 
         Height          =   315
         Left            =   1935
         TabIndex        =   11
         Top             =   285
         Width           =   2130
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3757;556"
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0"
      End
      Begin MSForms.TextBox txtTenantName 
         Height          =   285
         Left            =   1935
         TabIndex        =   42
         Top             =   285
         Visible         =   0   'False
         Width           =   2130
         VariousPropertyBits=   679495709
         BackColor       =   16777215
         BorderStyle     =   1
         Size            =   "3757;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdTask_Assigned 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtDateCompleted 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   18
      Top             =   2325
      Width           =   3255
   End
   Begin VB.TextBox txtNextRemDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtExpCompletionDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1965
      Width           =   1695
   End
   Begin VB.TextBox txtExpStartDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtDateReported 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   1605
      Width           =   3255
   End
   Begin VB.TextBox txtJobDetail 
      Appearance      =   0  'Flat
      Height          =   1305
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   3405
      Width           =   8295
   End
   Begin VB.CommandButton cmdTask_Assigned 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   9240
      TabIndex        =   22
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdType 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   8640
      TabIndex        =   37
      Top             =   1965
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "(dd/mm/yyyy)"
      Size            =   "1720;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label13 
      Height          =   255
      Left            =   1320
      TabIndex        =   36
      Top             =   3000
      Width           =   2415
      VariousPropertyBits=   8388627
      Caption         =   "(dd/mm/yyyy               hh:mm - 24 Hrs)"
      Size            =   "4260;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   465
   End
   Begin MSForms.ComboBox cboPropertyList 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   2
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1234;2822"
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   34
      Top             =   495
      Width           =   645
   End
   Begin MSForms.ComboBox cboClientList 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5741;556"
      TextColumn      =   2
      ColumnCount     =   8
      ListRows        =   20
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtLocation 
      Height          =   315
      Left            =   6360
      TabIndex        =   7
      Top             =   840
      Width           =   3255
      VariousPropertyBits=   746604571
      MaxLength       =   40
      BorderStyle     =   1
      Size            =   "5741;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Width           =   1095
      VariousPropertyBits=   8388627
      Caption         =   "Assigned To"
      Size            =   "1931;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboAssignedTo 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   2880
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5089;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblJobName 
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   480
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Subject"
      Size            =   "1508;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtJobName 
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   480
      Width           =   3255
      VariousPropertyBits=   746604571
      MaxLength       =   40
      BorderStyle     =   1
      Size            =   "5741;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboType 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   2880
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5089;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtNextRemTime 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   2640
      Width           =   855
      VariousPropertyBits=   746604571
      MaxLength       =   6
      BorderStyle     =   1
      Size            =   "1508;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtRef 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Top             =   115
      Width           =   3255
      VariousPropertyBits=   746604569
      MaxLength       =   9
      BorderStyle     =   1
      Size            =   "5741;556"
      SpecialEffect   =   0
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox cbAlarm 
      Height          =   285
      Left            =   3720
      TabIndex        =   15
      Top             =   2760
      Width           =   855
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1508;503"
      Value           =   "0"
      Caption         =   "Alarm"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboTaskOwner 
      Height          =   315
      Left            =   6360
      TabIndex        =   10
      Top             =   1200
      Width           =   2880
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5089;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8160
      TabIndex        =   21
      Top             =   4860
      Width           =   1455
      Caption         =   "Cancel"
      Size            =   "2566;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   4860
      Width           =   1455
      Caption         =   "Save"
      Size            =   "2566;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   2325
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Date Completed"
      Size            =   "2566;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label12 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Reminder"
      Size            =   "1508;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label11 
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   840
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Location"
      Size            =   "1508;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label9 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "1720;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   1965
      Width           =   1935
      VariousPropertyBits=   8388627
      Caption         =   "Expected Completion"
      Size            =   "3413;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   975
      VariousPropertyBits=   8388627
      Caption         =   "Start Date"
      Size            =   "1720;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   1605
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Date Entered"
      Size            =   "2355;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   1200
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "Task Owner"
      Size            =   "2143;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   735
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "1296;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Left            =   4815
      TabIndex        =   2
      Top             =   180
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Diary Entry No:"
      Size            =   "2355;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmMaintananceDairy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Created_Ref As String
Public isEdit As Boolean
Public UpdateRow As Integer
Public RecordType As String
Public CallingForm As String

Private Sub cboClientList_Change()
   If cboClientList.ListIndex < 0 Then Exit Sub

   Dim adoConn    As New ADODB.Connection

   adoConn.Open getConnectionString

   LoadProperty adoConn, cboPropertyList

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadProperty(adoConn As ADODB.Connection, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & cboClientList.Value & "' " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   ReDim Data(TotalCol, TotalRow - 1) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboP.Clear
   cboP.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub cmdCancel_Click()
   txtRef.text = ""
   txtJobName.text = ""
   cboType.text = ""
   cboTaskOwner.text = ""
   cboAssignedTo.text = ""
   txtDateReported.text = ""
   txtExpStartDate.text = ""
   txtExpCompletionDate.text = ""
   txtJobDetail.text = ""
   txtNextRemDate.text = ""
   txtNextRemTime.text = ""
   cbAlarm.Value = False
   txtDateCompleted.text = ""
   
   Unload Me
End Sub


Private Sub cmdDmdGridUnitLookup_Click()
   picDmdLeaseList.Visible = False
End Sub

Private Sub cmdSave_Click()
   Dim adoConn As New ADODB.Connection

   If (Not validateMandatory(cboClientList, "Client Name cannot be empty")) Then Exit Sub
   If (Not validateMandatory(cboPropertyList, "Property Name cannot be empty")) Then Exit Sub
   If (Not validateMandatory(txtJobName, "Dirary Name cannot be empty")) Then Exit Sub
   If (Not validateMandatory(cboType, "Type cannot be empty")) Then Exit Sub
   If (Not validateMandatory(cboTaskOwner, "Task Owner cannot be empty")) Then Exit Sub
   If (Not validateMandatory(cboAssignedTo, "Assigned To cannot be empty")) Then Exit Sub
   If (Not validateMandatory(txtDateReported, "Date Reported cannot be empty")) Then Exit Sub

   adoConn.Open getConnectionString

   If SavePropertyMaintenanceHistory(adoConn) Then
      If CallingForm = "M" Then
         ShowMsgInTaskBar "The Property maintenance event saved successfully."
         frmMaintenance.RefreshMaintenanceGrid adoConn
      End If
      If CallingForm = "P" Then
         ShowMsgInTaskBar "The Property maintenance event saved successfully."
         frmProperty2.LoadGridMaintenanceHistory adoConn
      End If
      If CallingForm = "U" Then
         ShowMsgInTaskBar "The Unit maintenance event saved successfully."
         frmUnits2.LoadGridMaintenanceHistory adoConn
      End If
   Else
       ShowMsgInTaskBar "Could not save maintenance event", , "N"
   End If

   adoConn.Close
   Set adoConn = Nothing

   Unload Me
End Sub

Private Sub cmdTask_Assigned_Click(Index As Integer)
'issue 474
'added by anol 26 Nov 2014
   If Index = 2 And optLessee.Value = True Then
         Dim szSQL As String
         Dim adoConn As New ADODB.Connection
         Dim adoRst As New ADODB.Recordset
         If cboPropertyList.Value = "" Then
         szSQL = "SELECT T.SageAccountNumber, T.Name, L.UnitNumber " & _
                  "FROM (Tenants AS T INNER JOIN LeaseDetails AS L ON " & _
                      "T.SageAccountNumber = L.SageAccountNumber) INNER JOIN Units AS U ON " & _
                      "L.UnitNumber = U.UnitNumber " & _
                  "WHERE L.Status " & _
                  "ORDER BY T.Name;"
         Else
            szSQL = "SELECT T.SageAccountNumber, T.Name, L.UnitNumber " & _
                  "FROM (Tenants AS T INNER JOIN LeaseDetails AS L ON " & _
                      "T.SageAccountNumber = L.SageAccountNumber) INNER JOIN Units AS U ON " & _
                      "L.UnitNumber = U.UnitNumber " & _
                  "WHERE L.Status AND U.PropertyID = '" & cboPropertyList.Value & "' " & _
                  "ORDER BY T.Name;"
         End If
         adoConn.Open getConnectionString
         adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim szHeader As String
   flxDmdLeaseList.Clear
   flxDmdLeaseList.Cols = 4
   flxDmdLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Lessee ID|<Lessee Name|<Unit Name|<ExpLes"
   flxDmdLeaseList.FormatString = szHeader$
   flxDmdLeaseList.ColWidth(0) = Label20(9).Left - flxDmdLeaseList.Left   '240        Solid column
   flxDmdLeaseList.ColWidth(1) = Label20(8).Left - Label20(9).Left - 20   '1400       'Tenant ID
   flxDmdLeaseList.ColWidth(2) = Label20(7).Left - Label20(8).Left - 20               'Tenant Name
   flxDmdLeaseList.ColWidth(3) = flxDmdLeaseList.Left + flxDmdLeaseList.Width - Label20(7).Left - 300 'Unit Name
   flxDmdLeaseList.Rows = 2


   Dim iRow As Integer
   iRow = 1
      While Not adoRst.EOF
         flxDmdLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
         flxDmdLeaseList.TextMatrix(iRow, 2) = adoRst!Name
         flxDmdLeaseList.TextMatrix(iRow, 3) = adoRst!UnitNumber
         iRow = iRow + 1
         adoRst.MoveNext
         If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
      Wend
   
      picDmdLeaseList.Top = fraReportedBy.Top + 185
      picDmdLeaseList.Left = txtTenantName.Left + 105
      picDmdLeaseList.Visible = True
      adoRst.Close
      Set adoRst = Nothing
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If
   'end of addition by anol
   Dim sSQLQuery As String
   frmSecondaryCode.PRIMARY_CODE_SHOW = "MNTJOB"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
      populateCombo adoConn, sSQLQuery, cboReportedBy
   adoConn.Close
   Set adoConn = Nothing
  
'
'
'
'
'
'   Dim sSQLQuery As String
'  ' Dim adoConn As New ADODB.Connection
'   Dim SelTaskOwner, SelAssignedTo As String
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "MNTJOB"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoConn.Open getConnectionString
'   sSQLQuery = "SELECT CODE, VALUE " & _
'               "FROM SECONDARYCODE " & _
'               "WHERE PRIMARYCODE = 'MNTJOB'"
'   SelTaskOwner = IIf(cboTaskOwner.text = "", "", cboTaskOwner.Value)
'   SelAssignedTo = IIf(cboAssignedTo.text = "", "", cboAssignedTo.Value)
'   populateCombo adoConn, sSQLQuery, cboTaskOwner
'   populateCombo adoConn, sSQLQuery, cboAssignedTo
'   cboTaskOwner.text = SelTaskOwner
'   cboAssignedTo.text = SelAssignedTo
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

Private Sub cmdType_Click()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   Dim selType As String

   selType = IIf(cboType.text = "", "", cboType.text)
   frmSecondaryCode.PRIMARY_CODE_SHOW = "MTYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1
   
   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MTYP'"
   populateCombo adoConn, sSQLQuery, cboType
   cboType.text = selType

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub flxDmdLeaseList_Click()
'Resolved by BOSL
'issue 474
'added by anol 26 Nov 2014
   txtTenantName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1)
   picDmdLeaseList.Visible = False
   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchName.text = ""
   txtDmdTenantSearchUnitName.text = ""
End Sub

Private Sub Form_Activate()
    txtRef.Enabled = False
End Sub

Private Sub Form_Load()
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   fraReportedBy.BackColor = MODULEBACKCOLOR
   cmdTask_Assigned(2).BackColor = MODULEBACKCOLOR
   optInternal_Reported.BackColor = MODULEBACKCOLOR
   optLessee.BackColor = MODULEBACKCOLOR
   Me.Height = 5850
   Me.Width = 9780
  
   LoadValues
   If (isEdit) Then
      loadEditValues
   Else
      txtRef.text = getNextRef
   End If
End Sub

Private Sub SupplierAccCombo()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String, Data() As String, i As Integer

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString
   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier " & _
           "ORDER BY SupplierName;"

   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRst.EOF Then GoTo NoRes

   ReDim Data(1, rstRst.RecordCount) As String
   cboAssignedTo.Clear

   For i = 0 To rstRst.RecordCount - 1
      Data(0, i) = CStr(rstRst!SupplierID)
      Data(1, i) = CStr(rstRst!SupplierName)
      rstRst.MoveNext
   Next i
   cboAssignedTo.Column() = Data()
   cboAssignedTo.BoundColumn = 1

NoRes:
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

Public Function getNextRef()
   Dim szSQL  As String, RefId As String, NewId As String
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim prefix As String
   
   If (RecordType = "J") Then
      prefix = "JOB"
      szSQL = "SELECT MAX(ID) AS RefFound From PropertyMaintHistory Where ID Like 'JOB%'"
   Else
      prefix = "DIA"
      szSQL = "SELECT MAX(ID) AS RefFound From PropertyMaintHistory Where ID Like 'DIA%'"
   End If
      
   adoConn.Open getConnectionString
   'find the largest available ID
   
   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   RefId = IIf(IsNull(rstRst!RefFound), IIf(RecordType = "J", "JOB000000", "DIA000000"), rstRst!RefFound)
   rstRst.Close

   NewId = CStr(CInt(Right(RefId, 5)) + 1)
   Dim i As Integer
     
   For i = Len(NewId) To 5
      NewId = "0" + NewId
   Next i
      
   Created_Ref = prefix + NewId
   getNextRef = Created_Ref
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If CallingForm = "M" Then _
      frmMaintenance.Enabled = True
   If CallingForm = "P" Then _
      frmProperty2.Enabled = True
   If CallingForm = "U" Then _
      frmUnits2.Enabled = True
End Sub

Private Sub optSupplier_Click()
   SupplierAccCombo
   cmdTask_Assigned(1).Visible = False
End Sub

Private Sub optInternal_Reported_Click()
'Resolved by BOSL
'added by anol 30 Nov 2014
'issue 474
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
           
   adoConn.Open getConnectionString
   'AssignedTo
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   populateCombo adoConn, sSQLQuery, cboReportedBy
   
   adoConn.Close
   Set adoConn = Nothing
   cmdTask_Assigned(2).Visible = True
   cboReportedBy.Visible = True
   txtTenantName.Visible = False
   picDmdLeaseList.Visible = False
End Sub

Private Sub optLessee_Click()
'Resolved by BOSL
'added by anol 30 Nov 2014
'issue 474
   cboReportedBy.Visible = False
   txtTenantName.Visible = True
   If txtTenantName.text = "" Then
      Call cmdTask_Assigned_Click(2)
   End If
End Sub
Private Sub txtDmdTenantSearchID_Change()
 'issue 474
   'added by anol 26 Nov 2014

   Dim i As Integer

   If Len(txtDmdTenantSearchID.text) > 0 Then
      txtDmdTenantSearchName.text = ""
      txtDmdTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 1), Len(txtDmdTenantSearchID.text))) <> UCase(txtDmdTenantSearchID.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub
Private Sub txtDmdTenantSearchName_Change()
 'issue 474
   'added by anol 26 Nov 2014

   Dim i As Integer

   If Len(txtDmdTenantSearchName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchUnitName.text = ""
   End If

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 2), Len(txtDmdTenantSearchName.text))) <> UCase(txtDmdTenantSearchName.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub
Private Sub txtDmdTenantSearchUnitName_Change()
   'issue 474
   'added by anol 26 Nov 2014
   Dim i As Integer

   If Len(txtDmdTenantSearchUnitName.text) > 0 Then
      txtDmdTenantSearchID.text = ""
      txtDmdTenantSearchName.text = ""
   End If

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 3), Len(txtDmdTenantSearchUnitName.text))) <> UCase(txtDmdTenantSearchUnitName.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub
Private Sub txtDateCompleted_Change()
   TextBoxChangeDate txtDateCompleted
End Sub

Private Sub txtDateCompleted_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateCompleted, KeyAscii
End Sub

Private Sub txtDateCompleted_LostFocus()
   TextBoxFormatDate txtDateCompleted
End Sub

Private Sub txtDateReported_Change()
   TextBoxChangeDate txtDateReported
End Sub

Private Sub txtDateReported_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateReported, KeyAscii
End Sub

Private Sub txtDateReported_LostFocus()
   TextBoxFormatDate txtDateReported
End Sub

Private Sub txtExpCompletionDate_Change()
   TextBoxChangeDate txtExpCompletionDate
End Sub

Private Sub txtExpCompletionDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtExpCompletionDate, KeyAscii
End Sub

Private Sub txtExpCompletionDate_LostFocus()
   TextBoxFormatDate txtExpCompletionDate
End Sub

Private Sub txtExpStartDate_Change()
   TextBoxChangeDate txtExpStartDate
End Sub

Private Sub txtExpStartDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtExpStartDate, KeyAscii
End Sub

Private Sub txtExpStartDate_LostFocus()
   TextBoxFormatDate txtExpStartDate
End Sub

Private Sub txtNextRemDate_Change()
   TextBoxChangeDate txtNextRemDate
End Sub

Private Sub txtNextRemDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtNextRemDate, KeyAscii
End Sub

Private Sub txtNextRemDate_LostFocus()
   TextBoxFormatDate txtNextRemDate
End Sub

Private Sub txtNextRemTime_Change()
   prsTime
End Sub

Private Sub txtNextRemTime_LostFocus()
   validateTime
   If txtNextRemTime.text = "" Then
         cbAlarm.Value = vbUnchecked
   End If
End Sub

Private Sub txtRef_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtRef_LostFocus()
   Dim szSQL  As String
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   
   adoConn.Open getConnectionString
   
   If (validateMandatory(txtRef, "Ref can not be empty")) Then
      If (txtRef.text <> Created_Ref) Then
         If (Left(txtRef.text, 3) = "JOB" Or Left(txtRef.text, 3) = "DIA") Then
            ShowMsgInTaskBar "User defined IDs cannot start with JOB or DIA", , "N"
         Else
            'Check whether user entered Ref allready exists
            szSQL = "SELECT ID From PropertyMaintHistory Where ID = '" + txtRef.text + "'"
            rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If (rstRst.RecordCount > 0) Then
               MsgBox "Ref Already Exist, Please change", , "N"
            End If
            rstRst.Close
            adoConn.Close
         End If
      End If
   End If
End Sub

Private Sub LoadValues()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
   
   adoConn.Open getConnectionString
   'reported by
   'added by anol 30 Nov 2014
   'issue 474
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   populateCombo adoConn, sSQLQuery, cboReportedBy
   'end of addition
   
   'Maintenance Type
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MTYP'"

   populateCombo adoConn, sSQLQuery, cboType

   'TaskOwner and AssignedTo
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   populateCombo adoConn, sSQLQuery, cboTaskOwner
   populateCombo adoConn, sSQLQuery, cboAssignedTo
   LoadCmbClient adoConn, cboClientList

   txtDateReported.text = Date
   SelTxtInCtrl txtDateReported

   txtExpStartDate.text = Date
   SelTxtInCtrl txtExpStartDate
End Sub

Private Sub LoadCmbClient(adoConn As ADODB.Connection, cboC As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Function validateMandatory(testCon As Control, Msg As String)
   If (testCon.text = "") Then
      MsgBox Msg, vbInformation, "Mandatory Field"
      testCon.SetFocus
      validateMandatory = False
      Exit Function
   End If
   validateMandatory = True
End Function

Private Function validateTime()
   Dim time, szaTempTime() As String

   If (Not txtNextRemTime.text = "") Then
      szaTempTime = Split(txtNextRemTime.text, ":")
      If (UBound(szaTempTime) = 1) Then
         If (Not (Val(szaTempTime(0)) < 24 And Val(szaTempTime(1)) < 60 And Val(szaTempTime(0)) >= 0 And Val(szaTempTime(1)) >= 0)) Then
            ShowMsgInTaskBar "Time is not valid, please follow correct format HH:MM (24 Hr)", , "N"
            'Modified by anol
            'issue 487 note 6
            '15 Oct 2014
            
         Else
            cbAlarm.Value = vbChecked
            'end of modiciation
         End If
      End If
   End If
End Function

Private Sub prsTime()
   If Len(txtNextRemTime.text) = 2 Then
      txtNextRemTime.text = txtNextRemTime + ":"
      txtNextRemTime.SelStart = Len(txtNextRemTime.text)
   End If
End Sub

Private Sub loadEditValues()
   txtRef.Enabled = False

   If CallingForm = "P" Then
      RecordType = Left(frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 0), 1)
      txtRef.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 3)
      txtJobName.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 4)
      cboType.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 1)
      txtDateReported.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 2)
      cboTaskOwner.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 5)
      cboAssignedTo.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 6)
      txtNextRemDate.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 7)
      txtDateCompleted.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 9)
      cbAlarm.Value = IIf(frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 8) = "Y", True, False)
      txtExpStartDate.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 11)
      txtExpCompletionDate.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 12)
      txtNextRemTime.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 15)
      txtLocation.text = frmProperty2.gridMaintenanceHistory.TextMatrix(UpdateRow, 10)
   End If

   If CallingForm = "U" Then
      RecordType = Left(frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 0), 1)
      txtRef.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 3)
      txtJobName.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 4)
      cboType.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 1)
      txtDateReported.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 2)
      cboTaskOwner.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 5)
      cboAssignedTo.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 6)
      txtNextRemDate.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 7)
      txtDateCompleted.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 9)
      cbAlarm.Value = IIf(frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 8) = "Y", True, False)
      txtExpStartDate.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 11)
      txtExpCompletionDate.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 12)
      txtNextRemTime.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 15)
      txtLocation.text = frmUnits2.gridMaintenanceHistory.TextMatrix(UpdateRow, 10)
   End If
'Resolved By BOSL
'issue 474 Note 2
'Modified by anol 28 Sep 2014
   If CallingForm = "M" Then
      Dim adoConn As New ADODB.Connection
      If adoConn.State = 0 Then
         adoConn.Open getConnectionString
      End If
            'Modified by anol
            'issue 487 note 5
            '15 Oct 2014
      Dim rsCheck As New ADODB.Recordset
      rsCheck.Open "SELECT H.location,H.Remindtime,H.Alarm from PropertyMaintHistory As H where H.PropertyID & '-' & H.ID ='" & frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 3) & "'", adoConn, adOpenStatic, adLockReadOnly
      If rsCheck.EOF = False Then
         txtNextRemTime.text = IIf(IsNull(rsCheck("Remindtime")) = True, "", rsCheck("Remindtime").Value)
         cbAlarm.Value = IIf(IsNull(rsCheck("Alarm")) = True, "", rsCheck("Alarm").Value)
         txtLocation.text = IIf(IsNull(rsCheck("location")) = True, "", rsCheck("location").Value)
      End If
      rsCheck.Close
      Set rsCheck = Nothing
      If adoConn.State = 1 Then
         adoConn.Close
      End If
      'end of modification
      
      '   SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
'                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName(4), H.TaskOwner(5),H.ReportedBy(6), " & _
'                "H.AssignedTo(7), H.RemindDate(8), IIF(H.Alarm, 'YES', 'NO')(9), H.DateCompleted(10), " & _
'                "H.BudgetCost(11), H.ExpectedStartDate(12), H.ExpectedCompletionDate(13), " & _
'                "H.Detail(14), H.ActualCost(15),  H.AssignedILz(16), " & _
'                "H.ReportedIS(17), H.RemindTime(18), H.Urgent(19), H.MaintenanceType(20), " & _
'                "H.ReportedFrom(21), H.FundID(22), H.OverrideBudget(23), H.FYrID(24), " & _
'                "H.BudgetPassed(25), P.PropertyID(26), P.ClientID(27), " & _
'                "IIf(AssignedIL='S',U.SupplierOfficeEmail,S1.Description) AS EmailAdd (28)" & _

      cboClientList.Value = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 27)
      cboPropertyList.Value = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 26)
      RecordType = Left(frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 0), 1)
      txtRef.text = Right(frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 3), 9)
      txtJobName.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 4)
      cboType.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 1)
      txtDateReported.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 2)
      cboTaskOwner.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 5)
      'added by anol 30 Nov 2014
      'issue 474
      'Reported by additon in Dairy
      With frmMaintenance.flxMaintenance
      If frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
       End With
        'End of addition
      cboAssignedTo.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 7)
      txtNextRemDate.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 8)
      txtDateCompleted.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 10)
      'cbAlarm.Value = IIf(frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 8) = "Y", True, False)
      txtExpStartDate.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 12)
      txtExpCompletionDate.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 13)
      'txtNextRemTime.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 15)
     ' txtLocation.text = frmMaintenance.flxMaintenance.TextMatrix(UpdateRow, 10)
   End If
'End of modification

   RetrieveMemo "PropertyMaintHistory", "Detail", txtRef.text, "ID", txtJobDetail

   SelTxtInCtrl txtDateCompleted
   SelTxtInCtrl txtNextRemDate
   SelTxtInCtrl txtNextRemTime
   SelTxtInCtrl txtExpCompletionDate
End Sub

Public Function SavePropertyMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection) As Boolean
   Dim rstMHistory_ As New ADODB.Recordset
   Dim rstID As New ADODB.Recordset
   Dim sSQLQuery_ As String, sSQLDelete As String, sSQLFilter As String, iRowIndex As Integer
   Dim lTableID As Long, szAlarmTime As String

   sSQLFilter = ""

'   On Error GoTo Exception

   If isEdit Then sSQLFilter = "WHERE PropertyID = '" & cboPropertyList.Value & "' AND ID = '" & txtRef.text & "'"
'      If CallingForm = "P" Then _
'         sSQLFilter = "WHERE PropertyID = '" & frmProperty2.txtPropertyID.text & "' AND ID = '" & txtRef.text & "'"
'      If CallingForm = "U" Then _
'         sSQLFilter = "WHERE PropertyID = '" & frmUnits2.txtUnitNo.text & "' AND ID = '" & txtRef.text & "'"
'   Else
'      sSQLFilter = ""
'   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyMAINTHISTORY " & sSQLFilter

   rstMHistory_.Open sSQLQuery_, conMHistory_, adOpenDynamic, adLockOptimistic

   If Not isEdit Then rstMHistory_.AddNew

   rstMHistory_!ID = txtRef.text
   rstMHistory_!propertyID = cboPropertyList.Value
'   If CallingForm = "P" Then _
'      rstMHistory_!PropertyID = frmProperty2.txtPropertyID.text
'   If CallingForm = "U" Then _
'      rstMHistory_!PropertyID = frmUnits2.txtUnitNo.text
   rstMHistory_!Job_DiaryName = txtJobName.text
   rstMHistory_!MaintenanceType = cboType.Value
   rstMHistory_!TaskOwner = IIf(cboTaskOwner.text = "", "", cboTaskOwner.text)
   rstMHistory_!AssignedTo = IIf(cboAssignedTo.text = "", "", cboAssignedTo.text)
   
   'added by anol 30 nov 2014
   'issue 474
   'Reported by addition in diary
   If optInternal_Reported.Value Then
      rstMHistory_!ReportedIS = "I"
      rstMHistory_!ReportedBy = cboReportedBy.Value
   Else
      rstMHistory_!ReportedIS = "L"
      rstMHistory_!ReportedBy = txtTenantName.text
   End If
   'End of addition
   
   rstMHistory_!ReportedDate = Format(IIf(txtDateReported.text = "", Now, txtDateReported.text), "DD/MM/YYYY")
   rstMHistory_!ExpectedStartDate = IIf(txtExpStartDate.text = "", "", Format(txtExpStartDate.text, "DD/MM/YYYY"))
   rstMHistory_!ExpectedCompletionDate = IIf(txtExpCompletionDate.text = "", "", Format(txtExpCompletionDate.text, "DD/MM/YYYY"))
   rstMHistory_!Detail = IIf(txtJobDetail.text = "", "", txtJobDetail.text)
   rstMHistory_!RemindTime = IIf(txtNextRemTime.text = "", "", txtNextRemTime.text)
   rstMHistory_!Location = txtLocation.text
   If (txtDateCompleted.text = "") Then
     rstMHistory_!DateCompleted = Null
   Else
     rstMHistory_!DateCompleted = CDate(Format(txtDateCompleted.text, "DD/MM/YYYY"))
   End If

   If (txtNextRemDate.text = "") Then
     rstMHistory_!RemindDate = Null
   Else
     rstMHistory_!RemindDate = CDate(Format(txtNextRemDate.text, "DD/MM/YYYY"))
   End If

   rstMHistory_!RecordType = RecordType
   rstMHistory_!ReportedFrom = CallingForm

   If cbAlarm.Value Then
      rstMHistory_!Alarm = True
      szAlarmTime = IIf(txtNextRemTime.text = "", "083000", Format(txtNextRemTime.text, "hhmm") & "00")

      If Not isEdit Then
         rstMHistory_!Reminder_ID = NewReminder(Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), szAlarmTime, txtJobName.text, "PropertyMAINTHISTORY", txtRef.text)
      Else
         UpdateReminder rstMHistory_!Reminder_ID, Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), szAlarmTime, txtJobName.text
      End If
   Else
      rstMHistory_!Alarm = False
   End If

   rstMHistory_.Update

   rstMHistory_.Close
   Set rstMHistory_ = Nothing

   SavePropertyMaintenanceHistory = True
   Exit Function

Exception:
   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
   rstMHistory_.Close

   Set rstMHistory_ = Nothing

   SavePropertyMaintenanceHistory = False
End Function
