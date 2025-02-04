VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMaintenanceJob 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maintenance - Job Entry"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15870
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaintenanceJob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15870
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   11565
      ScaleHeight     =   4425
      ScaleWidth      =   5580
      TabIndex        =   98
      Top             =   2025
      Visible         =   0   'False
      Width           =   5610
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3750
         Left            =   45
         TabIndex        =   100
         Top             =   675
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   6615
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   106
         Top             =   375
         Width           =   3915
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6906;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   105
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
         Left            =   1620
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   101
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0FFFF&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   15
         Left            =   45
         Top             =   75
         Width           =   5220
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   45
      TabIndex        =   32
      Top             =   45
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Job Entry"
      TabPicture(0)   =   "frmMaintenanceJob.frx":1202
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Memo/Attachment"
      TabPicture(1)   =   "frmMaintenanceJob.frx":121E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6405
         Left            =   90
         TabIndex        =   33
         Top             =   360
         Width           =   11490
         Begin VB.CommandButton cmdFund 
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
            Left            =   11055
            TabIndex        =   20
            Top             =   2385
            Width           =   300
         End
         Begin VB.CommandButton cmdBudgetYears 
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
            Left            =   11070
            TabIndex        =   14
            Top             =   2070
            Width           =   300
         End
         Begin VB.PictureBox picDmdLeaseList 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   3135
            Left            =   240
            ScaleHeight     =   3105
            ScaleWidth      =   6345
            TabIndex        =   36
            Top             =   5730
            Visible         =   0   'False
            Width           =   6375
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
               TabIndex        =   45
               Top             =   20
               Width           =   255
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
               TabIndex        =   44
               Top             =   300
               Width           =   1425
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
               TabIndex        =   43
               Top             =   300
               Width           =   2505
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
               TabIndex        =   42
               Top             =   300
               Width           =   1935
            End
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
               TabIndex        =   37
               Top             =   3240
               Visible         =   0   'False
               Width           =   6015
               Begin MSForms.ComboBox ComboBox1 
                  Height          =   315
                  Left            =   480
                  TabIndex        =   41
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
               Begin MSForms.ComboBox ComboBox2 
                  Height          =   315
                  Left            =   3675
                  TabIndex        =   40
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
                  TabIndex        =   39
                  Top             =   0
                  Width           =   465
               End
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
                  TabIndex        =   38
                  Top             =   0
                  Width           =   645
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
               Height          =   2490
               Left            =   45
               TabIndex        =   46
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
               Caption         =   "Unit Name"
               Height          =   195
               Index           =   2
               Left            =   4095
               TabIndex        =   52
               Top             =   45
               Width           =   735
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Index           =   1
               Left            =   1575
               TabIndex        =   51
               Top             =   45
               Width           =   405
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   50
               Top             =   45
               Width           =   165
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
               Caption         =   "Unit Name"
               Height          =   195
               Index           =   7
               Left            =   4080
               TabIndex        =   49
               Top             =   75
               Width           =   735
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               Height          =   195
               Index           =   8
               Left            =   1560
               TabIndex        =   48
               Top             =   75
               Width           =   405
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   47
               Top             =   75
               Width           =   165
            End
         End
         Begin VB.TextBox txtInsruction 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   1725
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   5370
            Width           =   9645
         End
         Begin VB.Frame fraReportedBy 
            Caption         =   "Reported By:"
            Height          =   735
            Left            =   45
            TabIndex        =   35
            Top             =   2070
            Width           =   5550
            Begin VB.OptionButton optLessee 
               Caption         =   "Lessee"
               Height          =   255
               Left            =   1065
               TabIndex        =   16
               Top             =   285
               Width           =   855
            End
            Begin VB.OptionButton optInternal_Reported 
               Caption         =   "Internal"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   285
               Value           =   -1  'True
               Width           =   855
            End
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
               Left            =   5160
               TabIndex        =   19
               Top             =   280
               Width           =   285
            End
            Begin MSForms.TextBox txtTenantName 
               Height          =   285
               Left            =   1935
               TabIndex        =   17
               Top             =   285
               Visible         =   0   'False
               Width           =   3210
               VariousPropertyBits=   679495709
               BackColor       =   16777215
               BorderStyle     =   1
               Size            =   "5662;503"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontEffects     =   1073750016
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cboReportedBy 
               Height          =   315
               Left            =   1935
               TabIndex        =   18
               Top             =   285
               Width           =   3210
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5662;556"
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
         End
         Begin VB.TextBox txtDateCompleted 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            MaxLength       =   10
            TabIndex        =   23
            Top             =   3105
            Width           =   3525
         End
         Begin VB.TextBox txtNextRemDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1725
            MaxLength       =   10
            TabIndex        =   26
            Top             =   3705
            Width           =   1095
         End
         Begin VB.TextBox txtExpCompletionDate 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7845
            MaxLength       =   10
            TabIndex        =   13
            Top             =   1725
            Width           =   3525
         End
         Begin VB.TextBox txtExpStartDate 
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
            Height          =   315
            Left            =   7845
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1380
            Width           =   3525
         End
         Begin VB.TextBox txtDateReported 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1725
            TabIndex        =   22
            Top             =   2865
            Width           =   3855
         End
         Begin VB.TextBox txtJobDetail 
            Appearance      =   0  'Flat
            Height          =   990
            Left            =   1725
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   4290
            Width           =   9645
         End
         Begin VB.Frame fraAssignedTo 
            Caption         =   "Assigned To:"
            Height          =   735
            Left            =   45
            TabIndex        =   34
            Top             =   1305
            Width           =   5535
            Begin VB.OptionButton optSupplier 
               Caption         =   "Supplier"
               Height          =   375
               Left            =   1080
               TabIndex        =   9
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optInternal 
               Caption         =   "Internal"
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
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
               Index           =   1
               Left            =   5130
               TabIndex        =   12
               Top             =   280
               Width           =   285
            End
            Begin MSForms.ComboBox cboAssignedTo 
               Height          =   315
               Left            =   2280
               TabIndex        =   11
               Top             =   280
               Width           =   2835
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "5001;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
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
            Index           =   0
            Left            =   11040
            TabIndex        =   7
            Top             =   945
            Width           =   285
         End
         Begin VB.CommandButton cmdType 
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
            Left            =   5295
            TabIndex        =   5
            Top             =   945
            Width           =   285
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
            Left            =   5310
            TabIndex        =   0
            Top             =   225
            Width           =   300
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
            Left            =   5310
            TabIndex        =   1
            Top             =   570
            Width           =   300
         End
         Begin MSForms.TextBox txtFund 
            Height          =   285
            Left            =   7830
            TabIndex        =   108
            Top             =   2400
            Width           =   3195
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "5636;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBudgetYears 
            Height          =   285
            Left            =   7845
            TabIndex        =   107
            Top             =   2085
            Width           =   3195
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "5636;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPropertyName 
            Height          =   315
            Left            =   1710
            TabIndex        =   97
            Top             =   570
            Width           =   3600
            VariousPropertyBits=   746604575
            Size            =   "6350;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtClientList 
            Height          =   285
            Left            =   1710
            TabIndex        =   96
            Top             =   225
            Width           =   3600
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "6350;503"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
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
            Left            =   150
            TabIndex        =   70
            Top             =   565
            Width           =   645
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
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
            Left            =   165
            TabIndex        =   69
            Top             =   225
            Width           =   465
         End
         Begin VB.Label lblRentCharges 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Budget Year:"
            Height          =   195
            Index           =   1
            Left            =   5805
            TabIndex        =   68
            Top             =   2070
            Width           =   930
         End
         Begin MSForms.CheckBox chkBudgetOverride 
            Height          =   285
            Left            =   4110
            TabIndex        =   25
            Top             =   3270
            Width           =   1515
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2681;503"
            Value           =   "0"
            Caption         =   "Budget Override"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fund:"
            Height          =   195
            Index           =   34
            Left            =   5805
            TabIndex        =   67
            Top             =   2430
            Width           =   390
         End
         Begin MSForms.CheckBox cbUrgent 
            Height          =   285
            Left            =   4725
            TabIndex        =   29
            Top             =   3945
            Width           =   855
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1508;503"
            Value           =   "0"
            Caption         =   "Urgent"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   255
            Left            =   165
            TabIndex        =   66
            Top             =   5370
            Width           =   855
            VariousPropertyBits=   8388627
            Caption         =   "Instruction:"
            Size            =   "1508;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label13 
            Height          =   255
            Left            =   1725
            TabIndex        =   65
            Top             =   4035
            Width           =   2415
            VariousPropertyBits=   8388627
            Caption         =   "(dd/mm/yyyy            hh:mm - 24 Hrs)"
            Size            =   "4260;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtBudgetCost 
            Height          =   315
            Left            =   1725
            TabIndex        =   24
            Top             =   3270
            Width           =   1335
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "2355;556"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtNextRemTime 
            Height          =   315
            Left            =   2925
            TabIndex        =   27
            Top             =   3705
            Width           =   1095
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "1931;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtActualCost 
            Height          =   315
            Left            =   7845
            TabIndex        =   21
            Top             =   2745
            Width           =   3525
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            Size            =   "6218;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtRef 
            Height          =   315
            Left            =   7005
            TabIndex        =   2
            Top             =   225
            Width           =   4350
            VariousPropertyBits=   746604573
            MaxLength       =   9
            BorderStyle     =   1
            Size            =   "7673;556"
            SpecialEffect   =   0
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboType 
            Height          =   315
            Left            =   1725
            TabIndex        =   4
            Top             =   945
            Width           =   3570
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6297;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox cbAlarm 
            Height          =   285
            Left            =   4725
            TabIndex        =   28
            Top             =   3690
            Width           =   735
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "1296;503"
            Value           =   "0"
            Caption         =   "Alarm"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboTaskOwner 
            Height          =   315
            Left            =   7005
            TabIndex        =   6
            Top             =   945
            Width           =   3990
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "7038;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label14 
            Height          =   180
            Left            =   5820
            TabIndex        =   64
            Top             =   3150
            Width           =   1335
            VariousPropertyBits=   8388627
            Caption         =   "Date Completed:"
            Size            =   "2355;317"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label12 
            Height          =   255
            Left            =   165
            TabIndex        =   63
            Top             =   3705
            Width           =   2295
            VariousPropertyBits=   8388627
            Caption         =   "Next Reminder:"
            Size            =   "4048;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label11 
            Height          =   180
            Left            =   5805
            TabIndex        =   62
            Top             =   2745
            Width           =   975
            VariousPropertyBits=   8388627
            Caption         =   "Actual Cost:"
            Size            =   "1720;317"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label10 
            Height          =   255
            Left            =   165
            TabIndex        =   61
            Top             =   3270
            Width           =   975
            VariousPropertyBits=   8388627
            Caption         =   "Budget Cost:"
            Size            =   "1720;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label9 
            Height          =   195
            Left            =   165
            TabIndex        =   60
            Top             =   4290
            Width           =   855
            VariousPropertyBits=   8388627
            Caption         =   "Job Details:"
            Size            =   "1508;344"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label8 
            Height          =   255
            Left            =   5805
            TabIndex        =   59
            Top             =   1725
            Width           =   2055
            VariousPropertyBits=   8388627
            Caption         =   "Expected Completion Date:"
            Size            =   "3625;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label7 
            Height          =   255
            Left            =   5805
            TabIndex        =   58
            Top             =   1380
            Width           =   1575
            VariousPropertyBits=   8388627
            Caption         =   "Expected Start Date:"
            Size            =   "2778;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label6 
            Height          =   255
            Left            =   165
            TabIndex        =   57
            Top             =   2865
            Width           =   1095
            VariousPropertyBits=   8388627
            Caption         =   "Date Reported:"
            Size            =   "1931;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   255
            Left            =   5805
            TabIndex        =   56
            Top             =   945
            Width           =   1095
            VariousPropertyBits=   8388627
            Caption         =   "Task Owner:"
            Size            =   "1931;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   255
            Left            =   165
            TabIndex        =   55
            Top             =   945
            Width           =   615
            VariousPropertyBits=   8388627
            Caption         =   "Type:"
            Size            =   "1085;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtJobName 
            Height          =   315
            Left            =   7005
            TabIndex        =   3
            Top             =   585
            Width           =   4365
            VariousPropertyBits=   746604571
            MaxLength       =   40
            BorderStyle     =   1
            Size            =   "7699;556"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label lblJobName 
            Height          =   255
            Left            =   5805
            TabIndex        =   54
            Top             =   585
            Width           =   855
            VariousPropertyBits=   8388627
            Caption         =   "Job Name:"
            Size            =   "1508;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label1 
            Height          =   255
            Left            =   5820
            TabIndex        =   53
            Top             =   240
            Width           =   735
            VariousPropertyBits=   8388627
            Caption         =   "Job  No.:"
            Size            =   "1296;450"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Memo"
         Height          =   3840
         Left            =   -74955
         TabIndex        =   71
         Top             =   315
         Width           =   11535
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   11385
            ScaleHeight     =   3015
            ScaleWidth      =   11385
            TabIndex        =   73
            Top             =   135
            Width           =   11385
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
               Left            =   11115
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtMemoAll 
               Height          =   2685
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               Top             =   315
               Width           =   11385
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
               Width           =   11070
            End
            Begin MSForms.Label lblSea 
               Height          =   195
               Left            =   135
               TabIndex        =   76
               Top             =   0
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
         Begin VB.CommandButton cmdUnitMemoCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   405
            Left            =   9870
            TabIndex        =   82
            Top             =   3270
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   405
            Left            =   7560
            TabIndex        =   81
            Top             =   3270
            Width           =   1125
         End
         Begin VB.CommandButton cmdUnitMemoEdit 
            Caption         =   "&Edit"
            Height          =   405
            Left            =   6375
            TabIndex        =   80
            Top             =   3270
            Width           =   1125
         End
         Begin VB.TextBox txtUnitMemo 
            Height          =   1335
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            Top             =   210
            Width           =   11385
         End
         Begin VB.CommandButton cmdUnitMemoNew 
            Caption         =   "&New"
            Height          =   405
            Left            =   5355
            TabIndex        =   78
            Top             =   3270
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&View All Memo"
            Height          =   405
            Left            =   3825
            TabIndex        =   77
            Top             =   3285
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   405
            Left            =   8730
            TabIndex        =   72
            Top             =   3270
            Width           =   1125
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridLeaseAnalysis 
            Height          =   1305
            Left            =   90
            TabIndex        =   83
            Top             =   1845
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2302
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
         Begin VB.TextBox txtLeaseAnalysisID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10170
            TabIndex        =   88
            Top             =   45
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSForms.Label Label17 
            Height          =   195
            Left            =   225
            TabIndex        =   87
            Top             =   1575
            Width           =   420
            VariousPropertyBits=   8388627
            Caption         =   "No"
            Size            =   "741;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label16 
            Height          =   195
            Left            =   1680
            TabIndex        =   86
            Top             =   1575
            Width           =   915
            VariousPropertyBits=   8388627
            Caption         =   "Description"
            Size            =   "1614;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label15 
            Height          =   195
            Left            =   9540
            TabIndex        =   85
            Top             =   1575
            Width           =   1095
            VariousPropertyBits=   8388627
            Caption         =   "User"
            Size            =   "1931;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label5 
            Height          =   195
            Left            =   630
            TabIndex        =   84
            Top             =   1575
            Width           =   420
            VariousPropertyBits=   8388627
            Caption         =   "Date"
            Size            =   "741;344"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   4
            Left            =   90
            Top             =   1575
            Width           =   11340
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   870
      Left            =   90
      TabIndex        =   89
      Top             =   6840
      Width           =   11670
      Begin VB.TextBox txtModifiedBy 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   95
         Top             =   180
         Width           =   2040
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   94
         Top             =   180
         Width           =   2040
      End
      Begin MSForms.Label Label21 
         Height          =   255
         Left            =   3465
         TabIndex        =   93
         Top             =   180
         Width           =   1440
         VariousPropertyBits=   8388627
         Caption         =   "Last Modified By:"
         Size            =   "2540;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label18 
         Height          =   255
         Left            =   180
         TabIndex        =   92
         Top             =   180
         Width           =   1440
         VariousPropertyBits=   8388627
         Caption         =   "User Name:"
         Size            =   "2540;450"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   420
         Left            =   9870
         TabIndex        =   91
         Top             =   225
         Width           =   1455
         Caption         =   "Cancel"
         Size            =   "2566;741"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   420
         Left            =   8190
         TabIndex        =   90
         Top             =   225
         Width           =   1455
         Caption         =   "Save"
         Size            =   "2566;741"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "frmMaintenanceJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isEdit As Boolean
Public UpdateRow As Integer
Public RecordType As String
Public CallingForm As String

Private szClient  As String
Private Created_Ref As String
Private flgLoadEdit As Boolean
Dim Job_ANALYSIS_NEW_ENTRY  As Boolean
Dim sTextBox As String
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
           rRow = 1
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
'Private Sub cboClientList_Change()
'   If cboClientList.ListIndex < 0 Then Exit Sub
'   szClient = txtClientList.tag
'
'   Dim adoConn    As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   LoadProperty adoConn, cboPropertyList
'   LoadFY adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub



Private Sub cboClientList_GotFocus()
   txtRef.Enabled = False
End Sub

'Private Sub cboPropertyList_Click()
'   If flgLoadEdit Then Exit Sub
'
''   If txtClientList.tag < 0 Then
''      cboClientList.SetFocus
''      Exit Sub
''   End If
'   If cboPropertyList.ListIndex < 0 Then Exit Sub
'   txtRef.text = txtPropertyName.Tag & "-" & getNextRef
'   'Resolved by BOSL
'   'issue 474 Note 5
'   'Modified by anol 28 Sep 2014
'   Dim rstCBY As New ADODB.Recordset
'   Dim Conn As New ADODB.Connection
'   Dim szSQL1 As String
'   If Conn.State = 0 Then
'      Conn.Open getConnectionString
'   End If
'   szSQL1 = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode,CBY " & _
'           "FROM Property " & _
'           "WHERE PropertyID = '" & txtPropertyName.Tag & "' " & _
'           "ORDER BY PropertyID;"
'    rstCBY.Open szSQL1, Conn, adOpenStatic, adLockReadOnly
'    If rstCBY.EOF = False Then
'         cboBudgetYears.Value = rstCBY("CBY").Value
'    Else
'         cboBudgetYears.ListIndex = -1
'   End If
'   rstCBY.Close
'   Set rstCBY = Nothing
'   If Conn.State = 1 Then
'      Conn.Close
'   End If
'
'End Sub

Private Sub LoadProperty(adoConn As ADODB.Connection, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
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





Private Sub cboAssignedTo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtExpCompletionDate.SetFocus
    End If
End Sub

Private Sub cboTaskOwner_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        txtExpStartDate.SetFocus
    End If
End Sub



Private Sub cboType_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
        cboTaskOwner.SetFocus
    End If
End Sub



Private Sub cmdBudgetYears_Click()
    sTextBox = "3"
    picClient.Left = 5805
    picClient.Top = 450
    picClient.Visible = True
    LoadGridFY
    cmdClientList.Enabled = False
    cmdProperty.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 2500
   flxClient.ColWidth(2) = 3500
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
   lblClientID.Caption = "Financial Year"
   lblClientName.Caption = "Financial Year Description"

   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   adoConn.Open getConnectionString
           
        szSQL = "SELECT FYrID, FinancialYear, FY_Description " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.PropertyID = '" & txtPropertyName.Tag & "' order by FinancialYear Desc ;"


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
      ShowMsgInTaskBar "Financial year has not been created.", "Y", "N"
   Else
            rRow = 1

        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = " " & Trim(rstRec.Fields.Item("FYrID").Value)
           flxClient.TextMatrix(rRow, 1) = Trim(rstRec.Fields.Item("FinancialYear").Value)
           flxClient.TextMatrix(rRow, 2) = Trim(rstRec.Fields.Item("FY_Description").Value)
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 585
    picClient.Visible = True
    LoadflxClient
    cmdProperty.Enabled = False
    cmdClientList.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdCloseMemo_Click()
   Picture2.Visible = False
   txtUnitMemo.SetFocus
   Command1.Visible = True
End Sub
Private Sub LoadCmbClient(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
                adoRst.Close
'                szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
            szSQL = "SELECT FYrID, FinancialYear, FY_Description, P.PropertyID, P.PropertyName " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.ClientID = '" & txtClientList.Tag & "' ORDER BY P.PropertyID;"
                adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields("PropertyName").Value
                        txtPropertyName.Tag = adoRst.Fields("PropertyID").Value
'                        txtBudgetYears.text = adoRst.Fields("FinancialYear").Value
'                        txtBudgetYears.Tag = adoRst.Fields("FYrID").Value
                        txtRef.text = txtPropertyName.Tag & "-" & getNextRef
                        LoadSingleFY adoConn
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
'                        txtBudgetYears.text = ""
'                        txtBudgetYears.Tag = ""
                End If
   Else
            adoRst.Close
   End If
               
End Sub
Private Sub cmdDelete_Click()
      Picture2.Visible = False
      If txtLeaseAnalysisID.text = "" Then
          ShowMsgInTaskBar "Please select a memo you want to delete", "Y"
          If gridLeaseAnalysis.Enabled = True Then
               gridLeaseAnalysis.SetFocus
          End If
          Exit Sub
      End If
      If MsgBox("Are you sure to delete memo?", vbQuestion + vbYesNo, "Delete Memo") = vbNo Then Exit Sub
      Dim adoConn As New ADODB.Connection
      If JobIDExists(txtRef.text) = True Then
         adoConn.Open getConnectionString
         adoConn.Execute "DELETE from MemoDetails where MemoID=" & Val(gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 1)) & " and sageaccountNumber='" & txtRef.text & "'"
         adoConn.Close
      Else
         gridLeaseAnalysis.RemoveItem gridLeaseAnalysis.row
      End If
      MsgBox "Memo has been deleted successfully", vbInformation + vbOKOnly, "Delete Memo"
      txtLeaseAnalysisID.text = ""
      txtUnitMemo.text = ""
      PopulateGridLeaseAnalysis
      Command1.Enabled = True
      Command1.Visible = False
      Picture2.Visible = True
      Call ViewMemo
      txtMemoAll.SetFocus
End Sub

Private Sub cmdDmdGridUnitLookup_Click()
   picDmdLeaseList.Visible = False
End Sub

Private Sub cmdFund_Click()
    sTextBox = "4"
    picClient.Left = 5805
    picClient.Top = 450
    picClient.Visible = True
    loadgridFund
    cmdClientList.Enabled = False
    cmdProperty.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdClientList.Enabled = True
    cmdProperty.Enabled = True
    cmdBudgetYears.Enabled = True
End Sub

Private Sub cmdProperty_Click()
    sTextBox = "2"
    picClient.Left = 915
    picClient.Top = 685
    picClient.Visible = True
    LoadPropertyList
    cmdProperty.Enabled = False
    cmdClientList.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    txtSearchClientID.SetFocus
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
   lblClientName.Caption = "Property Name"
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
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = ""
'           flxClient.TextMatrix(rRow, 2) = ""
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
Private Sub loadgridFund()
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
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"
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
           
         szSQL = "SELECT FundID, FundCode, FundName FROM Fund;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            rRow = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = ""
'           flxClient.TextMatrix(rRow, 2) = ""
'           flxClient.RowHeight(rRow) = 240
'           flxClient.AddItem ""
'           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = " " & rstRec.Fields.Item("FundID").Value
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("FundCode").Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("FundName").Value
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
Private Sub cmdSave_Click()
'Resolved by BOSL
'Modified by anol 28 Sep 2014
'issuer 474 note 8
   If cboType.ListIndex = -1 Then
      cboType.SetFocus
      ShowMsgInTaskBar "Please enter correct type", "Y", "Y"
      Exit Sub
   End If
   If txtRef.text = "" Then
         ShowMsgInTaskBar "Please Job ID missing", "Y", "Y"
      Exit Sub
   End If
   If txtClientList.Tag = "" Then
      cmdClientList.SetFocus
      ShowMsgInTaskBar "Please enter correct Client", "Y", "Y"
      Exit Sub
   End If
   If txtPropertyName.text = "" Then
      cmdProperty.SetFocus
      ShowMsgInTaskBar "Please enter correct Property", "Y", "Y"
      Exit Sub
   End If
   If cboTaskOwner.ListIndex = -1 Then
      cboTaskOwner.SetFocus
      ShowMsgInTaskBar "Please enter correct Task Owner", "Y", "Y"
      Exit Sub
   End If
   If cboAssignedTo.ListIndex = -1 Then
      cboAssignedTo.SetFocus
      ShowMsgInTaskBar "Please select Assigned To.", "Y", "Y"
      Exit Sub
   End If
   If optInternal_Reported.Value = True Then
      If cboReportedBy.ListIndex = -1 Then
         If cboReportedBy.Visible = True Then
            cboReportedBy.SetFocus
         End If
         ShowMsgInTaskBar "Please select reported by", "Y", "Y"
         Exit Sub
      End If
   End If
   If txtBudgetYears.Tag = "" Then
      cmdBudgetYears.SetFocus
      ShowMsgInTaskBar "Please select correct byudget year", "Y", "Y"
      Exit Sub
   End If
   If txtFund.text = "" Then
      cmdFund.SetFocus
      ShowMsgInTaskBar "Please enter correct fund", "Y", "Y"
      Exit Sub
   End If
   'End of modification
  ' If (Not validateMandatory(cboClientList, "Client Name cannot be empty")) Then Exit Sub

   'If (Not validateMandatory(cboPropertyList, "Property Name cannot be empty")) Then Exit Sub

   If (Not validateMandatory(txtJobName, "Job Name cannot be empty")) Then Exit Sub

   If (Not validateMandatory(cboType, "Type cannot be empty")) Then Exit Sub

   If (Not validateMandatory(cboTaskOwner, "Task Owner cannot be empty")) Then Exit Sub

   If (Not validateMandatory(cboAssignedTo, "Assigned To cannot be empty")) Then Exit Sub
   If optInternal_Reported.Value = True Then
      If (Not validateMandatory(cboReportedBy, "Reported by cannot be empty")) Then Exit Sub
   Else
      If (txtTenantName.text = "") Then
      MsgBox "Tenant Name by cannot be empty", vbInformation, "Mandatory Field"
      'cmdTask_Assigned.SetFocus
      Exit Sub
      End If
   End If
   If (Not validateMandatory(txtDateReported, "Date Reported cannot be empty")) Then Exit Sub

   If (Not validateMandatory(txtBudgetCost, "Please enter the amount")) Then Exit Sub

   'If (Not validateMandatory(cboBudgetYears, "Please select a budget year")) Then Exit Sub

   'If (Not validateMandatory(cboFund, "Fund cannot be empty")) Then Exit Sub

   If cbAlarm.Value > 0 Then
      If (Not validateMandatory(txtNextRemDate, "Reminder date cannot be empty")) Then
         Exit Sub
      End If
   End If

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   adoConn.BeginTrans
   If SavePropertyMaintenanceHistory(adoConn) Then
      If CallingForm = "P" Then
         ShowMsgInTaskBar "The Property maintenance event saved successfully.", "Y", "P"
         frmProperty2.LoadGridMaintenanceHistory adoConn
      End If
      If CallingForm = "U" Then
         ShowMsgInTaskBar "The Unit maintenance event saved successfully.", "Y", "P"
         frmUnits2.LoadGridMaintenanceHistory adoConn
      End If
      If CallingForm = "S" Then
         ShowMsgInTaskBar "The maintenance event saved successfully.", "Y", "P"
         frmSupplier.LoadGridMaintenanceHistory adoConn
      End If
      If CallingForm = "L" Then
         ShowMsgInTaskBar "The maintenance event saved successfully.", "Y", "P"
         frmLeasee1.LoadGridMaintenanceHistory adoConn
      End If
      If CallingForm = "M" Then
         ShowMsgInTaskBar "The maintenance event saved successfully.", "Y", "P"
         frmMaintenance.RefreshMaintenanceGrid adoConn
      End If
      'Save Memo here
      If MemoDataExists = True Then
            Dim rstLeaseAnalysis_ As New ADODB.Recordset
            rstLeaseAnalysis_.Open "SELECT * FROM MemoDetails", adoConn, adOpenDynamic, adLockOptimistic
            Dim i As Integer
            For i = 1 To gridLeaseAnalysis.Rows - 1
                  rstLeaseAnalysis_.AddNew
                  rstLeaseAnalysis_!MemoID = gridLeaseAnalysis.TextMatrix(i, 1)
                  rstLeaseAnalysis_!MemoType = "Job"
                  rstLeaseAnalysis_!SageAccountNumber = gridLeaseAnalysis.TextMatrix(i, 3)
                  rstLeaseAnalysis_!UpdateTime = gridLeaseAnalysis.TextMatrix(i, 4)
                  rstLeaseAnalysis_!MemoDescription = gridLeaseAnalysis.TextMatrix(i, 5)
                  rstLeaseAnalysis_!UserName = gridLeaseAnalysis.TextMatrix(i, 6)
                  rstLeaseAnalysis_.Update
            Next i
            rstLeaseAnalysis_.Close
      End If
      adoConn.CommitTrans
      adoConn.Close
      Set adoConn = Nothing
      Unload Me
    Else
         adoConn.RollbackTrans
         adoConn.Close
         Set adoConn = Nothing
         ShowMsgInTaskBar "Could not save maintenance event", , "N"
   End If
End Sub
Private Function MemoDataExists() As Boolean
     'Check if there is data in memo grid
      Dim Z As Integer
      If gridLeaseAnalysis.Rows > 1 Then
         For Z = 1 To gridLeaseAnalysis.Rows - 1
               If Len(gridLeaseAnalysis.TextMatrix(Z, 1)) > 0 Then
                  MemoDataExists = True
                  Exit Function
               End If
         Next Z
      End If
End Function
Private Function SaveLeaseAnalysis() As Boolean
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim conMemo As New ADODB.Connection
   Dim rstLease_ As New ADODB.Recordset
   conMemo.Open getConnectionString
   Dim sSQLQuery_ As String
   Dim sSQLFilter As String
   If Not Job_ANALYSIS_NEW_ENTRY Then
       sSQLFilter = "WHERE MemoID = " & Val(gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 1)) & " AND Memotype='Job' AND SageAccountNumber = '" & txtRef.text & "'"
   Else
       sSQLFilter = ""
   End If
   sSQLQuery_ = "SELECT * " & _
                "FROM MemoDetails " & sSQLFilter
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenDynamic, adLockOptimistic
   If Job_ANALYSIS_NEW_ENTRY Then rstLeaseAnalysis_.AddNew
   If Job_ANALYSIS_NEW_ENTRY = False Then
      rstLeaseAnalysis_!MemoID = txtLeaseAnalysisID.text
   Else
      rstLeaseAnalysis_!MemoID = NewMemoID()
   End If
   
   rstLeaseAnalysis_!MemoType = "Job"
   rstLeaseAnalysis_!SageAccountNumber = txtRef.text
   rstLeaseAnalysis_!MemoDescription = IIf(txtUnitMemo.text <> "", txtUnitMemo.text, "")
   rstLeaseAnalysis_!UpdateTime = Now
   rstLeaseAnalysis_!UserName = frmMMain.SystemUserName
   rstLeaseAnalysis_.Update
   rstLeaseAnalysis_.Close
   Set rstLease_ = Nothing
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   SaveLeaseAnalysis = True
End Function
Private Function NewMemoID() As Integer
   Dim conMemo As New ADODB.Connection
   conMemo.Open getConnectionString
   Dim szSQL As String
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT MAX(MemoID) AS x   " & _
                 "FROM MemoDetails where Memotype='Job';"
   rstSet.Open szSQL, conMemo, adOpenStatic, adLockReadOnly
   NewMemoID = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
   rstSet.Close
   Set rstSet = Nothing
   conMemo.Close
End Function
Private Function JobIDExists(strID As String) As Boolean
   Dim conMemo As New ADODB.Connection
   conMemo.Open getConnectionString
   Dim szSQL As String
   Dim rstSet As New ADODB.Recordset
   szSQL = "SELECT ID,LastModified,ModifiedBy  " & _
                 "FROM PropertyMaintHistory where ID='" & Right(strID, 9) & "';"
   rstSet.Open szSQL, conMemo, adOpenStatic, adLockReadOnly
   If rstSet.EOF = True Then
      JobIDExists = False
   Else
      JobIDExists = True
      txtModifiedBy.text = (IIf(IsNull(rstSet!LastModified), "", rstSet!LastModified))
      txtUserName.text = IIf(IsNull(rstSet!ModifiedBy), "", rstSet!ModifiedBy)
   End If
   
   rstSet.Close
   Set rstSet = Nothing
   conMemo.Close
End Function
Private Sub cmdUnitMemoCancel_Click()

   'Issue 488
   'Modified by anol 04 Oct 2014
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   'MemoButtonEnable False
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdDelete.Enabled = False
   txtUnitMemo.Locked = True
   gridLeaseAnalysis.Enabled = True
   Command1.Enabled = True
   Command1.Visible = False
   Picture2.Visible = True
   txtMemoAll.SetFocus
End Sub

Private Sub cmdUnitMemoEdit_Click()
'   If txtRef.text = "" Then
'      ShowMsgInTaskBar "This job is still not created.Create and save the job first.", "Y"
'      SSTab1.Tab = 0
'      Exit Sub
'   End If
'   If JobIDExists(txtRef.text) = False Then
'      ShowMsgInTaskBar "This job is still not saved.You need to save the job first", "Y"
'      SSTab1.Tab = 0
'      Exit Sub
'   End If
   'Modified by Anol 30 Nov 2014
   'Issue 474
      Picture2.Visible = False
      If txtLeaseAnalysisID.text = "" Then
            Command1.Enabled = True
      Else
            Command1.Enabled = False
      End If
      Command1.Visible = True
      If txtLeaseAnalysisID.text = "" Then
          ShowMsgInTaskBar "Please select the memo you would like to edit", "Y"
          If gridLeaseAnalysis.Enabled = True Then
               gridLeaseAnalysis.SetFocus
          End If
          Exit Sub
      End If
      
      
      cmdUnitMemoNew.Enabled = False
      cmdUnitMemoEdit.Enabled = False
      cmdUnitMemoSave.Enabled = True
      cmdUnitMemoCancel.Enabled = True
      gridLeaseAnalysis.Enabled = False
      txtUnitMemo.Locked = False
      Job_ANALYSIS_NEW_ENTRY = False
   If txtUnitMemo.Enabled = True Then
      txtUnitMemo.SetFocus
   End If
End Sub

Private Sub cmdUnitMemoNew_Click()
'   If txtRef.text = "" Then
'      ShowMsgInTaskBar "This job is still not created.Create and save the job first.", "Y"
'      SSTab1.Tab = 0
'      Exit Sub
'   End If
'   If JobIDExists(txtRef.text) = False Then
'      ShowMsgInTaskBar "This job is still not saved.You need to save the job first", "Y"
'      SSTab1.Tab = 0
'      Exit Sub
'   Else
      Job_ANALYSIS_NEW_ENTRY = True
      cmdUnitMemoNew.Enabled = False
      cmdUnitMemoEdit.Enabled = False
      cmdDelete.Enabled = False
      cmdUnitMemoSave.Enabled = True
      cmdUnitMemoCancel.Enabled = True
      gridLeaseAnalysis.Enabled = False
      txtUnitMemo.Locked = False
      Picture2.Visible = False
      txtUnitMemo.text = ""
      txtUnitMemo.SetFocus
   'End If
   
End Sub

Private Sub cmdUnitMemoSave_Click()
   
   If Len(txtUnitMemo.text) = 0 Then
      ShowMsgInTaskBar "Please enter description of memo", "Y"
      If txtUnitMemo.Enabled = True Then
         txtUnitMemo.SetFocus
      End If
      Exit Sub
   End If
  If gridLeaseAnalysis.row = 0 And Job_ANALYSIS_NEW_ENTRY = False Then
      ShowMsgInTaskBar "Please select a memo from list", "Y"
      Exit Sub
  End If
   If JobIDExists(txtRef.text) = True Then
      If SaveLeaseAnalysis Then
         ShowMsgInTaskBar "The memo has been saved successfully."
         PopulateGridLeaseAnalysis
      Else
         ShowMsgInTaskBar "Could not save memo.", , "N"
      End If
   Else
   'If job ID not in database then I need to just add to the grid
      With gridLeaseAnalysis
            gridLeaseAnalysis.RowHeight(0) = 0
            If Job_ANALYSIS_NEW_ENTRY Then
                  .AddItem ""
            End If
            .TextMatrix(.Rows - 1, 0) = .Rows - 1 '"SL"
            If Job_ANALYSIS_NEW_ENTRY = False Then
               .TextMatrix(.Rows - 1, 1) = txtLeaseAnalysisID.text '"MemoID"
            Else
               .TextMatrix(.Rows - 1, 1) = .Rows - 1 '"MemoID"
            End If
            .TextMatrix(.Rows - 1, 2) = "Job" '"MemoType"
            .TextMatrix(.Rows - 1, 3) = txtRef.text '"SageAccountNumber"
            .TextMatrix(.Rows - 1, 4) = Now  '"UpdateTime"
            .TextMatrix(.Rows - 1, 5) = txtUnitMemo.text '"MemoDescription"
            .TextMatrix(.Rows - 1, 6) = frmMMain.SystemUserName '"UserName"
      End With
   End If
   cmdUnitMemoNew.Enabled = True
   cmdUnitMemoEdit.Enabled = True
   cmdUnitMemoSave.Enabled = False
   cmdUnitMemoCancel.Enabled = False
   gridLeaseAnalysis.Enabled = True
   gridLeaseAnalysis.row = 0
   txtUnitMemo.text = ""
   txtLeaseAnalysisID.text = ""
   txtUnitMemo.Locked = True
   cmdDelete.Enabled = False
   
   txtMemoAll.text = ""
   Call ViewMemo
   
   Command1.Enabled = True
   Command1.Visible = False
   Picture2.Visible = True
   txtMemoAll.SetFocus
End Sub

Private Sub Command1_Click()
   txtLeaseAnalysisID.text = ""
   Picture2.Visible = True
   txtMemoAll.text = ""
   cmdCloseMemo.Refresh
   Call ViewMemo
   Command1.Visible = False
End Sub



Private Sub flxClient_Click()
    Frame1.Enabled = True
    Frame2.Enabled = True
        
            cmdClientList.Enabled = True
            cmdProperty.Enabled = True
            
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
        If sTextBox = "1" Then
               
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = ""
                txtPropertyName.Tag = ""
               
                Dim adoRst As New ADODB.Recordset
                Dim szSQL As String

                szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
                adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields(1).Value
                        txtPropertyName.Tag = adoRst.Fields(0).Value
                         txtRef.text = txtPropertyName.Tag & "-" & getNextRef
                        LoadSingleFY adoConn
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
                End If
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                If isEdit = False Then
                    txtRef.text = txtPropertyName.Tag & "-" & getNextRef
                End If
                LoadSingleFY adoConn
                txtJobName.SetFocus
        End If
        If sTextBox = "3" Then
                txtBudgetYears.text = flxClient.TextMatrix(flxClient.row, 1)
                txtBudgetYears.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                If optInternal_Reported.Value = True Then
                    txtTenantName.SetFocus
                End If
        End If
        If sTextBox = "4" Then
                txtFund.text = flxClient.TextMatrix(flxClient.row, 1)
                txtFund.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                txtDateReported.SetFocus
        End If
        adoConn.Close
       
        picClient.Visible = False
End Sub



Private Sub txtActualCost_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtDateCompleted.SetFocus
    End If
End Sub

Private Sub txtDateCompleted_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtBudgetCost.SetFocus
    End If
End Sub

Private Sub txtDateReported_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtActualCost.SetFocus
    End If
End Sub

Private Sub txtJobName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
            cboType.SetFocus
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
          Frame1.Enabled = True
          Frame2.Enabled = True
          
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
           End If
    End If
End Sub
Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub

'Private Sub Command2_Click()
'   'gridLeaseAnalysis.RowHeight(0) = 0
'   gridLeaseAnalysis.Rows = 1
'   MsgBox gridLeaseAnalysis.row
'
'End Sub

Private Sub flxDmdLeaseList_Click()
   txtTenantName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1)
   picDmdLeaseList.Visible = False
   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchName.text = ""
   txtDmdTenantSearchUnitName.text = ""
End Sub

Private Sub gridLeaseAnalysis_RowColChange()
      txtLeaseAnalysisID.text = gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 0)
      txtUnitMemo.text = gridLeaseAnalysis.TextMatrix(gridLeaseAnalysis.row, 5)
      cmdUnitMemoEdit.Enabled = True
      cmdDelete.Enabled = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
'      If txtRef.text = "" Then
'        ' MsgBox "This job is still not created.Create and save the job first."
'         SSTab1.Tab = 0
'         Exit Sub
'      End If
       If JobIDExists(txtRef.text) = True Then
         Call PopulateGridLeaseAnalysis
         Call ViewMemo
      End If
   End If
End Sub
Private Sub ViewMemo()
   'Issue 488
   'Added by anol 04 Nov 2014
'   Dim conMemo As New ADODB.Connection
'   Dim rstLeaseAnalysis_ As New ADODB.Recordset
'   Dim sSQLQuery_ As String
'   conMemo.Open getConnectionString
'   txtMemoAll.text = ""
'   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtRef.text & "' And  MemoType='Job' order by MemoID"
'   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
'   Dim strTemp As String
'   While Not rstLeaseAnalysis_.EOF
'         If Len(rstLeaseAnalysis_!UpdateTime) > 0 Then
'               strTemp = " -  "
'         Else
'               strTemp = ""
'         End If
'         If Len(txtMemoAll.text) > 0 Then txtMemoAll.text = txtMemoAll.text & vbCrLf & vbCrLf
'         txtMemoAll.text = txtMemoAll.text & Left(rstLeaseAnalysis_!UpdateTime, 11) & strTemp & rstLeaseAnalysis_!UserName & vbCrLf & vbCrLf & IIf(IsNull(rstLeaseAnalysis_!MemoDescription) = True, "", rstLeaseAnalysis_!MemoDescription)
'         rstLeaseAnalysis_.MoveNext
'   Wend
'
'   rstLeaseAnalysis_.Close
'   Set rstLeaseAnalysis_ = Nothing
''   conMemo.Close

   
   Dim i As Integer
   Dim strTemp As String
   txtMemoAll.text = ""
   For i = 1 To gridLeaseAnalysis.Rows - 1
       If Len(gridLeaseAnalysis.TextMatrix(i, 4)) > 0 Then
            strTemp = " -  "
       Else
            strTemp = ""
       End If
       If Len(txtMemoAll.text) > 0 Then txtMemoAll.text = txtMemoAll.text & vbCrLf & vbCrLf
       txtMemoAll.text = txtMemoAll.text & Left(gridLeaseAnalysis.TextMatrix(i, 4), 11) & strTemp & gridLeaseAnalysis.TextMatrix(i, 6) & vbCrLf & vbCrLf & gridLeaseAnalysis.TextMatrix(i, 5)
   Next i
   
   
End Sub
Public Sub PopulateGridLeaseAnalysis()
   'Issue 488
   'Added by anol 30 Nov 2014
   Dim conMemo As New ADODB.Connection
   Dim rstLeaseAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   conMemo.Open getConnectionString
   sSQLQuery_ = "SELECT * from MemoDetails where SageAccountNumber='" & txtRef.text & "' And  MemoType='Job' order by MemoID"
   rstLeaseAnalysis_.Open sSQLQuery_, conMemo, adOpenStatic, adLockReadOnly
   Dim iRow As Integer
   iRow = 1

   gridLeaseAnalysis.Clear
   gridLeaseAnalysis.Rows = 1
   gridLeaseAnalysis.Cols = 7
   gridLeaseAnalysis.RowHeight(0) = 0
   If rstLeaseAnalysis_.EOF = True Then
       gridLeaseAnalysis.Rows = 2
   End If
   'SetgridLeaseAnalysisHeader conPropertyAnalysis
   While Not rstLeaseAnalysis_.EOF
      gridLeaseAnalysis.AddItem ""
      gridLeaseAnalysis.TextMatrix(iRow, 0) = iRow
      gridLeaseAnalysis.TextMatrix(iRow, 1) = rstLeaseAnalysis_!MemoID 'colwidth 0
      gridLeaseAnalysis.TextMatrix(iRow, 2) = rstLeaseAnalysis_!MemoType 'colwidth 0
      gridLeaseAnalysis.TextMatrix(iRow, 3) = rstLeaseAnalysis_!SageAccountNumber 'colwidth 0
      gridLeaseAnalysis.TextMatrix(iRow, 4) = rstLeaseAnalysis_!UpdateTime
      gridLeaseAnalysis.TextMatrix(iRow, 5) = rstLeaseAnalysis_!MemoDescription
      gridLeaseAnalysis.TextMatrix(iRow, 6) = rstLeaseAnalysis_!UserName
      
      rstLeaseAnalysis_.MoveNext
     
      iRow = iRow + 1
   Wend

   rstLeaseAnalysis_.Close
   Set rstLeaseAnalysis_ = Nothing
   conMemo.Close
   If iRow > 0 Then
      gridLeaseAnalysis.row = 0
   End If
End Sub
Private Sub txtBudgetCost_Change()
   If IsNumeric(txtBudgetCost.text) = False Then
      txtBudgetCost.text = ""
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
Private Sub ConfigGridJobAnalysis()
   Dim szHeader As String
   gridLeaseAnalysis.Clear
   gridLeaseAnalysis.Rows = 1
   gridLeaseAnalysis.Cols = 7
    szHeader$ = "<SL|<Date|<Description|>User"
    gridLeaseAnalysis.FormatString = szHeader$
   gridLeaseAnalysis.TextMatrix(0, 0) = "" '"SL"
   gridLeaseAnalysis.TextMatrix(0, 1) = "" '"MemoID"
   gridLeaseAnalysis.TextMatrix(0, 2) = "" '"MemoType"
   gridLeaseAnalysis.TextMatrix(0, 3) = "" '"SageAccountNumber"
   gridLeaseAnalysis.TextMatrix(0, 4) = "" '"UpdateTime"
   gridLeaseAnalysis.TextMatrix(0, 5) = "" ' "MemoDescription"
   gridLeaseAnalysis.TextMatrix(0, 6) = "" ' "UserName"
     
   gridLeaseAnalysis.ColWidth(0) = 450
   gridLeaseAnalysis.ColWidth(1) = 0
   gridLeaseAnalysis.ColWidth(2) = 0
   gridLeaseAnalysis.ColWidth(3) = 0
   gridLeaseAnalysis.ColWidth(4) = Label16.Left - Label5.Left + 50
   gridLeaseAnalysis.ColWidth(5) = Label15.Left - Label16.Left
   gridLeaseAnalysis.ColWidth(6) = 1810
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
Private Sub cmdTask_Assigned_Click(Index As Integer)
'issue 474
'added by anol 26 Nov 2014
   If Index = 2 And optLessee.Value = True Then
         Dim szSQL As String
         Dim adoConn As New ADODB.Connection
         Dim adoRst As New ADODB.Recordset
         If txtPropertyName.Tag = "" Then
         
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
                  "WHERE L.Status AND U.PropertyID = '" & txtPropertyName.Tag & "' " & _
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
         flxDmdLeaseList.TextMatrix(iRow, 3) = adoRst!UNITNUMBER
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
   'Dim adoConn As New ADODB.Connection

   frmSecondaryCode.PRIMARY_CODE_SHOW = "MNTJOB"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   
'   If Index = 0 Then
      populateCombo adoConn, sSQLQuery, cboTaskOwner
      
      populateCombo adoConn, sSQLQuery, frmMaintenance.cboReportedBy
'   End If
   
'   If Index = 1 Then
      populateCombo adoConn, sSQLQuery, cboAssignedTo
'   End If

'   If Index = 2 Then
      populateCombo adoConn, sSQLQuery, cboReportedBy
'   End If
   adoConn.Close
   Set adoConn = Nothing
   'added by anol 26 Nov 2014
   'issue 474
   
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

Private Sub Form_Activate()
   
   If isEdit Then
      loadEditValues
'   Else
'      If CallingForm = "P" Then
'         txtRef.text = frmProperty2.txtPropertyID.text & "-" & GetNextRef
'         szClient = frmProperty2.cboClientID.Value
'      End If
'      If CallingForm = "M" Then
'         txtRef.text = frmMaintenance.txtPropertyName.Tag & "-" & GetNextRef
'         szClient = frmMaintenance.cboClientList.Value
'      End If
'      If CallingForm = "U" Then txtRef.text = frmUnits2.cboProperty.BoundText & "-" & GetNextRef
   End If

   If RecordType = "J" Then
      Label1.Caption = "Job No.:"
      lblJobName.Caption = "Job Name:"
      Label9.Caption = "Job Details"
   Else
      Label1.Caption = "Diary Entry No.:"
      lblJobName.Caption = "Diary Entry Name:"
      Label9.Caption = "Diary Entry Details"
   End If

   If CallingForm = "L" Then
      optLessee.Value = True
      optInternal_Reported.Enabled = False
      'rem out by anol 27 06 2016
      'cboReportedBy.Value = True
      cboReportedBy.Locked = True
   End If
End Sub

'Private Sub LoadCmbClient(adoConn As ADODB.Connection, cboC As Control)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
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
'   cboC.Column() = Data()
'
'NoRes:
'   adoRst.Close
'   Set adoRst = Nothing
'
'   Exit Sub
'
'ErrorHandler:
'   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"
'
'   adoRst.Close
'   Set adoRst = Nothing
'End Sub

Private Sub Form_Load()
   Me.Width = 11790
   Me.Height = 8250
   Me.Top = 0
   Me.Left = 0
   flgLoadEdit = False
'   Me.BackColor = MODULEBACKCOLOR
'   fraAssignedTo.BackColor = MODULEBACKCOLOR
'   fraReportedBy.BackColor = MODULEBACKCOLOR
'   optInternal.BackColor = fraAssignedTo.BackColor
'   optSupplier.BackColor = fraAssignedTo.BackColor
   optInternal_Reported.BackColor = fraReportedBy.BackColor
   optLessee.BackColor = fraReportedBy.BackColor
   optSupplier.BackColor = fraReportedBy.BackColor
   
   SSTab1.Tab = 0
   ConfigGridJobAnalysis
   Picture2.Left = 90
   Picture2.Top = 180
   cmdCloseMemo.Refresh
   Dim adoConn As New ADODB.Connection
   
   adoConn.Open getConnectionString
   'LoadValues adoConn          'this method loads data into comboes
   Dim sSQLQuery As String

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
   populateCombo adoConn, sSQLQuery, cboReportedBy
   LoadCmbClient adoConn 'LOADING CLIENT , PROPERTY AND BudgetYears
   'LoadCboFund adoConn
   LoadSingleFY adoConn
   txtDateReported.text = Date
   SelTxtInCtrl txtDateReported

   txtExpStartDate.text = Date
   SelTxtInCtrl txtExpStartDate
   'end of LoadValues
   adoConn.Close
   Set adoConn = Nothing
   txtRef.Enabled = False
End Sub
Private Sub LoadSingleFY(adoConn As ADODB.Connection)
  ' Dim rRow    As Integer
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String


        'Resolved by BOSL
        'issue no 471
        'Modified by anol 11 Sep 2014
        Dim rstSQL  As New ADODB.Recordset
        szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID " & _
                "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
                "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
        rstSQL.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        If Not rstSQL.EOF Then
             txtBudgetYears.Tag = IIf(IsNull(rstSQL.Fields.Item("FYrID").Value), "", rstSQL.Fields.Item("FYrID").Value)
             txtBudgetYears.text = IIf(IsNull(rstSQL.Fields.Item("CBY").Value), "", rstSQL.Fields.Item("CBY").Value)
        Else
            txtBudgetYears.text = ""
            txtBudgetYears.Tag = ""
           ' ShowMsgInTaskBar "Please set the financial year for this property inGlobal Data.", , "N"
        End If
        'End of modification
   'End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading financial year.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub
'Private Sub LoadValues(adoConn As ADODB.Connection)
'
'End Sub

'Private Sub LoadFY(adoConn As ADODB.Connection)
'   Dim rRow    As Integer
'   Dim adoRst  As New ADODB.Recordset
'   Dim szSQL   As String
'
'   szSQL = "SELECT FYrID, FinancialYear, FY_Description, FY_StDate, FY_EndDate " & _
'           "FROM   FinancialYear AS F " & _
'           "WHERE  F.ClientID = '" & txtClientList.Tag & "';"
''   If CallingForm = "P" Then
''      szSQL = "SELECT FYrID, FinancialYear, FY_Description, FY_StDate, FY_EndDate " & _
''              "FROM   FinancialYear AS F, Property AS P " & _
''              "WHERE  F.ClientID = P.ClientID AND " & _
''                     "P.PropertyID = '" & frmProperty2.txtPropertyID.text & "';"
''   End If
''   If CallingForm = "M" Then
''      szSQL = "SELECT FYrID, FinancialYear, FY_Description, FY_StDate, FY_EndDate " & _
''              "FROM   FinancialYear AS F, Property AS P " & _
''              "WHERE  F.ClientID = P.ClientID AND " & _
''                     "P.PropertyID = '" & frmMaintenance.txtPropertyName.Tag & "';"
''   End If
''   If CallingForm = "U" Then
''      szSQL = "SELECT FYrID, FinancialYear, FY_Description, FY_StDate, FY_EndDate " & _
''              "FROM   FinancialYear AS F, Property AS P " & _
''              "WHERE  F.ClientID = P.ClientID AND " & _
''                     "P.PropertyID = '" & frmUnits2.cboProperty.BoundText & "';"
''   End If
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      ShowMsgInTaskBar "Financial year has not created.", "Y", "N"
'   Else
'      ReDim Data(4, adoRst.RecordCount - 1) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = Trim(adoRst.Fields.Item("FYrID").Value)
'         Data(1, rRow) = Trim(adoRst.Fields.Item("FinancialYear").Value)
'         Data(2, rRow) = Trim(adoRst.Fields.Item("FY_Description").Value)
'         Data(3, rRow) = Trim(adoRst.Fields.Item("FY_StDate").Value)
'         Data(4, rRow) = Trim(adoRst.Fields.Item("FY_EndDate").Value)
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'
'      cboBudgetYears.Clear
'      cboBudgetYears.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ShowMsgInTaskBar "Error in Loading financial year.", , "N"
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Private Sub SupplierAccCombo()
   Dim adoConn As New ADODB.Connection
   Dim rstRst  As New ADODB.Recordset
   Dim szSQL   As String, Data() As String, i As Integer

   On Error GoTo ErrorHandler

'   Set the RDO Connections to the dataset
   adoConn.Open getConnectionString
   szSQL = "SELECT SupplierID, SupplierName " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'SUPPLIER' " & _
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
'   find the largest available ID
   
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
   txtRef.text = ""
   txtJobName.text = ""
   cboType.text = ""
   cboTaskOwner.text = ""
   cboAssignedTo.text = ""
   txtDateReported.text = ""
   txtExpStartDate.text = ""
   txtExpCompletionDate.text = ""
   txtJobDetail.text = ""
   txtBudgetCost.text = ""
   txtActualCost.text = ""
   txtNextRemDate.text = ""
   txtNextRemTime.text = ""
   cbAlarm.Value = False
   txtDateCompleted.text = ""
   isEdit = False
   flgLoadEdit = False

   If CallingForm = "P" Then _
      frmProperty2.Enabled = True
   If CallingForm = "M" Then _
      frmMaintenance.Enabled = True
   If CallingForm = "U" Then _
      frmUnits2.Enabled = True
   If CallingForm = "L" Then _
      frmLeasee1.Enabled = True
   If CallingForm = "S" Then _
      frmSupplier.Enabled = True
End Sub

Private Sub fraAssignedTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub optInternal_Click()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection
           
   adoConn.Open getConnectionString
   'AssignedTo
   sSQLQuery = "SELECT CODE, VALUE " & _
               "FROM SECONDARYCODE " & _
               "WHERE PRIMARYCODE = 'MNTJOB'"
   populateCombo adoConn, sSQLQuery, cboAssignedTo
   cmdTask_Assigned(1).Visible = True
End Sub

Private Sub optInternal_Reported_Click()
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
'   Dim szSQL As String
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'   'AssignedTo
'   szSQL = "SELECT T.SageAccountNumber, T.Name " & _
'           "FROM (Tenants AS T INNER JOIN LeaseDetails AS L ON " & _
'               "T.SageAccountNumber = L.SageAccountNumber) INNER JOIN Units AS U ON " & _
'               "L.UnitNumber = U.UnitNumber " & _
'           "WHERE L.Status AND U.PropertyID = '" & txtPropertyName.Tag & "' " & _
'           "ORDER BY T.Name;"
'   populateCombo adoConn, szSQL, cboReportedBy
'
'   adoConn.Close
'   Set adoConn = Nothing
   'cmdTask_Assigned(2).Visible = False
   cboReportedBy.Visible = False
   txtTenantName.Visible = True
   If txtTenantName.text = "" Then
      Call cmdTask_Assigned_Click(2)
   End If
End Sub

Private Sub optSupplier_Click()
   SupplierAccCombo
   cmdTask_Assigned(1).Visible = False
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
    If KeyAscii = 13 Then
        cmdBudgetYears.SetFocus
    End If
   TextBoxKeyPrsDate txtExpCompletionDate, KeyAscii
End Sub

Private Sub txtExpCompletionDate_LostFocus()
   TextBoxFormatDate txtExpCompletionDate
End Sub

Private Sub txtExpStartDate_Change()
   TextBoxChangeDate txtExpStartDate
End Sub

Private Sub txtExpStartDate_KeyPress(KeyAscii As Integer)
    If optInternal.Value = True And KeyAscii = 13 Then
           cboAssignedTo.SetFocus
    End If
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
   
'Resolved by BOSL
'Issue No: 0000487
'If the Reminder date is entered, set the alarm option on by default
'Modified By: Asif. 13 Oct 2014

   If txtNextRemDate.text <> "" Then
        cbAlarm.Value = True
   End If
End Sub

Private Sub txtNextRemTime_Change()
   prsTime
End Sub

Private Sub txtNextRemTime_LostFocus()
   validateTime
End Sub

Private Sub txtRef_LostFocus()
   If Not txtRef.Enabled Then Exit Sub
   If txtRef.text = "" Then Exit Sub

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
               ShowMsgInTaskBar "Ref Already Exist, Please change", , "N"
            End If
            rstRst.Close
            adoConn.Close
         End If
      End If
   End If
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
   flgLoadEdit = True
   txtRef.Enabled = False

   If CallingForm = "U" Then
'   IIF(RecordType = 'J', 'JOB', 'DIARY'), S.Value AS MaintenanceType,           0, 1
'   H.ReportedDate, H.ID AS Ref, H.Job_DiaryName, H.TaskOwner,                   2, 3, 4, 5
'   H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted,      6, 7, 8, 9
'   H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate,                 10, 11, 12
'   H.Detail, H.ActualCost, H.ReportedBy,                                        13, 14, 15
'   H.AssignedIL, H.ReportedIS, H.RemindTime, H.Urgent                           16, 17, 18, 19
'   H.MaintenanceType                                                            20
'   H.FundID, H.OverrideBudget, H.FYrID                                          22, 23, 24
'      With frmUnits2.gridMaintenanceHistory
'         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
'         txtRef.text = .TextMatrix(UpdateRow, 3)
'         txtJobName.text = .TextMatrix(UpdateRow, 4)
'         cboType.Value = .TextMatrix(UpdateRow, 20)
'         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
'         If .TextMatrix(UpdateRow, 16) = "I" Then
'            optInternal.Value = True
'         Else
'            optSupplier.Value = True
'         End If
'         cboAssignedTo.Value = .TextMatrix(UpdateRow, 6)
'         If .TextMatrix(UpdateRow, 17) = "I" Then
'            optInternal_Reported.Value = True
'         Else
'            optLessee.Value = True
'         End If
'         cboReportedBy.Value = .TextMatrix(UpdateRow, 15)
'         txtExpStartDate.text = .TextMatrix(UpdateRow, 11)
'         txtDateReported.text = .TextMatrix(UpdateRow, 2)
'         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 10), "0.00")
'         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 12)
'         txtNextRemDate.text = .TextMatrix(UpdateRow, 7)
'         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
'         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 8) = "YES", True, False)
'         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
'         txtDateCompleted.text = .TextMatrix(UpdateRow, 9)
'         txtActualCost.text = .TextMatrix(UpdateRow, 14)
'      End With
      szClient = frmMaintenance.txtClientList.Tag
      'szClient = frmMaintenance.cboClientList.Value
      With frmMaintenance.flxMaintenance
         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
         txtRef.text = .TextMatrix(UpdateRow, 3)
         txtJobName.text = .TextMatrix(UpdateRow, 4)
         cboType.Value = .TextMatrix(UpdateRow, 20)
         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
         'put client name and value here mark1
         'cboClientList.ListIndex = frmMaintenance.cboClientList.ListIndex
        ' cboPropertyList.ListIndex = frmMaintenance.cboPropertyList.ListIndex
         txtClientList.text = .TextMatrix(UpdateRow, 32)
         txtClientList.Tag = .TextMatrix(UpdateRow, 27)
         txtPropertyName.text = .TextMatrix(UpdateRow, 29)  'frmMaintenance.txtClientList.text
         txtPropertyName.Tag = .TextMatrix(UpdateRow, 26)
         If .TextMatrix(UpdateRow, 16) = "I" Then
            optInternal.Value = True
         Else
            optSupplier.Value = True
         End If
         cboAssignedTo.Value = .TextMatrix(UpdateRow, 7)
         
         If .TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
         
         txtExpStartDate.text = .TextMatrix(UpdateRow, 12)
         txtDateReported.text = .TextMatrix(UpdateRow, 2)
         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 11), "0.00")
         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 13)
         txtNextRemDate.text = .TextMatrix(UpdateRow, 8)
         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 9) = "YES", True, False)
         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
         txtDateCompleted.text = .TextMatrix(UpdateRow, 10)
         txtActualCost.text = .TextMatrix(UpdateRow, 15)
         txtFund.Tag = .TextMatrix(UpdateRow, 22)
          txtFund.text = .TextMatrix(UpdateRow, 33)
         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
         txtBudgetYears.Tag = .TextMatrix(UpdateRow, 24)
         txtBudgetYears.text = .TextMatrix(UpdateRow, 34)
      End With
   End If

   If CallingForm = "S" Then
'      With frmSupplier.gridMaintenanceHistory
'         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
'         txtRef.text = .TextMatrix(UpdateRow, 3)
'         txtJobName.text = .TextMatrix(UpdateRow, 4)
'         cboType.Value = .TextMatrix(UpdateRow, 20)
'         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
'         If .TextMatrix(UpdateRow, 16) = "I" Then
'            optInternal.Value = True
'         Else
'            optSupplier.Value = True
'         End If
'         cboAssignedTo.Value = .TextMatrix(UpdateRow, 6)
'         If .TextMatrix(UpdateRow, 17) = "I" Then
'            optInternal_Reported.Value = True
'         Else
'            optLessee.Value = True
'         End If
'         cboReportedBy.Value = .TextMatrix(UpdateRow, 15)
'         txtExpStartDate.text = .TextMatrix(UpdateRow, 11)
'         txtDateReported.text = .TextMatrix(UpdateRow, 2)
'         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 10), "0.00")
'         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 12)
'         txtNextRemDate.text = .TextMatrix(UpdateRow, 7)
'         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
'         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 8) = "YES", True, False)
'         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
'         txtDateCompleted.text = .TextMatrix(UpdateRow, 9)
'         txtActualCost.text = .TextMatrix(UpdateRow, 14)
'      End With
      szClient = frmMaintenance.txtClientList.Tag
      'szClient = frmMaintenance.cboClientList.Value
      With frmMaintenance.flxMaintenance
         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
         txtRef.text = .TextMatrix(UpdateRow, 3)
         txtJobName.text = .TextMatrix(UpdateRow, 4)
         cboType.Value = .TextMatrix(UpdateRow, 20)
         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
         'put client name and value here mark1
         'cboClientList.ListIndex = frmMaintenance.cboClientList.ListIndex
        ' cboPropertyList.ListIndex = frmMaintenance.cboPropertyList.ListIndex
         txtClientList.text = .TextMatrix(UpdateRow, 32)
         txtClientList.Tag = .TextMatrix(UpdateRow, 27)
         txtPropertyName.text = .TextMatrix(UpdateRow, 29)  'frmMaintenance.txtClientList.text
         txtPropertyName.Tag = .TextMatrix(UpdateRow, 26)
         If .TextMatrix(UpdateRow, 16) = "I" Then
            optInternal.Value = True
         Else
            optSupplier.Value = True
         End If
         cboAssignedTo.Value = .TextMatrix(UpdateRow, 7)
         
         If .TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
         
         txtExpStartDate.text = .TextMatrix(UpdateRow, 12)
         txtDateReported.text = .TextMatrix(UpdateRow, 2)
         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 11), "0.00")
         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 13)
         txtNextRemDate.text = .TextMatrix(UpdateRow, 8)
         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 9) = "YES", True, False)
         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
         txtDateCompleted.text = .TextMatrix(UpdateRow, 10)
         txtActualCost.text = .TextMatrix(UpdateRow, 15)
         txtFund.Tag = .TextMatrix(UpdateRow, 22)
          txtFund.text = .TextMatrix(UpdateRow, 33)
         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
         txtBudgetYears.Tag = .TextMatrix(UpdateRow, 24)
         txtBudgetYears.text = .TextMatrix(UpdateRow, 34)
      End With
   End If

   If CallingForm = "P" Then
'      szClient = frmProperty2.txtClientList.Tag
'      With frmProperty2.gridMaintenanceHistory
'         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
'         txtRef.text = .TextMatrix(UpdateRow, 3)
'         txtJobName.text = .TextMatrix(UpdateRow, 4)
'         cboType.Value = .TextMatrix(UpdateRow, 20)
'         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
'         If .TextMatrix(UpdateRow, 16) = "I" Then
'            optInternal.Value = True
'         Else
'            optSupplier.Value = True
'         End If
'         cboAssignedTo.Value = .TextMatrix(UpdateRow, 6)
'         If .TextMatrix(UpdateRow, 17) = "I" Then
'            optInternal_Reported.Value = True
'         Else
'            optLessee.Value = True
'         End If
'         cboReportedBy.Value = .TextMatrix(UpdateRow, 15)
'         txtExpStartDate.text = .TextMatrix(UpdateRow, 11)
'         txtDateReported.text = .TextMatrix(UpdateRow, 2)
'         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 10), "0.00")
'         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 12)
'         txtNextRemDate.text = .TextMatrix(UpdateRow, 7)
'         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
'         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 8) = "YES", True, False)
'         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
'         txtDateCompleted.text = .TextMatrix(UpdateRow, 9)
'         txtActualCost.text = .TextMatrix(UpdateRow, 14)
'         txtFund.Tag = .TextMatrix(UpdateRow, 22)
'         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
'         'cboBudgetYears.Value = .TextMatrix(UpdateRow, 24)
'      End With
      szClient = frmMaintenance.txtClientList.Tag
      'szClient = frmMaintenance.cboClientList.Value
      With frmMaintenance.flxMaintenance
         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
         txtRef.text = .TextMatrix(UpdateRow, 3)
         txtJobName.text = .TextMatrix(UpdateRow, 4)
         cboType.Value = .TextMatrix(UpdateRow, 20)
         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
         'put client name and value here mark1
         'cboClientList.ListIndex = frmMaintenance.cboClientList.ListIndex
        ' cboPropertyList.ListIndex = frmMaintenance.cboPropertyList.ListIndex
         txtClientList.text = .TextMatrix(UpdateRow, 32)
         txtClientList.Tag = .TextMatrix(UpdateRow, 27)
         txtPropertyName.text = .TextMatrix(UpdateRow, 29)  'frmMaintenance.txtClientList.text
         txtPropertyName.Tag = .TextMatrix(UpdateRow, 26)
         If .TextMatrix(UpdateRow, 16) = "I" Then
            optInternal.Value = True
         Else
            optSupplier.Value = True
         End If
         cboAssignedTo.Value = .TextMatrix(UpdateRow, 7)
         
         If .TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
         
         txtExpStartDate.text = .TextMatrix(UpdateRow, 12)
         txtDateReported.text = .TextMatrix(UpdateRow, 2)
         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 11), "0.00")
         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 13)
         txtNextRemDate.text = .TextMatrix(UpdateRow, 8)
         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 9) = "YES", True, False)
         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
         txtDateCompleted.text = .TextMatrix(UpdateRow, 10)
         txtActualCost.text = .TextMatrix(UpdateRow, 15)
         txtFund.Tag = .TextMatrix(UpdateRow, 22)
          txtFund.text = .TextMatrix(UpdateRow, 33)
         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
         txtBudgetYears.Tag = .TextMatrix(UpdateRow, 24)
         txtBudgetYears.text = .TextMatrix(UpdateRow, 34)
      End With
   End If

   If CallingForm = "M" Then
   
'   SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
'                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName(4), H.TaskOwner(5),H.ReportedBy(6), " & _
'                "H.AssignedTo(7), H.RemindDate(8), IIF(H.Alarm, 'YES', 'NO')(9), H.DateCompleted(10), " & _
'                "H.BudgetCost(11), H.ExpectedStartDate(12), H.ExpectedCompletionDate(13), " & _
'                "H.Detail(14), H.ActualCost(15),  H.AssignedIL(16), " & _
'                "H.ReportedIS(17), H.RemindTime(18), H.Urgent(19), H.MaintenanceType(20), " & _
'                "H.ReportedFrom(21), H.FundID(22), H.OverrideBudget(23), H.FYrID(24), " & _
'                "H.BudgetPassed(25), P.PropertyID(26), P.ClientID(27), " & _
'                "IIf(AssignedIL='S',U.SupplierOfficeEmail,S1.Description) AS EmailAdd (28)" & _

      szClient = frmMaintenance.txtClientList.Tag
      'szClient = frmMaintenance.cboClientList.Value
      With frmMaintenance.flxMaintenance
         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
         txtRef.text = .TextMatrix(UpdateRow, 3)
         txtJobName.text = .TextMatrix(UpdateRow, 4)
         cboType.Value = .TextMatrix(UpdateRow, 20)
         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
         'put client name and value here mark1
         'cboClientList.ListIndex = frmMaintenance.cboClientList.ListIndex
        ' cboPropertyList.ListIndex = frmMaintenance.cboPropertyList.ListIndex
         txtClientList.text = .TextMatrix(UpdateRow, 32)
         txtClientList.Tag = .TextMatrix(UpdateRow, 27)
         txtPropertyName.text = .TextMatrix(UpdateRow, 29)  'frmMaintenance.txtClientList.text
         txtPropertyName.Tag = .TextMatrix(UpdateRow, 26)
         If .TextMatrix(UpdateRow, 16) = "I" Then
            optInternal.Value = True
         Else
            optSupplier.Value = True
         End If
         cboAssignedTo.Value = .TextMatrix(UpdateRow, 7)
         
         If .TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
         
         txtExpStartDate.text = .TextMatrix(UpdateRow, 12)
         txtDateReported.text = .TextMatrix(UpdateRow, 2)
         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 11), "0.00")
         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 13)
         txtNextRemDate.text = .TextMatrix(UpdateRow, 8)
         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 9) = "YES", True, False)
         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
         txtDateCompleted.text = .TextMatrix(UpdateRow, 10)
         txtActualCost.text = .TextMatrix(UpdateRow, 15)
         txtFund.Tag = .TextMatrix(UpdateRow, 22)
          txtFund.text = .TextMatrix(UpdateRow, 33)
         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
         txtBudgetYears.Tag = .TextMatrix(UpdateRow, 24)
         txtBudgetYears.text = .TextMatrix(UpdateRow, 34)
      End With
   End If

   If CallingForm = "L" Then
'      With frmLeasee1.gridMaintenanceHistory
'         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
'         txtRef.text = .TextMatrix(UpdateRow, 3)
'         txtJobName.text = .TextMatrix(UpdateRow, 4)
'         cboType.Value = .TextMatrix(UpdateRow, 20)
'         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
'         If .TextMatrix(UpdateRow, 16) = "I" Then
'            optInternal.Value = True
'         Else
'            optSupplier.Value = True
'         End If
'         cboAssignedTo.Value = .TextMatrix(UpdateRow, 6)
'         If .TextMatrix(UpdateRow, 17) = "I" Then
'            optInternal_Reported.Value = True
'         Else
'            optLessee.Value = True
'         End If
'         cboReportedBy.Value = .TextMatrix(UpdateRow, 15)
'         txtExpStartDate.text = .TextMatrix(UpdateRow, 11)
'         txtDateReported.text = .TextMatrix(UpdateRow, 2)
'         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 10), "0.00")
'         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 12)
'         txtNextRemDate.text = .TextMatrix(UpdateRow, 7)
'         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
'         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 8) = "YES", True, False)
'         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
'         txtDateCompleted.text = .TextMatrix(UpdateRow, 9)
'         txtActualCost.text = .TextMatrix(UpdateRow, 14)
'      End With
'   SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
'                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName(4), H.TaskOwner(5),H.ReportedBy(6), " & _
'                "H.AssignedTo(7), H.RemindDate(8), IIF(H.Alarm, 'YES', 'NO')(9), H.DateCompleted(10), " & _
'                "H.BudgetCost(11), H.ExpectedStartDate(12), H.ExpectedCompletionDate(13), " & _
'                "H.Detail(14), H.ActualCost(15),  H.AssignedIL(16), " & _
'                "H.ReportedIS(17), H.RemindTime(18), H.Urgent(19), H.MaintenanceType(20), " & _
'                "H.ReportedFrom(21), H.FundID(22), H.OverrideBudget(23), H.FYrID(24), " & _
'                "H.BudgetPassed(25), P.PropertyID(26), P.ClientID(27), " & _
'                "IIf(AssignedIL='S',U.SupplierOfficeEmail,S1.Description) AS EmailAdd (28)" & _

      szClient = frmMaintenance.txtClientList.Tag
      'szClient = frmMaintenance.cboClientList.Value
      With frmLeasee1.gridMaintenanceHistory
         RecordType = Left(.TextMatrix(UpdateRow, 0), 1)
         txtRef.text = .TextMatrix(UpdateRow, 3)
         txtJobName.text = .TextMatrix(UpdateRow, 4)
         cboType.Value = .TextMatrix(UpdateRow, 20)
         cboTaskOwner.Value = .TextMatrix(UpdateRow, 5)
         'put client name and value here mark1
         'cboClientList.ListIndex = frmMaintenance.cboClientList.ListIndex
        ' cboPropertyList.ListIndex = frmMaintenance.cboPropertyList.ListIndex
         txtClientList.text = .TextMatrix(UpdateRow, 32)
         txtClientList.Tag = .TextMatrix(UpdateRow, 27)
         txtPropertyName.text = .TextMatrix(UpdateRow, 29)  'frmMaintenance.txtClientList.text
         txtPropertyName.Tag = .TextMatrix(UpdateRow, 26)
         If .TextMatrix(UpdateRow, 16) = "I" Then
            optInternal.Value = True
         Else
            optSupplier.Value = True
         End If
         cboAssignedTo.Value = .TextMatrix(UpdateRow, 7)
         
         If .TextMatrix(UpdateRow, 17) = "I" Then
            cboReportedBy.Value = .TextMatrix(UpdateRow, 6)
            optInternal_Reported.Value = True
         Else
            optLessee.Value = True
            txtTenantName.text = .TextMatrix(UpdateRow, 6)
            picDmdLeaseList.Visible = False
         End If
         
         txtExpStartDate.text = .TextMatrix(UpdateRow, 12)
         txtDateReported.text = .TextMatrix(UpdateRow, 2)
         txtBudgetCost.text = Format(.TextMatrix(UpdateRow, 11), "0.00")
         txtExpCompletionDate.text = .TextMatrix(UpdateRow, 13)
         txtNextRemDate.text = .TextMatrix(UpdateRow, 8)
         txtNextRemTime.text = .TextMatrix(UpdateRow, 18)
         cbAlarm.Value = IIf(.TextMatrix(UpdateRow, 9) = "YES", True, False)
         cbUrgent.Value = IIf(.TextMatrix(UpdateRow, 19) = "U", True, False)
         txtDateCompleted.text = .TextMatrix(UpdateRow, 10)
         txtActualCost.text = .TextMatrix(UpdateRow, 15)
         txtFund.Tag = .TextMatrix(UpdateRow, 22)
          txtFund.text = .TextMatrix(UpdateRow, 33)
         chkBudgetOverride.Value = .TextMatrix(UpdateRow, 23)
         txtBudgetYears.Tag = .TextMatrix(UpdateRow, 24)
         txtBudgetYears.text = .TextMatrix(UpdateRow, 34)
      End With
   End If

   RetrieveMemo "PropertyMaintHistory", "Detail", Mid(txtRef.text, 6), "ID", txtJobDetail
   RetrieveMemo "PropertyMaintHistory", "Instruction", Mid(txtRef.text, 6), "ID", txtInsruction

   SelTxtInCtrl txtDateCompleted
   SelTxtInCtrl txtNextRemDate
   SelTxtInCtrl txtNextRemTime
   SelTxtInCtrl txtExpCompletionDate
   Call JobIDExists(txtRef.text)
End Sub

Public Function SavePropertyMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection) As Boolean
   Dim rstMHistory_  As New ADODB.Recordset
   Dim rstID         As New ADODB.Recordset
   Dim sSQLQuery_    As String, sSQLDelete As String, sSQLFilter As String, iRowIndex As Integer
   Dim lTableID      As Long, szAlarmTime As String
   Dim szJobId       As String
   Dim szaTemp()     As String
   Dim bSuffFund     As Boolean     'sufficient fund?

   bSuffFund = True
   'Below line has been modified by anol
   'Message override was not working
   If GetBudgetBalance(conMHistory_) - Val(txtActualCost) < CCur(txtBudgetCost.text) Then
   'If GetBudgetBalance(conMHistory_) - GetActualBudgetBalance(conMHistory_) < CCur(txtBudgetCost.text) Then
   'Resolved by BOSL
   'issue 474
   'Modified by Anol 30 Nov 2014
      If chkBudgetOverride.Value = False Then
         If MsgBox("There are insufficient funds for this job. Do you wish to save the Job entry?", vbQuestion + vbYesNo, "Job Entry") = vbNo Then Exit Function
      End If
      bSuffFund = False
   End If

   sSQLFilter = ""
   If InStr(txtRef.text, "-") > 0 Then
      szaTemp = Split(txtRef.text, "-")
      szJobId = szaTemp(1)
   Else
      szJobId = txtRef.text
   End If

   If isEdit Then
'      If CallingForm = "P" Then _
'         sSQLFilter = "WHERE PropertyID = '" & frmProperty2.txtPropertyID.text & "' AND ID = '" & szJobId & "'"
'      If CallingForm = "M" Then _
'         sSQLFilter = "WHERE PropertyID = '" & frmMaintenance.txtPropertyName.Tag & "' AND ID = '" & szJobId & "'"
'      If CallingForm = "L" Then _
'         sSQLFilter = "WHERE ReportedBy = '" & frmLeasee1.txtTenantID.text & "'     AND ID = '" & szJobId & "'"
'      If CallingForm = "U" Then _
'         sSQLFilter = "WHERE UnitNumber = '" & frmUnits2.txtUnitNo.text & "'        AND ID = '" & szJobId & "'"
'      If CallingForm = "S" Then _
'         sSQLFilter = "WHERE AssignedTo = '" & frmSupplier.txtSupplierID.text & "'        AND ID = '" & szJobId & "'"
     If CallingForm = "P" Then _
         sSQLFilter = "WHERE  ID = '" & szJobId & "'"
      If CallingForm = "M" Then _
         sSQLFilter = "WHERE  ID = '" & szJobId & "'"
      If CallingForm = "L" Then _
         sSQLFilter = "WHERE  ID = '" & szJobId & "'"
      If CallingForm = "U" Then _
         sSQLFilter = "WHERE  ID = '" & szJobId & "'"
      If CallingForm = "S" Then _
         sSQLFilter = "WHERE  ID = '" & szJobId & "'"
   Else
      sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyMaintHistory " & sSQLFilter

   rstMHistory_.Open sSQLQuery_, conMHistory_, adOpenDynamic, adLockOptimistic

   If Not isEdit Then rstMHistory_.AddNew
'Exit Function
   rstMHistory_!Id = szJobId
'   If CallingForm = "P" Then _
'      rstMHistory_!PropertyID = frmProperty2.txtPropertyID.text
'   If CallingForm = "M" Then _
'      rstMHistory_!PropertyID = txtPropertyName.Tag
'   If CallingForm = "U" Then _
'      rstMHistory_!PropertyID = frmUnits2.txtUnitNo.text
'   If CallingForm = "L" Then _
'      rstMHistory_!PropertyID = frmLeasee1.txtTenantID.text

   rstMHistory_!propertyID = txtPropertyName.Tag
   rstMHistory_!Job_DiaryName = txtJobName.text
   rstMHistory_!MaintenanceType = cboType.Value
   rstMHistory_!TaskOwner = cboTaskOwner.Value
   
   

   rstMHistory_!ReportedDate = Format(IIf(txtDateReported.text = "", Now, txtDateReported.text), "DD/MM/YYYY")
   rstMHistory_!ExpectedStartDate = IIf(txtExpStartDate.text = "", "", Format(txtExpStartDate.text, "DD/MM/YYYY"))
   rstMHistory_!ExpectedCompletionDate = IIf(txtExpCompletionDate.text = "", "", Format(txtExpCompletionDate.text, "DD/MM/YYYY"))
   rstMHistory_!Detail = IIf(txtJobDetail.text = "", "", txtJobDetail.text)
   rstMHistory_!Instruction = IIf(txtInsruction.text = "", "", txtInsruction.text)
   rstMHistory_!BudgetCost = IIf(txtBudgetCost.text = "", "0.00", Format(txtBudgetCost.text, "0.00"))
   rstMHistory_!ActualCost = IIf(txtActualCost.text = "", "0.00", Format(txtActualCost.text, "0.00"))
   rstMHistory_!RemindTime = IIf(txtNextRemTime.text = "", "", txtNextRemTime.text)
      'Resolved by BOSL
      'Issue No: 0000474 Note 962 ,4
      'If the reminder date is not set, then uncheck the alarm check box.
      'Modified By: Asif. 14 Oct 2014
   rstMHistory_!LastModified = Now
   rstMHistory_!ModifiedBy = frmMMain.SystemUserName
   If (txtDateCompleted.text = "") Then
     rstMHistory_!DateCompleted = Null
   Else
     rstMHistory_!DateCompleted = CDate(Format(txtDateCompleted.text, "DD/MM/YYYY"))
   End If

   If (txtNextRemDate.text = "") Then
     rstMHistory_!RemindDate = Null
      'Resolved by BOSL
      'Issue No: 0000487
      'If the reminder date is not set, then uncheck the alarm check box.
      'Modified By: Asif. 14 Oct 2014
      cbAlarm.Value = False
   Else
     rstMHistory_!RemindDate = CDate(Format(txtNextRemDate.text, "DD/MM/YYYY"))
   End If

   rstMHistory_!RecordType = RecordType

   If cbAlarm.Value Then
      rstMHistory_!Alarm = True
      szAlarmTime = IIf(txtNextRemTime.text = "", "083000", Format(txtNextRemTime.text, "hhmm") & "00")

      'Resolved by BOSL
      'Issue No: 0000487
      'If the Reminder was not set for an existing Job, create a new reminder record.
      'Modified By: Asif. 13 Oct 2014
      If Not isEdit Or IsNull(rstMHistory_!Reminder_ID) Then
         rstMHistory_!Reminder_ID = NewReminder(Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), szAlarmTime, txtJobName.text, "PropertyMaintHistory", szJobId)
      Else
         UpdateReminder rstMHistory_!Reminder_ID, Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), szAlarmTime, txtJobName.text
      End If
   Else
      rstMHistory_!Alarm = False
   End If
   
   'Resolved by BOSL
   'Issue No: 0000474
   'modified by anol 30 Nov 2014
      
   rstMHistory_!AssignedIL = IIf(optInternal.Value, "I", "S")
   rstMHistory_!AssignedTo = cboAssignedTo.Value
   
   If optInternal_Reported.Value Then
      rstMHistory_!ReportedIS = "I"
      rstMHistory_!ReportedBy = cboReportedBy.Value
   Else
      rstMHistory_!ReportedIS = "L"
      rstMHistory_!ReportedBy = txtTenantName.text
      rstMHistory_!UNITNUMBER = GetUnitIDbyTenantID(txtTenantName.text, conMHistory_)
   End If
   
   
   
   rstMHistory_!Urgent = IIf(cbUrgent.Value, "U", "N")
   rstMHistory_!ReportedFrom = CallingForm
   rstMHistory_!fundID = txtFund.Tag
   rstMHistory_!BudgetPassed = bSuffFund
   rstMHistory_!OverrideBudget = chkBudgetOverride.Value
   rstMHistory_!FYrID = txtBudgetYears.Tag

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

'Private Sub LoadCboFund(adoConn As ADODB.Connection)
'   Dim rRow As Integer, iRec As Integer, Data() As String
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT FundID, FundCode, FundName FROM Fund;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      MsgBox "Fund has not been setup for this company.", vbExclamation, "Load Fund in Global"
'   Else
'      ReDim Data(2, adoRst.RecordCount) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = adoRst.Fields.Item("FundID").Value
'         Data(1, rRow) = adoRst.Fields.Item("FundCode").Value
'         Data(2, rRow) = adoRst.Fields.Item("FundName").Value
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'      cboFund.Clear
'      cboFund.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'End Sub

Private Function GetBudgetBalance(adoConn As ADODB.Connection) As Currency
   Dim szSQL      As String
   Dim adoRst     As New ADODB.Recordset

   szSQL = "SELECT G.TotalBudget " & _
           "FROM   GlobalSC AS G " & _
           "WHERE  G.PropertyID = '" & txtPropertyName.Tag & "' AND " & _
                  "G.FinancialYear = '" & txtBudgetYears.Tag & "' AND " & _
                  "G.Fund = " & txtFund.Tag & ";"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   If Not adoRst.EOF Then
      GetBudgetBalance = CCur(adoRst.Fields.Item(0).Value)
   Else
      GetBudgetBalance = 0
   End If
   adoRst.Close
   Set adoRst = Nothing
End Function

'Private Function GetActualBudgetBalance(adoConn As ADODB.Connection) As Currency
'   Dim cBalDr     As Currency
'   Dim cBalCr     As Currency
'
'   cBalDr = CalcuateBalanceDr_PnL(adoConn)
'   cBalCr = CalcuateBalanceCr_PnL(adoConn)
'
'   GetActualBudgetBalance = Format(cBalDr - cBalCr, "0.00")
'End Function

'Private Function CalcuateBalanceDr_PnL(adoConn As ADODB.Connection) As Currency
'comment out by anok 2016 Jun 26
'   Dim adoRst     As New ADODB.Recordset
'
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim szSQL_Fund As String
'
'   CalcuateBalanceDr_PnL = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'    szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'      Dim propertyID As String
'      If CallingForm = "P" Then
'         propertyID = frmProperty2.txtPropertyID.text
'      ElseIf CallingForm = "M" Then
'         propertyID = frmMaintenance.cboPropertyList.Value
'      ElseIf CallingForm = "L" Then
'         propertyID = frmLeasee1.txtTenantID.text
'      ElseIf CallingForm = "U" Then
'         propertyID = frmUnits2.txtUnitNo.text
'      ElseIf CallingForm = "S" Then
'         propertyID = frmSupplier.txtSupplierID.text
'      End If
'   szSQL_Fund = "(" & _
'                  "SELECT S.NC " & _
'                  "FROM   GlobalSC AS G INNER JOIN GlobalSCDtls AS S ON G.BudgetID = S.BudgetID " & _
'                  "WHERE  G.PropertyID = '" & propertyID & "' AND " & _
'                         "G.FinancialYear = '" & txtBudgetYears.Tag & "' AND " & _
'                         "G.Fund = " & txtFund.Tag & "" & _
'                ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE IN (" & szSQL_Fund & ") AND " & _
'               "N.TRANSACTION_DATE >= #" & Format(cboBudgetYears.Column(3), "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(cboBudgetYears.Column(4), "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 2 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 15 OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 11) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
''Debug.Print szSQL
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   If adoRst.EOF Then
'      CalcuateBalanceDr_PnL = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceDr_PnL = 0
'      Else
'         CalcuateBalanceDr_PnL = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
'Private Function CalcuateBalanceCr_PnL(adoConn As ADODB.Connection) As Currency
''Comment out by anol 2016 Jun 26
'   Dim adoRst     As New ADODB.Recordset
'
'   Dim szSQL      As String
'   Dim szSQL_S    As String
'   Dim szSQL_P    As String
'   Dim szSQL_I    As String
'   Dim szSQL_O    As String
'   Dim szSQL_Fund As String
'
'   CalcuateBalanceCr_PnL = 0
'
'   szSQL_S = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'S' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_P = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'P' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_I = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'I' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'   szSQL_O = "(" & _
'                  "SELECT S.Code " & _
'                  "FROM   NominalLedger AS S " & _
'                  "WHERE  S.CAType = 'O' AND S.ClientID = '" & szClient & "'" & _
'             ")"
'
'   szSQL_Fund = "(" & _
'                  "SELECT S.NC " & _
'                  "FROM   GlobalSC AS G INNER JOIN GlobalSCDtls AS S ON G.BudgetID = S.BudgetID " & _
'                  "WHERE  G.PropertyID = '" & frmProperty2.txtPropertyID.text & "' AND " & _
'                         "G.FinancialYear = '" & txtBudgetYears.Tag & "' AND " & _
'                         "G.Fund = " & txtFund.Tag & "" & _
'                ")"
'
'   szSQL = "SELECT SUM(N.AMOUNT) " & _
'           "FROM   NLPosting AS N " & _
'            "WHERE N.ClientID = '" & szClient & "' AND N.NOMINAL_CODE IN (" & szSQL_Fund & ") AND " & _
'               "N.TRANSACTION_DATE >= #" & Format(cboBudgetYears.Column(3), "dd mmmm yyyy") & "# AND N.TRANSACTION_DATE <= #" & Format(cboBudgetYears.Column(4), "dd mmmm yyyy") & "# AND " & _
'               "("
'   szSQL = szSQL & _
'                   "(N.TRANSACTION_TYPE = 7 AND (N.NOMINAL_CODE IN " & _
'                        szSQL_I & _
'                   ")) OR " & _
'                   "N.TRANSACTION_TYPE = 16 OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 12) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_O & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 1 OR N.TRANSACTION_TYPE = 23) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 7 OR N.TRANSACTION_TYPE = 8 OR N.TRANSACTION_TYPE = 9) AND (N.NOMINAL_CODE NOT IN " & _
'                        szSQL_P & _
'                   ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 12 AND (N.AMOUNT_TYPE = 'A' OR N.AMOUNT_TYPE = 'V')) OR " & _
'                   "((N.TRANSACTION_TYPE = 2 OR N.TRANSACTION_TYPE = 3 OR N.TRANSACTION_TYPE = 4) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_S & _
'                   ")) OR " & _
'                   "((N.TRANSACTION_TYPE = 6 OR N.TRANSACTION_TYPE = 24) AND (N.NOMINAL_CODE IN " & _
'                        szSQL_P & _
'                  ")) OR " & _
'                   "(N.TRANSACTION_TYPE = 11 AND N.AMOUNT_TYPE = 'B') " & _
'               ")"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'   If adoRst.EOF Then
'      CalcuateBalanceCr_PnL = 0
'   Else
'      If IsNull(adoRst.Fields.Item(0).Value) Then
'         CalcuateBalanceCr_PnL = 0
'      Else
'         CalcuateBalanceCr_PnL = CCur(adoRst.Fields.Item(0).Value)
'      End If
'   End If
'   adoRst.Close
'   Set adoRst = Nothing
'End Function
Private Sub txtTenantName_Change()

End Sub

Private Sub txtTenantName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdFund.SetFocus
    End If
End Sub
