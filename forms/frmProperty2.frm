VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProperty2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Property"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19155
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperty2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   19155
   Begin VB.PictureBox picSupList 
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
      Height          =   3075
      Left            =   12690
      ScaleHeight     =   3045
      ScaleWidth      =   5625
      TabIndex        =   253
      Top             =   4230
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtSupplierSearch 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4230
         TabIndex        =   257
         Top             =   300
         Width           =   1245
      End
      Begin VB.TextBox txtSupplierSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1275
         TabIndex        =   256
         Top             =   300
         Width           =   2940
      End
      Begin VB.TextBox txtSupplierSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   255
         TabIndex        =   255
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton cmdGridUnitClose2 
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
         Index           =   1
         Left            =   5360
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   15
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplierList 
         Height          =   2370
         Left            =   45
         TabIndex        =   258
         Top             =   645
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   4180
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
            Name            =   "Arial"
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
         Index           =   0
         Left            =   345
         TabIndex        =   261
         Top             =   75
         Width           =   165
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LandLord Name"
         Height          =   210
         Index           =   1
         Left            =   1350
         TabIndex        =   260
         Top             =   45
         Width           =   1140
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Balance"
         Height          =   210
         Index           =   2
         Left            =   4095
         TabIndex        =   259
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   45
         Top             =   60
         Width           =   5280
      End
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4175
      Left            =   12735
      ScaleHeight     =   4140
      ScaleWidth      =   6225
      TabIndex        =   214
      Top             =   0
      Visible         =   0   'False
      Width           =   6250
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
         TabIndex        =   215
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3435
         Left            =   45
         TabIndex        =   216
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6059
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
         TabIndex        =   222
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   221
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label Label8 
         Height          =   195
         Left            =   120
         TabIndex        =   220
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
      Begin MSForms.Label Label7 
         Height          =   195
         Left            =   1620
         TabIndex        =   219
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
         TabIndex        =   218
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
         TabIndex        =   217
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
         Index           =   0
         Left            =   45
         Top             =   75
         Width           =   5850
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
      Height          =   360
      Left            =   3960
      ScaleHeight     =   360
      ScaleWidth      =   2655
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   60
         Width           =   2475
      End
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
      Left            =   -120
      ScaleHeight     =   45
      ScaleWidth      =   12675
      TabIndex        =   38
      Top             =   3300
      Width           =   12735
   End
   Begin VB.PictureBox fmePropertyLookup 
      BackColor       =   &H00B3C0C6&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3600
      Left            =   6660
      ScaleHeight     =   3540
      ScaleWidth      =   8940
      TabIndex        =   27
      Top             =   7590
      Visible         =   0   'False
      Width           =   9000
      Begin VB.CommandButton cmdGridPropertyLookup 
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
         Left            =   8685
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPropertyLookup 
         Height          =   2850
         Left            =   15
         TabIndex        =   33
         Top             =   630
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   5027
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
      Begin MSForms.TextBox TextBox3 
         Height          =   315
         Left            =   7650
         TabIndex        =   212
         Top             =   270
         Width           =   1110
         VariousPropertyBits=   746604571
         Size            =   "1958;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   315
         Left            =   6750
         TabIndex        =   211
         Top             =   270
         Width           =   885
         VariousPropertyBits=   746604571
         Size            =   "1561;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   315
         Left            =   4230
         TabIndex        =   210
         Top             =   270
         Width           =   2505
         VariousPropertyBits=   746604571
         Size            =   "4419;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertySearch 
         Height          =   315
         Left            =   1035
         TabIndex        =   32
         Top             =   270
         Width           =   3225
         VariousPropertyBits=   746604571
         Size            =   "5689;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label5 
         Height          =   195
         Left            =   7650
         TabIndex        =   209
         Top             =   45
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Total Area"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   195
         Left            =   6705
         TabIndex        =   208
         Top             =   45
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Post Code"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label3 
         Height          =   195
         Left            =   4365
         TabIndex        =   207
         Top             =   45
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Address"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   45
         TabIndex        =   206
         Top             =   45
         Width           =   960
         VariousPropertyBits=   8388627
         Caption         =   "Code"
         Size            =   "1693;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1095
         TabIndex        =   205
         Top             =   45
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Name"
         Size            =   "2090;344"
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
         Height          =   285
         Index           =   15
         Left            =   15
         Top             =   0
         Width           =   8640
      End
      Begin MSForms.TextBox txtSearchProperty 
         Height          =   315
         Left            =   0
         TabIndex        =   31
         Top             =   270
         Width           =   1020
         VariousPropertyBits=   746604571
         Size            =   "1799;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame fmeProperty 
      Caption         =   "Property Information"
      ForeColor       =   &H00000000&
      Height          =   3225
      Left            =   120
      TabIndex        =   17
      Top             =   30
      Width           =   12375
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Property"
         Height          =   345
         Left            =   5940
         TabIndex        =   15
         Top             =   2700
         Width           =   1365
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
         Left            =   4095
         TabIndex        =   0
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmdUploadImageAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11760
         TabIndex        =   35
         ToolTipText     =   "Add new image"
         Top             =   2820
         Width           =   555
      End
      Begin VB.CommandButton cmdImgDelete 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11760
         TabIndex        =   34
         ToolTipText     =   "Delete current image"
         Top             =   2520
         Width           =   555
      End
      Begin VB.CommandButton cmdNewProperty 
         Caption         =   "&New Property"
         Height          =   345
         Left            =   720
         TabIndex        =   12
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelProperty 
         Caption         =   "&Cancel Property"
         Height          =   345
         Left            =   2025
         TabIndex        =   13
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton cmdSaveProperty 
         Caption         =   "&Save Property"
         Height          =   345
         Left            =   3345
         TabIndex        =   11
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton cmdEditProperty 
         Caption         =   "&Edit Property"
         Height          =   345
         Left            =   4650
         TabIndex        =   14
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton cmdCloseProperty 
         Caption         =   "C&lose"
         Height          =   345
         Left            =   7320
         TabIndex        =   16
         Top             =   2700
         Width           =   1275
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1575
         TabIndex        =   213
         Top             =   360
         Width           =   2835
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "5001;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Details:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   68
         Top             =   1800
         Width           =   1155
      End
      Begin MSForms.TextBox txtContactDetails 
         Height          =   795
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Width           =   2835
         VariousPropertyBits=   -1467987941
         MaxLength       =   200
         Size            =   "5001;1402"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Manager/Contact:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   67
         Top             =   1440
         Width           =   1395
      End
      Begin MSForms.TextBox txtManager 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   2835
         VariousPropertyBits=   746604571
         MaxLength       =   80
         Size            =   "5001;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdPropertyLookup 
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   780
         Width           =   300
         Caption         =   """"
         Size            =   "529;450"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtPropertyID 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   750
         Width           =   2850
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         MaxLength       =   10
         Size            =   "5027;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyName 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1110
         Width           =   2835
         VariousPropertyBits=   746604571
         MaxLength       =   100
         Size            =   "5001;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblImageName 
         Height          =   195
         Left            =   8760
         TabIndex        =   37
         Top             =   120
         Width           =   3600
         VariousPropertyBits=   8388627
         Caption         =   "Image Name:"
         Size            =   "6350;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image imgPropertyPicture 
         BorderStyle     =   1  'Fixed Single
         Height          =   2730
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2850
      End
      Begin MSForms.CommandButton cmdImgLeftMove 
         Height          =   255
         Left            =   11760
         TabIndex        =   36
         ToolTipText     =   "Next image"
         Top             =   2220
         Width           =   555
         PicturePosition =   262148
         Size            =   "979;450"
         Picture         =   "frmProperty2.frx":08CA
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox txtProAddressLine1 
         Height          =   315
         Left            =   5370
         TabIndex        =   6
         Top             =   360
         Width           =   2955
         VariousPropertyBits=   746604571
         MaxLength       =   70
         Size            =   "5212;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProAddressLine2 
         Height          =   315
         Left            =   5370
         TabIndex        =   7
         Top             =   780
         Width           =   2955
         VariousPropertyBits=   746604571
         MaxLength       =   70
         Size            =   "5212;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProAddressLine3 
         Height          =   315
         Left            =   5370
         TabIndex        =   8
         Top             =   1200
         Width           =   2955
         VariousPropertyBits=   746604571
         MaxLength       =   70
         Size            =   "5212;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProAddressLine4 
         Height          =   315
         Left            =   5370
         TabIndex        =   9
         Top             =   1620
         Width           =   2955
         VariousPropertyBits=   746604571
         MaxLength       =   70
         Size            =   "5212;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProPostCode 
         Height          =   315
         Left            =   5370
         TabIndex        =   10
         Top             =   2040
         Width           =   1245
         VariousPropertyBits=   746604571
         MaxLength       =   50
         Size            =   "2196;556"
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Name:"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Post Code:"
         Height          =   255
         Left            =   4500
         TabIndex        =   19
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Left            =   4500
         TabIndex        =   18
         Top             =   360
         Width           =   1035
      End
   End
   Begin TabDlg.SSTab tabProperty 
      Height          =   5295
      Left            =   90
      TabIndex        =   29
      Top             =   3420
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   9
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
      TabCaption(0)   =   "&Property Analysis"
      TabPicture(0)   =   "frmProperty2.frx":11A4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Maintenance History"
      TabPicture(1)   =   "frmProperty2.frx":11C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLeaseHeading"
      Tab(1).Control(1)=   "Frame1(0)"
      Tab(1).Control(2)=   "fraJS_PO"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&Utilities"
      TabPicture(2)   =   "frmProperty2.frx":11DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label82(10)"
      Tab(2).Control(1)=   "Label82(5)"
      Tab(2).Control(2)=   "Label82(4)"
      Tab(2).Control(3)=   "Label82(6)"
      Tab(2).Control(4)=   "Label82(9)"
      Tab(2).Control(5)=   "Label82(3)"
      Tab(2).Control(6)=   "Label82(8)"
      Tab(2).Control(7)=   "Label82(1)"
      Tab(2).Control(8)=   "Label82(2)"
      Tab(2).Control(9)=   "Label82(7)"
      Tab(2).Control(10)=   "Label82(0)"
      Tab(2).Control(11)=   "cboUnitUtilityStatus"
      Tab(2).Control(12)=   "cboAuthority_Supplier"
      Tab(2).Control(13)=   "cboUtilitiesType"
      Tab(2).Control(14)=   "gridUtilities"
      Tab(2).Control(15)=   "txtUnitUtilityCom"
      Tab(2).Control(16)=   "txtUnitUtilityStDt"
      Tab(2).Control(17)=   "cmdUnitStatus"
      Tab(2).Control(18)=   "txtDateVacated"
      Tab(2).Control(19)=   "txtFinalReading"
      Tab(2).Control(20)=   "txtUnitUtilityIniReading"
      Tab(2).Control(21)=   "txtUnitUtilitiesID"
      Tab(2).Control(22)=   "cmdUtilitiesNew"
      Tab(2).Control(23)=   "cmdUtilitiesEdit"
      Tab(2).Control(24)=   "cmdUtilitiesCancel"
      Tab(2).Control(25)=   "cmdUtilitiesSave"
      Tab(2).Control(26)=   "txtUtilitiesReference"
      Tab(2).Control(27)=   "txtChargeRate"
      Tab(2).Control(28)=   "cmdSetUtilitiesType"
      Tab(2).ControlCount=   29
      TabCaption(3)   =   "&Insurance"
      TabPicture(3)   =   "frmProperty2.frx":11F8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "gridInsurance"
      Tab(3).Control(1)=   "fraInsurance"
      Tab(3).Control(2)=   "cmdInsuranceNew"
      Tab(3).Control(3)=   "cmdInsuranceEdit"
      Tab(3).Control(4)=   "cmdInsuranceCancel"
      Tab(3).Control(5)=   "cmdInsuranceSave"
      Tab(3).Control(6)=   "txtPropertyInsuranceID"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "&Health && Safety"
      TabPicture(4)   =   "frmProperty2.frx":1214
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Landlor&d"
      TabPicture(5)   =   "frmProperty2.frx":1230
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdSupplier"
      Tab(5).Control(1)=   "Frame1(2)"
      Tab(5).Control(2)=   "cmdAddLandlord"
      Tab(5).Control(3)=   "cmdSaveLandlord"
      Tab(5).Control(4)=   "cmdDeleteLandlord"
      Tab(5).Control(5)=   "flxLandlordGrid"
      Tab(5).Control(6)=   "Frame2"
      Tab(5).Control(7)=   "txtLLID"
      Tab(5).Control(8)=   "txtLLName"
      Tab(5).Control(9)=   "Label1(18)"
      Tab(5).Control(10)=   "lblLandlord(0)"
      Tab(5).Control(11)=   "lblLandlord(1)"
      Tab(5).Control(12)=   "lblLandlord(2)"
      Tab(5).Control(13)=   "lblLandlord(3)"
      Tab(5).Control(14)=   "lblLandlord(4)"
      Tab(5).Control(15)=   "lblLandlord(5)"
      Tab(5).Control(16)=   "lblLandlord(6)"
      Tab(5).Control(17)=   "lblLandlord(7)"
      Tab(5).ControlCount=   18
      TabCaption(6)   =   "Memo && A&ttachment"
      TabPicture(6)   =   "frmProperty2.frx":124C
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame1(5)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.CommandButton cmdSupplier 
         Caption         =   ".."
         Height          =   315
         Left            =   -68610
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   540
         Width           =   285
      End
      Begin VB.Frame Frame1 
         Caption         =   "LandLord Address:"
         Enabled         =   0   'False
         Height          =   3975
         Index           =   2
         Left            =   -74685
         TabIndex        =   223
         Top             =   4500
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtSupplierAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   233
            Top             =   573
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   232
            Top             =   1572
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   231
            Top             =   906
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   230
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   229
            Top             =   3090
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   228
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   227
            Top             =   2700
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   226
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            TabIndex        =   225
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   224
            Top             =   1239
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   210
            Index           =   15
            Left            =   120
            TabIndex        =   240
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   210
            Index           =   9
            Left            =   120
            TabIndex        =   239
            Top             =   1575
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   210
            Index           =   10
            Left            =   120
            TabIndex        =   238
            Top             =   2310
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Web:"
            Height          =   210
            Index           =   13
            Left            =   120
            TabIndex        =   237
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   236
            Top             =   2700
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            Height          =   210
            Index           =   14
            Left            =   120
            TabIndex        =   235
            Top             =   3090
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   234
            Top             =   1920
            Width           =   750
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Property Analysis Information"
         Height          =   4875
         Left            =   -75000
         TabIndex        =   177
         Top             =   360
         Width           =   12345
         Begin VB.CommandButton cmdAnalysis 
            Caption         =   "..."
            Height          =   315
            Left            =   2100
            TabIndex        =   191
            Top             =   525
            Width           =   255
         End
         Begin VB.CommandButton cmdAnalysisNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   6450
            TabIndex        =   200
            Top             =   4410
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   7530
            TabIndex        =   201
            Top             =   4410
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisCancel 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   10770
            TabIndex        =   204
            Top             =   4410
            Width           =   975
         End
         Begin VB.CommandButton cmdAnalysisSave 
            Caption         =   "&Save"
            Height          =   315
            Left            =   8610
            TabIndex        =   202
            Top             =   4410
            Width           =   975
         End
         Begin VB.TextBox txtPropertyAnalysisID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11835
            TabIndex        =   178
            Top             =   135
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdDeleteAnalysis 
            Caption         =   "De&lete"
            Height          =   315
            Left            =   9690
            TabIndex        =   203
            Top             =   4410
            Width           =   975
         End
         Begin VB.TextBox txtAnalysisDescription 
            Height          =   1440
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   199
            Top             =   2880
            Width           =   12045
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridPropertyAnalysis 
            Height          =   1665
            Left            =   180
            TabIndex        =   198
            Top             =   900
            Width           =   11985
            _ExtentX        =   21140
            _ExtentY        =   2937
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
         Begin MSForms.TextBox txtAnalysisValue1 
            Height          =   330
            Left            =   7695
            TabIndex        =   196
            Top             =   525
            Width           =   1260
            VariousPropertyBits=   679495707
            MaxLength       =   10
            Size            =   "2222;582"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cboAnalysisOption 
            Height          =   315
            Left            =   2385
            TabIndex        =   192
            Top             =   540
            Width           =   1365
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "2408;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtAnalysisPercentage 
            Height          =   330
            Left            =   6345
            TabIndex        =   195
            Top             =   525
            Width           =   1335
            VariousPropertyBits=   679495707
            MaxLength       =   5
            Size            =   "2355;582"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtAnalysisQuantity 
            Height          =   315
            Left            =   5055
            TabIndex        =   194
            Top             =   525
            Width           =   1275
            VariousPropertyBits=   679495707
            MaxLength       =   4
            Size            =   "2249;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox txtAnalysisValue 
            Height          =   315
            Left            =   3765
            TabIndex        =   193
            Top             =   525
            Width           =   1275
            VariousPropertyBits=   679495707
            MaxLength       =   4
            Size            =   "2249;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.ComboBox cboAnalysisType 
            Height          =   315
            Left            =   210
            TabIndex        =   190
            Top             =   525
            Width           =   1845
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "3254;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            Caption         =   "Total Area:"
            Height          =   255
            Left            =   210
            TabIndex        =   189
            Top             =   4470
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Description:"
            Height          =   255
            Index           =   1
            Left            =   11250
            TabIndex        =   188
            Top             =   180
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Analysis Type:"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   187
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Select Option:"
            Height          =   255
            Index           =   2
            Left            =   2385
            TabIndex        =   186
            Top             =   315
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Area            (sq ft/ sq m):"
            Height          =   435
            Index           =   3
            Left            =   3810
            TabIndex        =   185
            Top             =   105
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Quantity:"
            Height          =   255
            Index           =   4
            Left            =   5115
            TabIndex        =   184
            Top             =   285
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Percentage: (%)"
            Height          =   255
            Index           =   5
            Left            =   6375
            TabIndex        =   183
            Top             =   285
            Width           =   1155
         End
         Begin MSForms.TextBox txtAnalysisTotalArea 
            Height          =   315
            Left            =   1200
            TabIndex        =   182
            Top             =   4440
            Width           =   1515
            VariousPropertyBits=   746604571
            Size            =   "2672;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Caption         =   "Value  (£)         (sq ft/ sq m):"
            Height          =   435
            Index           =   6
            Left            =   7695
            TabIndex        =   181
            Top             =   135
            Width           =   855
         End
         Begin MSForms.TextBox txtAnalysisReference 
            Height          =   330
            Left            =   8955
            TabIndex        =   197
            Top             =   520
            Width           =   3165
            VariousPropertyBits=   679495707
            MaxLength       =   33
            Size            =   "5583;582"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Caption         =   "Reference"
            Height          =   300
            Index           =   7
            Left            =   8955
            TabIndex        =   180
            Top             =   225
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Description:"
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   179
            Top             =   2610
            Width           =   1020
         End
      End
      Begin VB.Frame fraJS_PO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   -65520
         TabIndex        =   173
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
         Begin VB.CommandButton cmdQuoteReq 
            Caption         =   "Job Quote Req"
            Height          =   300
            Left            =   60
            TabIndex        =   176
            Top             =   780
            Width           =   1335
         End
         Begin VB.CommandButton cmdAsJS 
            Caption         =   "Job Sheet"
            Height          =   300
            Left            =   60
            TabIndex        =   175
            Top             =   60
            Width           =   1335
         End
         Begin VB.CommandButton cmdAsPO 
            Caption         =   "Job P/Order"
            Height          =   300
            Left            =   60
            TabIndex        =   174
            Top             =   420
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPropertyInsuranceID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74160
         TabIndex        =   161
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdInsuranceSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -64545
         TabIndex        =   146
         Top             =   4035
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -63630
         TabIndex        =   147
         Top             =   4035
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -65460
         TabIndex        =   149
         Top             =   4035
         Width           =   855
      End
      Begin VB.CommandButton cmdInsuranceNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -66375
         TabIndex        =   148
         Top             =   4035
         Width           =   855
      End
      Begin VB.Frame fraInsurance 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   975
         Left            =   -74880
         TabIndex        =   131
         Top             =   360
         Width           =   12255
         Begin VB.CommandButton cmdUsage 
            Caption         =   "..."
            Height          =   315
            Left            =   10080
            TabIndex        =   143
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtTelephone 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7965
            ScrollBars      =   2  'Vertical
            TabIndex        =   141
            Top             =   555
            Width           =   1000
         End
         Begin VB.CommandButton cmdSetInsurer 
            Caption         =   "..."
            Height          =   315
            Left            =   1320
            TabIndex        =   133
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   10305
            ScrollBars      =   2  'Vertical
            TabIndex        =   144
            Top             =   555
            Width           =   1300
         End
         Begin VB.CommandButton cmdSetInsuranceType 
            Caption         =   "..."
            Height          =   315
            Left            =   2900
            TabIndex        =   135
            Top             =   555
            Width           =   255
         End
         Begin VB.TextBox txtPolicyNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3150
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   136
            Top             =   555
            Width           =   1140
         End
         Begin VB.TextBox txtSumInsured 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4290
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   555
            Width           =   990
         End
         Begin VB.TextBox txtAnnualPR 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5265
            ScrollBars      =   2  'Vertical
            TabIndex        =   138
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtStartDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6165
            ScrollBars      =   2  'Vertical
            TabIndex        =   139
            Top             =   555
            Width           =   900
         End
         Begin VB.TextBox txtExpiryDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7065
            ScrollBars      =   2  'Vertical
            TabIndex        =   140
            Top             =   555
            Width           =   900
         End
         Begin VB.CommandButton cmdUtilitiesAttach 
            Caption         =   "::"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11620
            TabIndex        =   145
            Top             =   555
            Width           =   255
         End
         Begin MSDataListLib.DataCombo cboInsurer 
            Bindings        =   "frmProperty2.frx":1268
            DataSource      =   "adoInsurer"
            Height          =   315
            Left            =   120
            TabIndex        =   132
            Top             =   555
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboInsuranceType 
            Bindings        =   "frmProperty2.frx":1281
            DataSource      =   "adoInsuranceType"
            Height          =   315
            Left            =   1575
            TabIndex        =   134
            Top             =   555
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboUsage 
            Bindings        =   "frmProperty2.frx":12A0
            DataSource      =   "adoInsUsage"
            Height          =   315
            Left            =   8970
            TabIndex        =   142
            Top             =   555
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Usage"
            Height          =   195
            Index           =   8
            Left            =   8970
            TabIndex        =   160
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Tel"
            Height          =   195
            Index           =   7
            Left            =   7965
            TabIndex        =   159
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Insurance Type"
            Height          =   195
            Index           =   1
            Left            =   1575
            TabIndex        =   158
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Attach."
            Height          =   195
            Index           =   10
            Left            =   11620
            TabIndex        =   157
            Top             =   315
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Comment"
            Height          =   195
            Index           =   9
            Left            =   10305
            TabIndex        =   156
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Insurer"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   155
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Policy No"
            Height          =   195
            Index           =   2
            Left            =   3150
            TabIndex        =   154
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "Sum Insured"
            Height          =   495
            Index           =   3
            Left            =   4290
            TabIndex        =   153
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Annual PR"
            Height          =   495
            Index           =   4
            Left            =   5265
            TabIndex        =   152
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Start Date"
            Height          =   435
            Index           =   5
            Left            =   6165
            TabIndex        =   151
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Expiry Date"
            Height          =   495
            Index           =   6
            Left            =   7065
            TabIndex        =   150
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdSetUtilitiesType 
         Caption         =   "..."
         Height          =   315
         Left            =   -72825
         TabIndex        =   103
         Top             =   720
         Width           =   220
      End
      Begin VB.TextBox txtChargeRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66660
         TabIndex        =   110
         Top             =   720
         Width           =   780
      End
      Begin VB.TextBox txtUtilitiesReference 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70965
         TabIndex        =   105
         Top             =   720
         Width           =   1050
      End
      Begin VB.CommandButton cmdUtilitiesSave 
         Caption         =   "&Save"
         Height          =   315
         Left            =   -65085
         TabIndex        =   118
         Top             =   4020
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesCancel 
         Caption         =   "&Cancel"
         Height          =   315
         Left            =   -64020
         TabIndex        =   117
         Top             =   4020
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesEdit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   -66090
         TabIndex        =   116
         Top             =   4020
         Width           =   975
      End
      Begin VB.CommandButton cmdUtilitiesNew 
         Caption         =   "&New"
         Height          =   315
         Left            =   -67125
         TabIndex        =   115
         Top             =   4020
         Width           =   975
      End
      Begin VB.TextBox txtUnitUtilitiesID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -63165
         TabIndex        =   114
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtUnitUtilityIniReading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -65925
         TabIndex        =   111
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtFinalReading 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -65040
         TabIndex        =   112
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtDateVacated 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67590
         TabIndex        =   109
         Top             =   720
         Width           =   945
      End
      Begin VB.CommandButton cmdUnitStatus 
         Caption         =   "..."
         Height          =   315
         Left            =   -68730
         TabIndex        =   107
         Top             =   720
         Width           =   220
      End
      Begin VB.TextBox txtUnitUtilityStDt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68520
         TabIndex        =   108
         Top             =   720
         Width           =   945
      End
      Begin VB.TextBox txtUnitUtilityCom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -64170
         TabIndex        =   113
         Top             =   720
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Notes"
         Height          =   3975
         Index           =   5
         Left            =   120
         TabIndex        =   97
         Top             =   360
         Width           =   12195
         Begin VB.TextBox txtMemo 
            Height          =   2475
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   166
            Top             =   240
            Width           =   11775
         End
         Begin VB.CommandButton cmdMemoCancel 
            Caption         =   "&Cancel"
            Height          =   300
            Left            =   7800
            TabIndex        =   165
            Top             =   3555
            Width           =   975
         End
         Begin VB.CommandButton cmdMemoSave 
            Caption         =   "&Save"
            Height          =   300
            Left            =   5400
            TabIndex        =   164
            Top             =   3555
            Width           =   975
         End
         Begin VB.CommandButton cmdMemoEdit 
            Caption         =   "&Edit"
            Height          =   300
            Left            =   3000
            TabIndex        =   163
            Top             =   3555
            Width           =   975
         End
         Begin VB.Frame Frame17 
            Caption         =   "Attactment Files:"
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
            Height          =   735
            Left            =   240
            TabIndex        =   98
            Top             =   2760
            Width           =   11775
            Begin VB.CommandButton cmdDeleteFile 
               Caption         =   "&Delete File"
               Height          =   375
               Left            =   10080
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   240
               Width           =   1350
            End
            Begin VB.CommandButton cmdClinetAddAtch 
               Caption         =   "&Add New"
               Height          =   375
               Left            =   6840
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   240
               Width           =   1350
            End
            Begin VB.CommandButton cmdOpenFile 
               Caption         =   "&Open File"
               Height          =   375
               Left            =   8520
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   240
               Width           =   1350
            End
            Begin MSForms.ComboBox cmbFiles 
               Height          =   285
               Left            =   120
               TabIndex        =   167
               Top             =   240
               Width           =   4890
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "8625;503"
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1763;4233"
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Property Maintenance"
         Height          =   3885
         Index           =   0
         Left            =   -74880
         TabIndex        =   51
         Top             =   360
         Width           =   12225
         Begin VB.CommandButton cmdEmailJS_PO 
            Caption         =   "Email"
            Height          =   355
            Left            =   9240
            TabIndex        =   172
            Top             =   3360
            Width           =   1275
         End
         Begin VB.Frame Frame1 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   168
            Top             =   3255
            Width           =   3375
            Begin VB.OptionButton optAll 
               Caption         =   "View All"
               Height          =   255
               Left            =   120
               TabIndex        =   171
               Top             =   160
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optJobs 
               Caption         =   "View Jobs"
               Height          =   255
               Left            =   1080
               TabIndex        =   170
               Top             =   160
               Width           =   1095
            End
            Begin VB.OptionButton optDiary 
               Caption         =   "View Diary"
               Height          =   255
               Left            =   2160
               TabIndex        =   169
               Top             =   160
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdNewMHistory 
            Caption         =   "View Job"
            Height          =   355
            Left            =   3600
            TabIndex        =   55
            Top             =   3360
            Width           =   1395
         End
         Begin VB.CommandButton cmdEditMHistory 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   355
            Left            =   7680
            TabIndex        =   54
            Top             =   3360
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton cmdPrintJobSheet 
            Caption         =   "Print"
            Height          =   355
            Left            =   10800
            TabIndex        =   53
            Top             =   3360
            Width           =   1275
         End
         Begin VB.CommandButton cmdAddDiary 
            Caption         =   "View &Diary Entry"
            Height          =   355
            Left            =   5160
            TabIndex        =   52
            Top             =   3360
            Visible         =   0   'False
            Width           =   1395
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridMaintenanceHistory 
            Height          =   2565
            Left            =   120
            TabIndex        =   56
            Top             =   690
            Width           =   11955
            _ExtentX        =   21087
            _ExtentY        =   4524
            _Version        =   393216
            Cols            =   10
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
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Budget / Location"
            Height          =   435
            Index           =   10
            Left            =   10920
            TabIndex        =   69
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Assigned To"
            Height          =   435
            Index           =   6
            Left            =   7080
            TabIndex        =   66
            Top             =   255
            Width           =   795
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Task Owner"
            Height          =   255
            Index           =   5
            Left            =   5640
            TabIndex        =   65
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Job Item / Diary Entry"
            Height          =   495
            Index           =   4
            Left            =   4440
            TabIndex        =   64
            Top             =   255
            Width           =   975
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Reminder"
            Height          =   435
            Index           =   7
            Left            =   8280
            TabIndex        =   63
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm"
            Height          =   195
            Index           =   8
            Left            =   9600
            TabIndex        =   62
            Top             =   255
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Maintenance Type"
            Height          =   435
            Index           =   1
            Left            =   840
            TabIndex        =   61
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Type"
            Height          =   480
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Reported"
            Height          =   480
            Index           =   2
            Left            =   2145
            TabIndex        =   59
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Ref"
            Height          =   435
            Index           =   3
            Left            =   3000
            TabIndex        =   58
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            Height          =   435
            Index           =   9
            Left            =   9600
            TabIndex        =   57
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdAddLandlord 
         Caption         =   "&Add Landlord"
         Height          =   345
         Left            =   -67890
         TabIndex        =   41
         Top             =   4500
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveLandlord 
         Caption         =   "&Save Landlord"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -66450
         TabIndex        =   40
         Top             =   4500
         Width           =   1395
      End
      Begin VB.CommandButton cmdDeleteLandlord 
         Caption         =   "&Delete Landlord"
         Enabled         =   0   'False
         Height          =   345
         Left            =   -65010
         TabIndex        =   39
         Top             =   4500
         Width           =   1395
      End
      Begin VB.Frame Frame4 
         Caption         =   "Health && Safety"
         Height          =   4425
         Left            =   -74910
         TabIndex        =   30
         Top             =   390
         Width           =   12255
         Begin VB.TextBox txtUnitSafetyID 
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
            Height          =   285
            Left            =   8880
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkCertificate 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10995
            TabIndex        =   81
            Top             =   780
            Width           =   255
         End
         Begin VB.CommandButton cmdSafetyNew 
            Caption         =   "&New"
            Height          =   315
            Left            =   7800
            TabIndex        =   86
            Top             =   3465
            Width           =   975
         End
         Begin VB.CommandButton cmdSafetyEdit 
            Caption         =   "&Edit"
            Height          =   315
            Left            =   8875
            TabIndex        =   85
            Top             =   3465
            Width           =   975
         End
         Begin VB.CommandButton cmdSafetyCancel 
            Caption         =   "&Cancel"
            Height          =   315
            Left            =   11025
            TabIndex        =   84
            Top             =   3465
            Width           =   975
         End
         Begin VB.CommandButton cmdSafetySave 
            Caption         =   "&Save"
            Height          =   315
            Left            =   9950
            TabIndex        =   83
            Top             =   3465
            Width           =   975
         End
         Begin VB.CommandButton cmdSafety 
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
            Height          =   315
            Left            =   1920
            TabIndex        =   71
            Top             =   720
            Width           =   215
         End
         Begin VB.TextBox txtRef 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3345
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   720
            Width           =   1545
         End
         Begin VB.TextBox txtNextDueDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5850
            TabIndex        =   75
            Top             =   720
            Width           =   990
         End
         Begin VB.TextBox txtSafetyTelephone 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8370
            TabIndex        =   78
            Top             =   720
            Width           =   1155
         End
         Begin VB.TextBox txtDateChk 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4875
            TabIndex        =   74
            Top             =   720
            Width           =   990
         End
         Begin VB.TextBox txtComment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9495
            TabIndex        =   79
            Top             =   720
            Width           =   1155
         End
         Begin VB.CheckBox chkAlarm 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   10665
            TabIndex        =   80
            Top             =   780
            Width           =   255
         End
         Begin VB.CommandButton cmdInspectedBy 
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
            Height          =   315
            Left            =   8160
            TabIndex        =   77
            Top             =   720
            Width           =   215
         End
         Begin VB.CommandButton cmdAttachment 
            Caption         =   "::"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11340
            TabIndex        =   82
            Top             =   720
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridSafety 
            Height          =   2355
            Left            =   120
            TabIndex        =   88
            Top             =   1080
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   4154
            _Version        =   393216
            Cols            =   3
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSDataListLib.DataCombo cboSafetyType 
            Bindings        =   "frmProperty2.frx":12BA
            DataSource      =   "adoSafetyType"
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboSchedule 
            Bindings        =   "frmProperty2.frx":12D6
            DataSource      =   "adoSafetyStatus"
            Height          =   315
            Left            =   2145
            TabIndex        =   72
            Top             =   720
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cboInspectedBy 
            Bindings        =   "frmProperty2.frx":12F4
            DataSource      =   "adoInspector"
            Height          =   315
            Left            =   6810
            TabIndex        =   76
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ListField       =   "Value"
            BoundColumn     =   "Code"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label41 
            Caption         =   "Reference"
            Height          =   255
            Index           =   2
            Left            =   3345
            TabIndex        =   96
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label41 
            Caption         =   "Type"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "Next Inspection"
            Height          =   375
            Index           =   4
            Left            =   5850
            TabIndex        =   94
            Top             =   300
            Width           =   990
         End
         Begin VB.Label Label41 
            Caption         =   "Inspected By"
            Height          =   255
            Index           =   5
            Left            =   6840
            TabIndex        =   93
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label41 
            Caption         =   "Comment"
            Height          =   255
            Index           =   7
            Left            =   9495
            TabIndex        =   92
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label41 
            Caption         =   "Schedule"
            Height          =   255
            Index           =   1
            Left            =   2145
            TabIndex        =   91
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label41 
            Caption         =   "Date Checked"
            Height          =   345
            Index           =   3
            Left            =   4875
            TabIndex        =   90
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label41 
            Caption         =   "Contact / Telephone"
            Height          =   375
            Index           =   6
            Left            =   8370
            TabIndex        =   89
            Top             =   300
            Width           =   1095
         End
         Begin VB.Image VerticalLabel 
            Height          =   585
            Index           =   0
            Left            =   10665
            Picture         =   "frmProperty2.frx":130F
            Top             =   105
            Width           =   195
         End
         Begin VB.Image VerticalLabel 
            Height          =   435
            Index           =   1
            Left            =   10995
            Picture         =   "frmProperty2.frx":16F9
            Top             =   255
            Width           =   225
         End
         Begin VB.Image VerticalLabel 
            Height          =   600
            Index           =   2
            Left            =   11340
            Picture         =   "frmProperty2.frx":1A8D
            Top             =   100
            Width           =   210
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxLandlordGrid 
         Height          =   3075
         Left            =   -74730
         TabIndex        =   42
         Top             =   1260
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5424
         _Version        =   393216
         Cols            =   3
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUtilities 
         Height          =   2835
         Left            =   -74880
         TabIndex        =   119
         Top             =   1140
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   3
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSDataListLib.DataCombo cboUtilitiesType 
         Bindings        =   "frmProperty2.frx":1E97
         DataSource      =   "adoUtilitiesType"
         Height          =   315
         Left            =   -74040
         TabIndex        =   102
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboAuthority_Supplier 
         Bindings        =   "frmProperty2.frx":1EB6
         DataSource      =   "adoSupplier"
         Height          =   315
         Left            =   -72630
         TabIndex        =   104
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "SupplierName"
         BoundColumn     =   "SupplierID"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboUnitUtilityStatus 
         Bindings        =   "frmProperty2.frx":1ED0
         DataSource      =   "adoStatus"
         Height          =   315
         Left            =   -69930
         TabIndex        =   106
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "Value"
         BoundColumn     =   "Code"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridInsurance 
         Height          =   2640
         Left            =   -74835
         TabIndex        =   162
         Top             =   1335
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   4657
         _Version        =   393216
         Cols            =   3
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "LandLord Alternative Address :"
         Enabled         =   0   'False
         Height          =   4005
         Left            =   -70320
         TabIndex        =   241
         Top             =   4455
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtSupplierOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   246
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   245
            Top             =   1530
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   244
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   243
            Top             =   1050
            Width           =   2655
         End
         Begin VB.TextBox txtSupplierOfficeAddressLine4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   242
            Top             =   1950
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   248
            Top             =   2520
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   247
            Top             =   600
            Width           =   615
         End
      End
      Begin MSForms.TextBox txtLLID 
         Height          =   315
         Left            =   -72930
         TabIndex        =   252
         Top             =   540
         Width           =   1590
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "2805;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtLLName 
         Height          =   315
         Left            =   -71310
         TabIndex        =   250
         Top             =   540
         Width           =   2670
         VariousPropertyBits=   746604571
         MaxLength       =   10
         BorderStyle     =   1
         Size            =   "4710;556"
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
         Caption         =   "LandLord A/C:"
         Height          =   195
         Index           =   18
         Left            =   -74730
         TabIndex        =   249
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Height          =   435
         Index           =   0
         Left            =   -74880
         TabIndex        =   130
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Index           =   7
         Left            =   -66660
         TabIndex        =   129
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   -72630
         TabIndex        =   128
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   1
         Left            =   -74040
         TabIndex        =   127
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Initial Reading"
         Height          =   435
         Index           =   8
         Left            =   -65925
         TabIndex        =   126
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   195
         Index           =   3
         Left            =   -70965
         TabIndex        =   125
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Final    Reading"
         Height          =   435
         Index           =   9
         Left            =   -65040
         TabIndex        =   124
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "End  Date"
         Height          =   435
         Index           =   6
         Left            =   -67590
         TabIndex        =   123
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   4
         Left            =   -69930
         TabIndex        =   122
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   435
         Index           =   5
         Left            =   -68520
         TabIndex        =   121
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   195
         Index           =   10
         Left            =   -64170
         TabIndex        =   120
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   -74730
         TabIndex        =   50
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Address Line 1"
         Height          =   255
         Index           =   1
         Left            =   -73335
         TabIndex        =   49
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Address Line 2"
         Height          =   255
         Index           =   2
         Left            =   -71655
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Address Line 3"
         Height          =   255
         Index           =   3
         Left            =   -70095
         TabIndex        =   47
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Post Code"
         Height          =   255
         Index           =   4
         Left            =   -68415
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Tel"
         Height          =   255
         Index           =   5
         Left            =   -67335
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Mobile"
         Height          =   255
         Index           =   6
         Left            =   -66255
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLandlord 
         Caption         =   "Account Balance"
         Height          =   255
         Index           =   7
         Left            =   -65055
         TabIndex        =   43
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLeaseHeading 
         Alignment       =   2  'Center
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
         Left            =   -72645
         TabIndex        =   24
         Top             =   660
         Width           =   7875
      End
   End
   Begin MSAdodcLib.Adodc adoSafetyType 
      Height          =   330
      Left            =   360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Safety Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSafetyStatus 
      Height          =   330
      Left            =   360
      Top             =   8400
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Safety Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInspector 
      Height          =   330
      Left            =   360
      Top             =   8040
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Inspector"
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
   Begin MSAdodcLib.Adodc adoUtilitiesType 
      Height          =   330
      Left            =   2520
      Top             =   8040
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Unitlities Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoSupplier 
      Height          =   330
      Left            =   2520
      Top             =   8400
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Suppliers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoStatus 
      Height          =   330
      Left            =   2520
      Top             =   8760
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Status"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsuranceType 
      Height          =   330
      Left            =   4680
      Top             =   8760
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Insurance Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsurer 
      Height          =   330
      Left            =   4680
      Top             =   8040
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Insurer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoInsUsage 
      Height          =   330
      Left            =   4680
      Top             =   8400
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
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
      Caption         =   "Usage"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label31 
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8175
      TabIndex        =   23
      Top             =   2745
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmProperty2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LOAD_PROPERTY_PROPERTYID  As String
Public CLIENT_NAME               As String

Dim NEWMODE_                     As Boolean
Dim SEARCHPropertyMODE_          As Boolean
Dim UNIT_INSURANCE_NEW_ENTRY     As Boolean
Dim M_HISTORY_NEW_ENTRY_         As Boolean
Dim Property_ANALYSIS_NEW_ENTRY  As Boolean
Dim HEALTH_SAFETY_NEW_ENTRY      As Boolean
Dim Property_INSURANCE_NEW_ENTRY As Boolean
Dim Property_UTILITIES_NEW_ENTRY As Boolean
Dim IMAGE_FILE_NAME_             As String
Dim szEditingPropID              As String

Dim DSN_ALARM_                   As String
Dim lblAsJS_PO                   As String
'Private HEALTH_SAFETY_ID As String
Public HEALTH_N_SAFETY_ATTACH    As Boolean
Dim UNIT_UTILITIES_NEW_ENTRY     As Boolean
Private INSURANCE_ID             As String
Dim sTextBox As String
Dim szaSupplierBalance() As String
Private Sub cboLandlord_Change()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    Call loadlandlorddetails(txtLLID.text, adoConn)
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub cboLandlord_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    
    
End Sub
Private Sub loadlandlorddetails(Id As String, adoConn As ADODB.Connection)
    Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   Dim sSQLQuery_ As String, sFilter As String
   If Id = "" Then Exit Sub
   'MousePointer = vbHourglass

   sSQLQuery_ = "SELECT S.SupplierID, S.SupplierName, S.SupplierAddressLine1, S.SupplierAddressLine2, " & _
                  "S.SupplierAddressLine3,S.SupplierAddressLine4, S.SupplierPostCode, " & _
                  "S.SupplierOfficeEmail, S.SupplierPersonalEmail, V.VAT_RATE, " & _
                  "S.SupplierHomeTel, S.SupplierMobile, S.SupplierOfficeAddressLine1, " & _
                  "S.SupplierOfficeAddressLine2, S.SupplierOfficeAddressLine3,S.SupplierOfficeAddressLine4, " & _
                  "S.SupplierOfficePostCode, S.PLControl, S.PLControlName, " & _
                  "S.SupplierOfficeTel, S.SupplierMemo, S.VATReg,S.VATCode, " & _
                  "S.BacsRef, S.SupplierType, S.creditlimit, S.nominalcode, " & _
                  "S.AccountType, S.PaymentType, S.PaymentTerms, S.SortCode, S.AcNo, S.AcName, S.BPR, " & _
                  "N.Name AS NN " & _
                "FROM (Supplier AS S LEFT OUTER JOIN NominalLedger AS N " & _
                  "ON S.NominalCode = N.Code) LEFT OUTER JOIN tlbVatCode AS V " & _
                  "ON S.VATCode = V.VAT_CODE " & _
                "WHERE S.SupplierID = '" & Id & "';"
'Debug.Print sSQLQuery_
   rstRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly
   If Not rstRst.EOF Then
'        txtLLID.text = rstRst!SupplierName
        txtSupplierAddressLine1.text = IIf(IsNull(rstRst!SupplierAddressLine1), "", rstRst!SupplierAddressLine1)
        txtSupplierAddressLine2.text = IIf(IsNull(rstRst!SupplierAddressLine2), "", rstRst!SupplierAddressLine2)
        txtSupplierAddressLine3.text = IIf(IsNull(rstRst!SupplierAddressLine3), "", rstRst!SupplierAddressLine3)
        txtSupplierAddressLine4.text = IIf(IsNull(rstRst!SupplierAddressLine4), "", rstRst!SupplierAddressLine4)
        txtSupplierPostCode.text = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
        txtSupplierHomeTel.text = IIf(IsNull(rstRst!SupplierHomeTel), "", rstRst!SupplierHomeTel)
        txtSupplierMobile.text = IIf(IsNull(rstRst!SupplierMobile), "", rstRst!SupplierMobile)
        txtSupplierOfficeTel.text = IIf(IsNull(rstRst!SupplierOfficeTel), "", rstRst!SupplierOfficeTel)
        txtSupplierOfficeEmail.text = IIf(IsNull(rstRst!SupplierOfficeEmail), "", rstRst!SupplierOfficeEmail)
        txtSupplierPersonalEmail.text = IIf(IsNull(rstRst!SupplierPersonalEmail), "", rstRst!SupplierPersonalEmail)
        
        txtSupplierOfficeAddressLine1.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine1), "", rstRst!SupplierOfficeAddressLine1)
        txtSupplierOfficeAddressLine2.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine2), "", rstRst!SupplierOfficeAddressLine2)
        txtSupplierOfficeAddressLine3.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine3), "", rstRst!SupplierOfficeAddressLine3)
        txtSupplierOfficeAddressLine4.text = IIf(IsNull(rstRst!SupplierOfficeAddressLine4), "", rstRst!SupplierOfficeAddressLine4)
        txtSupplierOfficePostCode.text = IIf(IsNull(rstRst!SupplierOfficePostCode), "", rstRst!SupplierOfficePostCode)
   End If
   rstRst.Close
   Set rstRst = Nothing
End Sub
Private Sub cmdDelete_Click()
    If txtPropertyID.text = "" Then
      ShowMsgInTaskBar "Please select a Property to continue."
      Exit Sub
   End If
   'Check the existence of Unit under it or any transaction is made
   Dim adoConn As New ADODB.Connection
   Dim rsTransaction As New ADODB.Recordset
   adoConn.Open getConnectionString
   rsTransaction.Open "Select * from Property where PropertyID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If rsTransaction.EOF Then
        MsgBox "This property ID was not found in the database", vbInformation, "Not found"
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   rsTransaction.Open "Select UnitID from tlbpayment where unitID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Property cannot be deleted, because there are transactions entered against it.", vbInformation, "Cannot Delete"
        cmdCancelProperty_Click
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   rsTransaction.Open "Select PropertyID from tlbBankPayment where PropertyID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Property cannot be deleted, because there are bank transactions entered against it.", vbInformation, "Cannot Delete"
        cmdCancelProperty_Click
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   
   rsTransaction.Open "Select UnitNumber from Units where PropertyID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Property cannot be deleted, because there are units linked to it. Please remove the linked units and try again", vbInformation, "Cannot Delete"
        cmdCancelProperty_Click
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   rsTransaction.Open "Select PropertyID from DemandTypes where PropertyID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If Not rsTransaction.EOF Then
        MsgBox "This Property cannot be deleted. There is some DemandTypes exists with this property ID", vbInformation, "Cannot Delete"
        cmdCancelProperty_Click
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   'Delete property
    If MsgBox("  Are you sure you wish to delete this property information?" & (Chr(13) + Chr(10)) & _
             "", vbYesNo + vbQuestion, _
             "Delete property information") = vbYes Then
            adoConn.Execute "DELETE FROM Property where PropertyID='" & txtPropertyID.text & "'"
            MsgBox "Delete Successful", vbInformation
   End If
   adoConn.Close
   cmdCancelProperty_Click
   FocusControl cmdPropertyLookup
End Sub

Private Sub cmdGridUnitClose2_Click(Index As Integer)
    picSupList.Visible = False
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
End Sub



Private Sub cmdSupplier_Click()
   Call PrepareList

   picSupList.Top = 3915
   picSupList.Left = 1890
   picSupList.Visible = True
   picSupList.ZOrder 0
   
   txtLLID.Locked = True
   'added by anol 08 Apr 2015
   FocusControl txtSupplierSearchID
End Sub

Private Sub flxSupplierList_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    txtLLID.text = flxSupplierList.TextMatrix(flxSupplierList.row, 1)
    txtLLName.text = flxSupplierList.TextMatrix(flxSupplierList.row, 2)
    Call loadlandlorddetails(txtLLID.text, adoConn)
    picSupList.Visible = False
    adoConn.Close
    Set adoConn = Nothing
End Sub





Private Sub gridPropertyAnalysis_Click()
        txtPropertyAnalysisID.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 0)
        txtAnalysisReference.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 8)
        cboAnalysisType.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 1)
        txtAnalysisDescription.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 2)
        cboAnalysisOption.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 3)
        txtAnalysisValue.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 4)
        txtAnalysisQuantity.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 5)
        txtAnalysisPercentage.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 6)
        txtAnalysisValue1.text = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.row, 7)
        If cmdSaveProperty.Enabled = True Then
            cmdAnalysisEdit.Enabled = True
            cmdDeleteAnalysis.Enabled = True
            cmdAnalysisCancel.Enabled = True
            cmdAnalysisNew.Enabled = False
            cmdAnalysisSave.Enabled = False
        End If
End Sub

Private Sub tabProperty_Click(PreviousTab As Integer)

   Dim adoConn As New ADODB.Connection
'   Dim rslandlord As New ADODB.Recordset
'   Dim sSQLQuery As String
'   If tabProperty.Tab = 5 Then
       adoConn.Open getConnectionString
'
'       sSQLQuery = "SELECT Landlord, SupplierName " & _
'                    "FROM Landlord L Left join supplier S on L.landlord=S.SupplierID where PropertyID='" & txtPropertyID.text & "';"
'
'       rslandlord.Open sSQLQuery, adoConn, adOpenStatic, adLockReadOnly
'       If Not rslandlord.EOF Then
'            txtLLID.text = rslandlord("Landlord").Value
'            txtLLName.text = rslandlord("SupplierName").Value
        If tabProperty.Tab = 5 Then
            Call loadlandlorddetails(txtLLID.text, adoConn)
       End If
       adoConn.Close
       Set adoConn = Nothing
'
'    End If
    Select Case PreviousTab
          Case 5:           'Rent Charges
            If cmdSaveLandlord.Enabled Then
               If MsgBox("Do you what to save changes to this Landlord?", vbQuestion + vbYesNo, "Landlord") = vbYes Then
                    tabProperty.Tab = PreviousTab
                    FocusControl cmdSaveLandlord
               Else
                    cmdSaveLandlord.Enabled = False
                    cmdAddLandlord.Enabled = True
                    FocusControl cmdAddLandlord
               End If
            End If
    End Select
End Sub

Private Sub txtAnalysisValue1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtAnalysisValue1, KA, 4
   KeyAscii = KA
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdClientList.SetFocus
    End If
End Sub

Private Sub txtPropertyID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        If txtPropertyName.Enabled Then txtPropertyName.SetFocus
    End If
End Sub

Private Sub txtPropertyName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        If txtManager.Enabled Then txtManager.SetFocus
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
Private Sub flxClient_Click()
            picClient.Visible = False
            If sTextBox = "1" Then
                    
                    txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                    cboClientID_Change
                    If NEWMODE_ = True Then
                        If txtPropertyName.Enabled Then txtPropertyName.SetFocus
                    Else
                        If cmdPropertyLookup.Enabled Then cmdPropertyLookup.SetFocus
                    End If
            End If
    
        
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub
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
           flxClient.AddItem ""
            flxClient.TextMatrix(rRow, 1) = "ALL"
            flxClient.TextMatrix(rRow, 2) = "ALL Clients"
            flxClient.RowHeight(rRow) = 240
            
           rRow = 2
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
Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 1700
    picClient.Top = 390
    picClient.Visible = True
    Call LoadflxClient
    txtSearchClientName.Enabled = True
    
   txtSearchClientID.Enabled = True
    txtSearchClientID.SetFocus
End Sub

Private Sub txtPropertySearch_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        gridPropertyLookup.SetFocus
    End If
End Sub

Private Sub txtSearchProperty_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchProperty.text) > 0 Then
        txtPropertySearch.text = ""
   End If

   For i = gridPropertyLookup.Rows - 1 To 1 Step -1
        gridPropertyLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridPropertyLookup.TextMatrix(i, 0)), UCase(txtSearchProperty.text), vbTextCompare) = 0 Then
              gridPropertyLookup.RowHeight(i) = 0
        End If
        If gridPropertyLookup.RowHeight(i) = 240 Then
              gridPropertyLookup.row = i
        End If
        gridPropertyLookup.col = 1
   Next i
   
End Sub

Private Sub txtSearchProperty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           gridPropertyLookup.SetFocus
    End If
'    If KeyCode = 13 Then
'           txtPropertySearch.SetFocus
'    End If
End Sub



Private Sub txtPropertySearch_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtPropertySearch.text) > 0 Then
        txtSearchProperty.text = ""
   End If

   For i = gridPropertyLookup.Rows - 1 To 1 Step -1
        gridPropertyLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridPropertyLookup.TextMatrix(i, 1)), UCase(txtPropertySearch.text), vbTextCompare) = 0 Then
            gridPropertyLookup.RowHeight(i) = 0
        End If
        If gridPropertyLookup.RowHeight(i) = 240 Then
            gridPropertyLookup.row = i
        End If
        'gridPropertyLookup.col = 1
   Next i
End Sub
Public Function PopulatePropertyLookup(strFilter_ As String)
   Dim conProperty_  As New ADODB.Connection
   Dim rstProperty_  As New ADODB.Recordset
   Dim sSQLQuery_    As String
   Dim iRow          As Integer

   'On Error Resume Next
   conProperty_.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   If txtClientList.text <> "" Then
      sSQLQuery_ = "SELECT PropertyID, PropertyNAME, " & _
                     "iif(isnull(ProAddressLine1),'',ProAddressLine1) + ' ' + iif(isnull(ProAddressLine2),'',ProAddressLine2) + ' ' +  iif(isnull(ProAddressLine3),'',ProAddressLine3) + ' ' + iif(isnull(ProAddressLine4),'',ProAddressLine4) as Address, " & _
                     "ProPOSTCODE , TotalArea ,Property.ClientID, Client.ClientName " & _
                   "From Property,Client " & _
                   "WHERE Client.ClientID = Property.ClientID AND Property.ClientID = '" & txtClientList.Tag & "';"
   Else
      sSQLQuery_ = "SELECT PropertyID, PropertyNAME, " & _
                     "iif(isnull(ProAddressLine1),'',ProAddressLine1) + ' ' + iif(isnull(ProAddressLine2),'',ProAddressLine2) + ' ' +  iif(isnull(ProAddressLine3),'',ProAddressLine3) + ' ' + iif(isnull(ProAddressLine4),'',ProAddressLine4) as Address, " & _
                     "ProPOSTCODE , TotalArea, Property.ClientID, Client.ClientName " & _
                   "From Property,Client WHERE Client.ClientID = Property.ClientID;"
   End If
   If txtClientList.Tag = "ALL" Then
        sSQLQuery_ = "SELECT PropertyID, PropertyNAME, " & _
                     "iif(isnull(ProAddressLine1),'',ProAddressLine1) + ' ' + iif(isnull(ProAddressLine2),'',ProAddressLine2) + ' ' +  iif(isnull(ProAddressLine3),'',ProAddressLine3) + ' ' + iif(isnull(ProAddressLine4),'',ProAddressLine4) as Address, " & _
                     "ProPOSTCODE , TotalArea, Property.ClientID, Client.ClientName " & _
                   "From Property,Client WHERE Client.ClientID = Property.ClientID;"
   End If
   rstProperty_.Open sSQLQuery_, conProperty_, adOpenStatic, adLockReadOnly

   iRow = 1

   gridPropertyLookup.Clear
   gridPropertyLookup.Rows = 2
   'gridPropertyLookup.Cols = 5
   gridPropertyLookup.Cols = 7
   ConfigurFlexGrid
    gridPropertyLookup.ColWidth(5) = 0
   gridPropertyLookup.ColWidth(6) = 0
   lblClientName.Caption = "Property Name"
   lblClientID.Caption = "Property ID"
   'On Error Resume Next
    While Not rstProperty_.EOF
         'fixed by anol 20160916
         Call TotalareofProperty(IIf(IsNull(rstProperty_!propertyID), "", rstProperty_!propertyID), conProperty_)
        rstProperty_.MoveNext
    Wend
    rstProperty_.Requery
    If rstProperty_.RecordCount > 0 Then
        rstProperty_.MoveFirst
    End If
   While Not rstProperty_.EOF
        gridPropertyLookup.col = 0
        gridPropertyLookup.row = 1
        gridPropertyLookup.TextMatrix(iRow, 0) = IIf(IsNull(rstProperty_!propertyID), "", rstProperty_!propertyID)
       
        gridPropertyLookup.TextMatrix(iRow, 1) = IIf(IsNull(rstProperty_!PropertyName), "", rstProperty_!PropertyName)
        gridPropertyLookup.TextMatrix(iRow, 2) = IIf(IsNull(rstProperty_!Address), "", rstProperty_!Address)
        gridPropertyLookup.TextMatrix(iRow, 3) = IIf(IsNull(rstProperty_!PROPOSTCODE), "", rstProperty_!PROPOSTCODE)
        gridPropertyLookup.TextMatrix(iRow, 4) = IIf(IsNull(rstProperty_!TotalArea), "", rstProperty_!TotalArea)
        gridPropertyLookup.TextMatrix(iRow, 5) = IIf(IsNull(rstProperty_!ClientID), "", rstProperty_!ClientID)
        gridPropertyLookup.TextMatrix(iRow, 6) = IIf(IsNull(rstProperty_!ClientName), "", rstProperty_!ClientName)
        gridPropertyLookup.RowHeight(iRow) = 280
        rstProperty_.MoveNext
        If Not rstProperty_.EOF Then gridPropertyLookup.AddItem ""
        iRow = iRow + 1
   Wend

   rstProperty_.Close
   conProperty_.Close
   Set rstProperty_ = Nothing
   Set conProperty_ = Nothing
End Function
Private Sub ConfigurFlexGrid()
   fmePropertyLookup.Visible = True
   gridPropertyLookup.Visible = True

   gridPropertyLookup.RowHeight(0) = 350
   gridPropertyLookup.row = 0
'   Dim i As Integer
'   For i = 0 To gridPropertyLookup.Cols - 1
'        gridPropertyLookup.col = i
'        gridPropertyLookup.CellFontBold = True
'   Next i

    gridPropertyLookup.ColWidth(0) = 1100
    gridPropertyLookup.TextMatrix(0, 0) = "Code"
    gridPropertyLookup.ColAlignment(0) = vbLeftJustify
    gridPropertyLookup.ColWidth(1) = 3300 '2100
    gridPropertyLookup.TextMatrix(0, 1) = "Name"
    gridPropertyLookup.ColAlignment(1) = vbLeftJustify
    gridPropertyLookup.ColWidth(2) = 3200 '2500
    gridPropertyLookup.TextMatrix(0, 2) = "Address"
    gridPropertyLookup.ColAlignment(3) = vbLeftJustify
    gridPropertyLookup.ColWidth(3) = 1000
    gridPropertyLookup.TextMatrix(0, 3) = "Post Code"
    gridPropertyLookup.ColAlignment(4) = vbLeftJustify
    gridPropertyLookup.ColWidth(4) = 1000
    gridPropertyLookup.TextMatrix(0, 4) = "Total Area"
    gridPropertyLookup.RowHeight(0) = 0
    gridPropertyLookup.ColAlignment(5) = vbLeftJustify
End Sub

Private Sub cboClientID_Change()
'Resolved by BOSL
'Issue No: 0000467
'Modified By: Asif. 04 Sep 2014
   If txtProAddressLine1.Locked = True Then ' And txtPropertyName.text <> ""
        Dim selectedClient As String
        If Not IsNull(txtClientList.Tag) Then
            selectedClient = txtClientList.Tag
        End If
        'cmdCancelProperty_Click
          ' ComponentEnableModeProperty Me, DefaultMode
          Dim ctrl As Control
          For Each ctrl In Me.Controls
               Select Case TypeName(ctrl)
                   Case "TextBox"
                       If ctrl.Name <> "txtClientList" Then
                            ctrl.Enabled = False
                            ctrl.text = ""
                       End If
                   Case "CheckBox"
                       ctrl.Enabled = False
                   Case "ComboBox"
                       ctrl.Enabled = False
               End Select
           Next ctrl

           Me.Controls("gridPropertyLookup").Visible = False

           Me.cmdNewProperty.Enabled = True
           Me.cmdEditProperty.Enabled = True
           Me.cmdSaveProperty.Enabled = False
           Me.cmdCancelProperty.Enabled = False
           Me.cmdCloseProperty.Enabled = True
           
           
         imgPropertyPicture.Picture = LoadPicture()
         lblImageName.Caption = "Image Name:"
         NEWMODE_ = False
         SEARCHPropertyMODE_ = True
         txtPropertyID.Enabled = True
         txtPropertyName.Enabled = True
         cmdPropertyLookup.Enabled = True
        ' cboClientID.Enabled = True
         If txtPropertyID.text = "" Then
             Exit Sub
         End If
         tabProperty.Enabled = True
        txtClientList.Tag = selectedClient
   End If
'Resolved by BOSL
'Issue No: 0000494
'Modified By: Anol. 08 Dec 2014
   PropertyAnalysisButtonMode DefaultMode
   gridPropertyAnalysis.Clear
   gridPropertyAnalysis.Rows = 2
End Sub

Private Function check_dupplicateLandlord(adoConn As ADODB.Connection) As Boolean
    Dim rsPropertyLandlord As New ADODB.Recordset
    rsPropertyLandlord.Open "Select * from PropertyLandlord where propertyID='" & txtPropertyID.text & "' and landlordID='" & txtLLID.text & "'", adoConn, adOpenKeyset, adLockReadOnly
    If Not rsPropertyLandlord.EOF Then
        check_dupplicateLandlord = True
    End If
End Function
Private Sub cmdAddLandlord_Click()
   cmdSupplier.Enabled = True
   cmdSaveLandlord.Enabled = True
   cmdAddLandlord.Enabled = False
   FocusControl cmdSupplier
'   cmdSupplier.Enabled = True
'   If txtLLID.text = "" Then
'        MsgBox "Please select a landlord to add", vbInformation, "Warning"
'        Exit Sub
'   End If
'
'
'   cmdSaveLandlord.Enabled = True
'   cmdAddLandlord.Enabled = False
'   If MsgBox("Do you want to add a landlord to this property?", vbYesNo, "add landlord") = vbYes Then
'         adoConn.Open getConnectionString
'
'         If check_dupplicateLandlord(adoConn) = False Then
'            rsPropertyLandlord.Open "Select * from PropertyLandlord", adoConn, adOpenDynamic, adLockOptimistic
'            With rsPropertyLandlord
'            .AddNew
'            !propertyID = txtPropertyID.text
'            !landLordID = txtLLID.text
'            .Update
'            End With
'             MsgBox "This landlord added successfully.", vbInformation, "Warning!"
'             loadflxLandlordGrid adoConn
'             adoConn.Close
'             Set adoConn = Nothing
'         Else
'            MsgBox "This landlord already exists in the current list.", vbInformation, "Warning!"
'            Exit Sub
'         End If
'   End If
'    cmdSaveLandlord.Enabled = False
'   cmdAddLandlord.Enabled = True
'
   
   
   'Added by anol 13 May 2015
'   'Lanlord table needs to be updated from Supplier table
'   'Should be alwys syncronized with supplier table where supplier type = 'LL'
'   'So I am deleting all record from landlord table and transferring all where supplier type = 'LL' from supplier table
'   Dim adoConn As New ADODB.Connection
'
'   If adoConn.State = 0 Then
'        adoConn.Open getConnectionString
'   End If
'   adoConn.Execute "DELETE FROM Landlord"
'   Dim rsSupplier As New ADODB.Recordset
'   Dim rsLandlord As New ADODB.Recordset
'   rsSupplier.Open "Select * from Supplier where SupplierType='LL'", adoConn, adOpenStatic, adLockReadOnly
'   rsLandlord.Open "Select * from Landlord", adoConn, adOpenDynamic, adLockOptimistic
'   While Not rsSupplier.EOF
'           rsLandlord.AddNew
'           rsLandlord!landLordID = "L-" & rsSupplier!SupplierID
'           rsLandlord!LandlordName = rsSupplier!SupplierName
'           rsLandlord!LandlordAddressLine1 = rsSupplier!SupplierAddressLine1
'           rsLandlord!LandlordAddressLine2 = rsSupplier!SupplierAddressLine2
'           rsLandlord!LandlordAddressLine3 = rsSupplier!SupplierAddressLine3
'           rsLandlord!LandlordAddressLine4 = rsSupplier!SupplierAddressLine4
'           rsLandlord!LandlordPostCode = rsSupplier!SupplierPostCode
'           rsLandlord!LandlordOfficeEmail = rsSupplier!SupplierOfficeEmail
'           rsLandlord!LandlordPersonalEmail = rsSupplier!SupplierPersonalEmail
'           rsLandlord!LandlordHomeTel = rsSupplier!SupplierHomeTel
'           rsLandlord!LandlordMobile = rsSupplier!SupplierMobile
'           rsLandlord!LandlordOfficeAddressLine1 = rsSupplier!SupplierOfficeAddressLine1
'           rsLandlord!LandlordOfficeAddressLine2 = rsSupplier!SupplierOfficeAddressLine2
'           rsLandlord!LandlordOfficeAddressLine3 = rsSupplier!SupplierOfficeAddressLine3
'           rsLandlord!LandlordOfficeAddressLine4 = rsSupplier!SupplierOfficeAddressLine4
'           rsLandlord!LandlordOfficePostCode = rsSupplier!SupplierOfficePostCode
'           rsLandlord!LandlordOfficeTel = rsSupplier!SupplierOfficeTel
'           rsLandlord!LandlordMemo = rsSupplier!SupplierMemo
'           rsLandlord!LandLordSageSuppAC = rsSupplier!SageSuppAC
'           rsLandlord!VATReg = rsSupplier!VATReg
'           rsLandlord!AcBalance = rsSupplier!AcBalance
'           rsLandlord!BacsRef = rsSupplier!BacsRef
'           rsLandlord.Update
'        rsSupplier.MoveNext
'   Wend
'   adoConn.Close
'   Set adoConn = Nothing

End Sub

Private Sub cmdAnalysis_Click()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection

   frmSecondaryCode.PRIMARY_CODE_SHOW = "ATYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoConn.Open getConnectionString
   ' Analysis Type
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'ATYP'"

   populateCombo adoConn, sSQLQuery, cboAnalysisType

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdAnalysisCancel_Click()
   PropertyAnalysisButtonMode DefaultMode
   fmeProperty.Enabled = True
   cmdAnalysisEdit.Enabled = False
End Sub

Private Sub cmdAsJS_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If lblAsJS_PO = "Print as..." Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue Mid(gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3), 6)

      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "JOB SHEET"

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
   If lblAsJS_PO = "Email as..." Then
      
   End If
   cmdAsJS_KeyPress 27
End Sub

Private Sub cmdAsJS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame1(0).Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdAsPO_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   If lblAsJS_PO = "Print as..." Then
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue Mid(gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3), 6)

      Report.ParameterFields(2).AddCurrentValue "Job Name"
      Report.ParameterFields(3).AddCurrentValue "PURCHASE ORDER"

      Load frmReport
      frmReport.LoadReportViewer Report
   End If

   cmdAsJS_KeyPress 27
End Sub

Private Sub cmdAsPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame1(0).Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdDeleteAnalysis_Click()
   If txtPropertyAnalysisID.text = "" Then Exit Sub

   If MsgBox("Do you wish to delete the seleted analysis?", vbQuestion + vbYesNo, "Deleting Analysis") = vbNo Then Exit Sub

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   adoConn.Execute "DELETE * FROM PropertyAnalysis WHERE PropertyAnalysisID = " & txtPropertyAnalysisID.text & ";"

   PropertyAnalysisButtonMode DefaultMode
   PopulateGridPropertyAnalysis adoConn

   adoConn.Close
   Set adoConn = Nothing

   fmeProperty.Enabled = True
   cmdDeleteAnalysis.Enabled = False
   ShowMsgInTaskBar "Property analysis has been deleted successfully", "Y", "P"
End Sub

Private Sub cmdAnalysisEdit_Click()
   If txtPropertyAnalysisID.text = "" Then Exit Sub
   PropertyAnalysisButtonMode EditMode
   Property_ANALYSIS_NEW_ENTRY = False
   'Modified by Anol 02 Nov 2014
   'Issue 494  Property Form - Property analysis Form
   fmeProperty.Enabled = False
   If cboAnalysisType.Enabled = True Then
      cboAnalysisType.SetFocus
   End If
   txtAnalysisValue.Locked = False
    txtAnalysisQuantity.Locked = False
    txtAnalysisPercentage.Locked = False
    txtAnalysisValue1.Locked = False
    txtAnalysisReference.Locked = False
End Sub

Private Sub cmdAnalysisNew_Click()
   PropertyAnalysisButtonMode NewEntryMode
   Property_ANALYSIS_NEW_ENTRY = True
   fmeProperty.Enabled = False
   txtAnalysisValue.Locked = False
    txtAnalysisQuantity.Locked = False
    txtAnalysisPercentage.Locked = False
    txtAnalysisValue1.Locked = False
    txtAnalysisReference.Locked = False
   'issue 494
   cboAnalysisType.SetFocus
End Sub

Private Sub cmdAnalysisSave_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   If cboAnalysisType.text = "" Then
      ShowMsgInTaskBar "Please enter Analysis Type."
      cboAnalysisType.SetFocus
      Exit Sub
   End If
   If cboAnalysisOption.ListIndex = -1 Then
       ShowMsgInTaskBar "Please select Analysis Option."
       cboAnalysisOption.SetFocus
       Exit Sub
   End If
    
   Dim rsTransaction As New ADODB.Recordset
   'adoConn.Open getConnectionString
   rsTransaction.Open "Select * from Property where PropertyID='" & txtPropertyID.text & "'", adoConn, adOpenKeyset
   If rsTransaction.EOF Then
        MsgBox "This property ID was not found in the database", vbInformation, "Not found"
        FocusControl cmdPropertyLookup
        rsTransaction.Close
        Exit Sub
   End If
   rsTransaction.Close
   If IsNull(cboAnalysisType.Value) = True Then
        MsgBox "Please enter Analysis Type"
        FocusControl cboAnalysisType
        Exit Sub
   End If
   
   If IsNull(cboAnalysisOption.Value) = True Then
        MsgBox "Please select Analysis Option."
        FocusControl cboAnalysisOption
        Exit Sub
   End If
   If SavePropertyAnalysis(adoConn) Then
      PopulateGridPropertyAnalysis adoConn
      ShowMsgInTaskBar "The property analysis has been saved successfully."
      SetTotalArea
   Else
       ShowMsgInTaskBar "Could not save Property analysis", , "N"
   End If
   adoConn.Close
   Set adoConn = Nothing
   PropertyAnalysisButtonMode DefaultMode
   fmeProperty.Enabled = True
   cmdAnalysisEdit.Enabled = False
End Sub

Private Sub cmdAttachment_Click()
   Me.Enabled = False
   Load frmAttachment

   If HEALTH_SAFETY_NEW_ENTRY Then
      If txtUnitSafetyID.text = "" Then txtUnitSafetyID.text = UniqueID()
   Else
      txtUnitSafetyID.text = gridSafety.TextMatrix(gridSafety.row, 0)
   End If

   HEALTH_N_SAFETY_ATTACH = False

   frmAttachment.OwnerID = txtUnitSafetyID.text
   frmAttachment.CallerForm = "Property"
   frmAttachment.Show
End Sub

Private Sub cmdCancelProperty_Click()
   ComponentEnableModeProperty Me, DefaultMode
   imgPropertyPicture.Picture = LoadPicture()
   lblImageName.Caption = "Image Name:"
   NEWMODE_ = False
   SEARCHPropertyMODE_ = True
   txtPropertyID.Enabled = True
   txtPropertyName.Enabled = True
   cmdPropertyLookup.Enabled = True
  ' cboClientID.Enabled = True
   cmdClientList.Enabled = True
   If txtPropertyID.text = "" Then
       Exit Sub
   End If
   tabProperty.Enabled = True
   gridPropertyAnalysis.Clear
   gridPropertyAnalysis.Rows = 2
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub

   AddNewAttachmentInCombo cmbFiles, "Property", txtPropertyID.text

   ShowMsgInTaskBar "The file has been saved successfully."
End Sub

Private Sub cmdCloseProperty_Click()
Unload Me
End Sub


Private Sub cmdCurrentTenant_Click()
Load frmLeasee1
frmLeasee1.Show
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub

   If MsgBox("Are you sure you want to delete the file " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachmentCombo cmbFiles, cmbFiles.Column(2), txtPropertyID.text, "Property"

   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdDeleteLandlord_Click()
   Dim landLordID As String
   
   landLordID = flxLandlordGrid.TextMatrix(flxLandlordGrid.row, 8)
   
   If (landLordID = "") Then
      ShowMsgInTaskBar "Select a Row to delete."
      Exit Sub
   Else
      If MsgBox("Are you sure you want to delete the Landlord from this Property?", vbQuestion + vbYesNo, "Landlord - Delete") = vbNo Then Exit Sub

      Dim adoConn As New ADODB.Connection

      adoConn.Open getConnectionString

      adoConn.Execute "DELETE * FROM PropertyLandlord WHERE LandlordID = '" & landLordID & "' AND PropertyId ='" & txtPropertyID.text & "';"

      adoConn.Close
      Set adoConn = Nothing
      
      PopulateCodes
      
      Dim conProperty_ As New ADODB.Connection
      conProperty_.Open getConnectionString
      ConfigFlxLandlordGrid
      loadflxLandlordGrid conProperty_
      
      conProperty_.Close
      Set conProperty_ = Nothing
   
      ShowMsgInTaskBar "Selected Landlord deleted from Property."
   End If
   
End Sub
'
'Private Sub cmdEditMHistory_Click()
'If txtID.text = "" Then
'    Exit Sub
'End If
'
'MaintenanceHistoryButtonMode EditMode
'M_HISTORY_NEW_ENTRY_ = False
''cboMaintenanceType.SetFocus
'End Sub

Private Sub cmdEditMHistory_Click()
   If gridMaintenanceHistory.TextMatrix(1, 0) = "" Then Exit Sub

   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 21) <> "P" Then
      MsgBox "You cannot amend this record from property module." & Chr(13) & _
             "This " & gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) & " " & _
             "has been added from " & _
             IIf(gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 21) = "U", "Unit", "Lessee") & " " & _
             "module.", vbCritical + vbOKOnly, "Maintenance Entry"

      Exit Sub
   End If

   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      frmMaintenanceJob.isEdit = True
      frmMaintenanceJob.CallingForm = "P"                         'Property
      frmMaintenanceJob.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintenanceJob
      frmMaintenanceJob.ZOrder 0
      frmMaintenanceJob.Show
   Else
      frmMaintananceDairy.isEdit = True
      frmMaintananceDairy.CallingForm = "P"                         'Property
      frmMaintananceDairy.UpdateRow = gridMaintenanceHistory.row
      Load frmMaintananceDairy
      frmMaintananceDairy.ZOrder 0
      frmMaintananceDairy.Show
   End If
   Me.Enabled = False
End Sub

Private Sub cmdEditProperty_Click()
   If txtPropertyID.text = "" Then
      ShowMsgInTaskBar "Please select a Property to continue."
      Exit Sub
   End If

   NEWMODE_ = False
   SEARCHPropertyMODE_ = False
   ComponentEnableModeProperty Me, EditMode
   cmdClientList.Enabled = False
   szEditingPropID = txtPropertyID.text
'   cboClientID.SetFocus
    tabProperty.Enabled = False
    cmdPropertyLookup.Enabled = False
    
    cmdAnalysisNew.Enabled = True
    cmdAnalysisEdit.Enabled = False
    cmdAnalysisSave.Enabled = False
    cmdDeleteAnalysis.Enabled = False
    cmdAnalysisCancel.Enabled = False
    
    cmdDelete.Enabled = False
    tabProperty.Enabled = True
    
    txtAnalysisValue.Locked = True
    txtAnalysisQuantity.Locked = True
    txtAnalysisPercentage.Locked = True
    txtAnalysisValue1.Locked = True
    txtAnalysisReference.Locked = True
    txtProAddressLine1.Enabled = True
    txtProAddressLine2.Enabled = True
    txtProAddressLine3.Enabled = True
    txtProAddressLine4.Enabled = True
    txtProPostCode.Enabled = True
    txtContactDetails.Enabled = True
    txtManager.Enabled = True
    txtPropertyID.Locked = True
End Sub

Private Sub cmdEmailJS_PO_Click()
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      fraJS_PO.Top = Frame1(0).Top + cmdEmailJS_PO.Top + cmdEmailJS_PO.Height - fraJS_PO.Height
      fraJS_PO.Left = cmdEmailJS_PO.Left + Frame1(0).Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
      Frame1(0).Enabled = False
      lblAsJS_PO = "Email as..."
   End If
End Sub

Private Sub cmdGridPropertyLookup_Click()
    fmePropertyLookup.Visible = False
    fmeProperty.Enabled = True
    tabProperty.Enabled = True
End Sub

Private Sub cmdImgDelete_Click()
   If imgPropertyPicture.Picture = 0 Then Exit Sub
   If MsgBox("Are you sure to delete the image?", vbQuestion + vbYesNo, "Delete Image") = vbNo Then Exit Sub
   DeleteImage imgPropertyPicture, IMAGE_FILE_NAME_, txtPropertyID.text, "Property"
   ShowMsgInTaskBar "File has been deleted successfully"
End Sub

Private Sub cmdImgLeftMove_Click()
   IMAGE_FILE_NAME_ = MoveNextImage(imgPropertyPicture, txtPropertyID.text, "Property", IMAGE_FILE_NAME_, lblImageName)
End Sub

Private Sub cmdInspectedBy_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "IPT"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInspector.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IPT'"

   adoInspector.RecordSource = sSQLQuery_
   adoInspector.CommandType = adCmdText
   adoInspector.Refresh
End Sub

Private Sub cmdInsuranceCancel_Click()
   InsuranceButtonMode DefaultMode
End Sub
'
'Private Sub cmdInsuranceEdit_Click()
'   If txtPropertyInsuranceID.text = "" Then
'       Exit Sub
'   End If
'
'   InsuranceButtonMode EditMode
'   Property_INSURANCE_NEW_ENTRY = False
'End Sub

Private Sub cmdInsuranceEdit_Click()
   If txtPropertyInsuranceID.text = "" Then
      Exit Sub
   End If

   InsuranceButtonMode EditMode
   UNIT_INSURANCE_NEW_ENTRY = False
   fmeProperty.Enabled = False
End Sub

Private Sub cmdInsuranceNew_Click()
   InsuranceButtonMode NewEntryMode
   UNIT_INSURANCE_NEW_ENTRY = True
   INSURANCE_ID = ""
   fmeProperty.Enabled = False
End Sub

Private Sub cmdInsuranceSave_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If SavePropertyInsurance(adoConn) Then
      ShowMsgInTaskBar "The insurance information have been saved successfully."
      PopulateInsurance adoConn
   Else
       ShowMsgInTaskBar "Could not save the Insurance Information", , "N"
   End If

   InsuranceButtonMode DefaultMode
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdLandlord_Click()
Load frmClientNew4
frmClientNew4.Show
End Sub
'
'Private Sub cmdMntJob_Click(Index As Integer)
'   Dim sSQLQuery As String
'   Dim adoConn As New ADODB.Connection
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "MNTJOB"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoConn.Open getConnectionString
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'MNTJOB'"
'   populateCombo adoConn, sSQLQuery, cboTaskOwner
'   populateCombo adoConn, sSQLQuery, cboContact
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub
'
'Private Sub cmdMType_Click()
'   Dim sSQLQuery As String
'   Dim adoConn As New ADODB.Connection
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "MTYP"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoConn.Open getConnectionString
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'MTYP'"
'   populateCombo adoConn, sSQLQuery, cboMaintenanceType
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub
'
'Private Sub cmdNewMHistory_Click()
'MaintenanceHistoryButtonMode NewEntryMode
'M_HISTORY_NEW_ENTRY_ = True
'End Sub
'
'Private Sub cmdMntJob_Click(Index As Integer)
'
'End Sub

Private Sub cmdMemoCancel_Click()
   MemoButtonMode DefaultMode
End Sub

Private Sub cmdMemoEdit_Click()
   MemoButtonMode EditMode
End Sub

Public Sub MemoButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    
    Select Case mode
        Case ComponentMode.DefaultMode
            cmdMemoEdit.Enabled = True
            cmdMemoSave.Enabled = False
            cmdMemoCancel.Enabled = False
            
            txtMemo.Enabled = False
            Frame17.Enabled = False
            cmbFiles.Enabled = False
            
        Case ComponentMode.GridRowOnSelection
            cmdMemoEdit.Enabled = True
            cmdMemoSave.Enabled = False
            cmdMemoCancel.Enabled = False
            
            txtMemo.Enabled = False
            Frame17.Enabled = False
            cmbFiles.Enabled = False

        Case ComponentMode.NewEntryMode
            cmdMemoEdit.Enabled = False
            cmdMemoSave.Enabled = True
            cmdMemoCancel.Enabled = True
            
            txtMemo.Enabled = True
            Frame17.Enabled = True
            cmbFiles.Enabled = True
                    
        Case ComponentMode.EditMode
            cmdMemoEdit.Enabled = False
            cmdMemoSave.Enabled = True
            cmdMemoCancel.Enabled = True
            
            txtMemo.Enabled = True
            Frame17.Enabled = True
            cmbFiles.Enabled = True
    End Select
End Sub

Private Sub cmdMemoSave_Click()
   If SavePropertyMemo Then
      ShowMsgInTaskBar "The memo has been saved successfully."
   Else
      ShowMsgInTaskBar "Could not save the memo successfully", , "N"
   End If

   MemoButtonMode DefaultMode
End Sub

Private Sub cmdNewMHistory_Click()
'   If txtPropertyID.text = "" Then Exit Sub
'
'   Load frmMaintenanceJob
'   With frmMaintenanceJob
'      .CallingForm = "P"          'Calling from property form
'      .RecordType = "J"
'      .lblJobName.Caption = "Job Name"
'      .Label1.Caption = "Job No."
'      .txtRef.Enabled = True
'      .isEdit = False
'      .Show
'      .ZOrder 0
'   End With
'
'   Me.Enabled = False
    If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
        With frmMaintenanceJob
          .isEdit = True
          .CallingForm = "P"          'Calling from property form
          .RecordType = "J"
          .lblJobName.Caption = "Job Name"
          .Label1.Caption = "Job No."
          .txtRef.Enabled = True
          .UpdateRow = gridMaintenanceHistory.row
          .Frame1.Enabled = False
          .Show
          .ZOrder 0
       End With
    Else
        ShowMsgInTaskBar "Please select a job."
    End If
End Sub

Private Sub cmdAddDiary_Click()
   If txtPropertyID.text = "" Then Exit Sub

   With frmMaintananceDairy
      .CallingForm = "P"          'Calling from property form
      .isEdit = False
      .RecordType = "D"
      .lblJobName.Caption = "Diary Name"
      .Label1.Caption = "Diary Entry No."
      Load frmMaintananceDairy
      .txtRef.Enabled = True
      .isEdit = False
      .Show
      .ZOrder 0
   End With

   Me.Enabled = False
End Sub

Private Sub ConfigureTabs()
   tabProperty.Tab = 0
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2

   gridUtilities.Clear
   gridUtilities.Rows = 2
   gridUtilities.Cols = 8
End Sub

Private Sub cmdNewProperty_Click()
   ConfigureTabs
   
   cmdClientList.Enabled = True
   If txtClientList.Tag = "" Then
       ShowMsgInTaskBar "Please select a client to continue."
       cmdClientList.SetFocus
       Exit Sub
   End If

   NEWMODE_ = True
   SEARCHPropertyMODE_ = False
   ComponentEnableModeProperty Me, NewEntryMode
   imgPropertyPicture.Picture = LoadPicture()
   lblImageName.Caption = "Image Name:"
   'issue 298 COmponenet not enable by anol 20170202
   txtManager.Enabled = True
   txtContactDetails.Enabled = True
   txtProAddressLine1.Enabled = True
   txtProAddressLine2.Enabled = True
   txtProAddressLine3.Enabled = True
   txtProAddressLine4.Enabled = True
   txtProPostCode.Enabled = True
   txtContactDetails.Locked = False
   txtProAddressLine1.Locked = False
   txtProAddressLine2.Locked = False
   txtProAddressLine3.Locked = False
   txtProAddressLine4.Locked = False
   txtProPostCode.Locked = False
   tabProperty.Enabled = False
   cmdPropertyLookup.Enabled = False
   txtPropertyID.Locked = False
   If txtPropertyID.Enabled Then txtPropertyID.SetFocus
   cmdCloseProperty.Enabled = True
   gridPropertyAnalysis.Clear
   gridPropertyAnalysis.Rows = 2
   cmdDelete.Enabled = False
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   'MousePointer = vbHourglass

   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      ShowMsgInTaskBar "File has been moved from original location."

   MousePointer = vbDefault
End Sub

Private Sub cmdPrintJobSheet_Click()
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      fraJS_PO.Top = Frame1(0).Top + cmdEmailJS_PO.Top + cmdEmailJS_PO.Height - fraJS_PO.Height
      fraJS_PO.Left = cmdPrintJobSheet.Left + Frame1(0).Left
      fraJS_PO.Visible = True
      cmdAsJS.SetFocus
      Frame1(0).Enabled = False
      lblAsJS_PO = "Print as..."
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\JobSheet.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue Mid(gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 3), 6)

   Report.ParameterFields(2).AddCurrentValue "Diary Entry"
   Report.ParameterFields(3).AddCurrentValue "DIARY ENTRY"

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub cmdQuoteReq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      fraJS_PO.Visible = False
      Frame1(0).Enabled = True
      gridMaintenanceHistory.SetFocus
   End If
End Sub

Private Sub cmdSafety_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "STYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoSafetyType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'STYP'"

   adoSafetyType.RecordSource = sSQLQuery_
   adoSafetyType.CommandType = adCmdText
   adoSafetyType.Refresh
End Sub

Private Sub cmdSafetyCancel_Click()
HealthSafetyButtonMode DefaultMode
End Sub

Private Sub cmdSafetyEdit_Click()
   If txtUnitSafetyID.text = "" Then Exit Sub

   HealthSafetyButtonMode EditMode
   HEALTH_SAFETY_NEW_ENTRY = False
   If gridSafety.TextMatrix(gridSafety.row, 11) = "Yes" Then
      HEALTH_N_SAFETY_ATTACH = True
   Else
      HEALTH_N_SAFETY_ATTACH = False
   End If
End Sub

Private Sub cmdSafetyNew_Click()
   HealthSafetyButtonMode NewEntryMode
   HEALTH_SAFETY_NEW_ENTRY = True
   txtUnitSafetyID.text = ""
End Sub

Private Sub cmdSafetySave_Click()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   If SaveHealthSafety(adoConn) Then
      ShowMsgInTaskBar "The Health and Safety Information have been saved successfully."
      PopulateHealthSafety adoConn
   Else
       ShowMsgInTaskBar "Could not save The Health and Safety Information", , "N"
   End If

   HealthSafetyButtonMode DefaultMode
   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdSaveLandlord_Click()
   Dim adoConn As New ADODB.Connection
   Dim rsPropertyLandlord  As New ADODB.Recordset
    If txtLLID.text = "" Then
        MsgBox "Please select a landlord to add", vbInformation, "Warning"
        Exit Sub
   End If
    adoConn.Open getConnectionString
    If check_dupplicateLandlord(adoConn) = True Then
         MsgBox "This landlord already exists in the current list.", vbInformation, "Warning!"
         Exit Sub
     End If
'   cmdSaveLandlord.Enabled = True
'   cmdAddLandlord.Enabled = False
   If MsgBox("Do you want to add a landlord to this property?", vbYesNo, "add landlord") = vbYes Then
            rsPropertyLandlord.Open "Select * from PropertyLandlord", adoConn, adOpenDynamic, adLockOptimistic
            With rsPropertyLandlord
                .AddNew
                !propertyID = txtPropertyID.text
                !landLordID = txtLLID.text
                .Update
            End With
            MsgBox "This landlord added successfully.", vbInformation, "Warning!"
            loadflxLandlordGrid adoConn
            txtLLID.text = ""
            txtLLName.text = ""
            adoConn.Close
            Set adoConn = Nothing
   End If
   cmdSaveLandlord.Enabled = False
   cmdAddLandlord.Enabled = True
   
   
'   Dim conProperty_ As New ADODB.Connection
'   Dim rstProperty_ As New ADODB.Recordset
'   Dim sSQLQuery_ As String
'    If txtPropertyID.text = "" Then
'        MsgBox "Please select the property", vbInformation, "Warning"
'        Exit Sub
'    End If
'   'On Error GoTo Exception
'
'      'Set the ado Connections to the dataset
'      conProperty_.Open getConnectionString
'
'      sSQLQuery_ = "SELECT * FROM PROPERTYLANDLORD"
'
'      rstProperty_.Open sSQLQuery_, conProperty_, adOpenDynamic, adLockOptimistic
'
'      rstProperty_.AddNew
'
'      rstProperty_!propertyID = txtPropertyID.text
'      rstProperty_!landLordID = txtLLID.text
'
'      rstProperty_.Update
'
'      rstProperty_.Close
'      conProperty_.Close
'
'      Set rstProperty_ = Nothing
'      Set conProperty_ = Nothing
'
'      PopulateCodes
'
'      'Set the ado Connections to the dataset
'      conProperty_.Open getConnectionString
'      loadflxLandlordGrid conProperty_
'
'      cmdSaveLandlord.Enabled = False
'
'      cmdAddLandlord.Enabled = True
'
'      conProperty_.Close
'      Set conProperty_ = Nothing
End Sub

Private Sub cmdSaveProperty_Click()
   If InStr(frmMMain.rtxtMessageDisplay.text, "Property ID already exits") > 0 Then Exit Sub

   If txtClientList.text = "" Then
       ShowMsgInTaskBar "Please select a client to continue."
       txtClientList.text = ""
      cmdClientList.SetFocus
       Exit Sub
   End If
   If txtClientList.Tag = "ALL" Then
        MsgBox "Please select a client from the list to save information"
        Exit Sub
   End If
   If txtPropertyName.text = "" Then
      ShowMsgInTaskBar "Please enter a Property Name to continue."
      txtPropertyName.SetFocus
      txtPropertyName.text = ""
      Exit Sub
   End If

   If txtPropertyID.text = "" Then
       ShowMsgInTaskBar "Please enter a code of Property to continue."
       txtPropertyID.SetFocus
       Exit Sub
   End If
'   If Len(txtPropertyID.text) < 4 Then
'      ShowMsgInTaskBar "The property code must be 4 charecters long.", , "N"
'      SelTxtInCtrl txtPropertyID
'      txtPropertyID.SetFocus
'      Exit Sub
'   End If
   If MsgBox("Do you wish to save property information?", vbQuestion + vbYesNo, "Property") = vbYes Then
      If SavePropertyInformation Then
         ShowMsgInTaskBar "The Property Information added successfully."
      Else
         ShowMsgInTaskBar "An error occured while saving the Property Information.", , "N"
      End If
   End If
   NEWMODE_ = False
   ComponentEnableModeProperty Me, SavedMode
   SEARCHPropertyMODE_ = True
   txtPropertyID.Enabled = True
   txtPropertyName.Enabled = True
   tabProperty.Enabled = True
   cmdPropertyLookup.Enabled = True
'   cboClientID.Enabled = True
   txtClientList.text = ""
   txtPropertyID.text = ""
   txtPropertyName.text = ""
   txtProAddressLine1.text = ""
   txtProAddressLine2.text = ""
   txtProAddressLine3.text = ""
   txtProAddressLine4.text = ""
   txtProPostCode.text = ""
   txtManager.text = ""
   txtContactDetails.text = ""
   gridPropertyAnalysis.Clear
   gridPropertyAnalysis.Rows = 2
End Sub

Private Sub cmdSetInsuranceType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "ITYP"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsuranceType.Refresh
End Sub
'
'Private Sub cmdSetRechargeableFrequency_Click()
'   Dim sSQLQuery As String
'   Dim adoConn As New ADODB.Connection
'
'   frmSecondaryCode.PRIMARY_CODE_SHOW = "RFRQ"
'   Load frmSecondaryCode
'   frmSecondaryCode.Show 1
'
'   adoConn.Open getConnectionString
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'RFRQ'"
'   populateCombo adoConn, sSQLQuery, cboRechargeableFreq
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub cmdSetInsurer_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "IRER"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsurer.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'IRER'"

   adoInsurer.RecordSource = sSQLQuery_
   adoInsurer.CommandType = adCmdText
   adoInsurer.Refresh
End Sub

Private Sub cmdSetUtilitiesType_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UTIL"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoUtilitiesType.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTIL'"

   adoUtilitiesType.RecordSource = sSQLQuery_
   adoUtilitiesType.CommandType = adCmdText
   adoUtilitiesType.Refresh
End Sub

Private Sub cmdPropertyLookup_Click()
'   If cboClientID.text = "" Then
'      ShowMsgInTaskBar "Please select a client to continue."
'      Exit Sub
'   End If
    fmeProperty.Enabled = False
    tabProperty.Enabled = False
   fmePropertyLookup.Top = fmeProperty.Top + txtPropertyID.Top + txtPropertyID.Height + 5
   fmePropertyLookup.Left = fmeProperty.Left + txtPropertyID.Left
   fmePropertyLookup.Visible = True
   fmePropertyLookup.ZOrder 0
   gridPropertyLookup.Visible = True
   txtSearchProperty.text = ""
   txtSearchProperty.Enabled = True
   
   txtPropertySearch.text = ""
   txtPropertySearch.Enabled = True
   txtSearchProperty.SetFocus

   'PopulatePropertyLookup " WHERE CLIENTID = '" & txtClientList.tag & "'"
   PopulatePropertyLookup ""
   txtSearchProperty.Locked = False
   txtPropertySearch.Locked = False
End Sub

Private Sub cmdSupCancel_Click()
   Frame1(7).Enabled = False
   Frame1(8).Enabled = False
End Sub
'
'Private Sub cmdSupEdit_Click()
'   Frame1(7).Enabled = True
'   Frame1(8).Enabled = True
'   txtSupp1.Enabled = True
'   txtSupp2.Enabled = True
'   txtSupp3.Enabled = True
'   txtSuppCaption1.Enabled = True
'   txtSuppCaption2.Enabled = True
'   txtSuppCaption3.Enabled = True
'   txtDtFlgDate.Enabled = True
'   txtDtFlgDt2.Enabled = True
'   txtDtFlgDt3.Enabled = True
'   txtDtFlgDesc.Enabled = True
'   txtDtFlgDesc2.Enabled = True
'   txtDtFlgDesc3.Enabled = True
'End Sub
'
'Private Sub cmdSupSave_Click()
'   Frame1(7).Enabled = False
'   Frame1(8).Enabled = False
'
'   If MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Supplementary") = vbNo Then Exit Sub
'
'   Dim conConn As New ADODB.Connection
'   Dim rstSQL As New ADODB.Recordset
'   Dim szSQL As String
'
'   conConn.Open getConnectionString
'
'   szSQL = "SELECT SuppCaption1, SuppCaption2, SuppCaption3, " & _
'               "SuppText1, SuppText2, SuppText3, " & _
'               "DateFlagDt1, DateFlagDt2, DateFlagDt3, " & _
'               "DateFlagDescription1, DateFlagDescription2, DateFlagDescription3 " & _
'           "FROM Property " & _
'           "WHERE PropertyID = '" & txtPropertyID.text & "';"
'   rstSQL.Open szSQL, conConn, adOpenDynamic, adLockOptimistic
'
'   With rstSQL
'      .Fields.Item(0).Value = lblSupplementary1.Caption
'      .Fields.Item(1).Value = lblSupplementary2.Caption
'      .Fields.Item(2).Value = lblSupplementary3.Caption
'      .Fields.Item(3).Value = txtSupp1.text
'      .Fields.Item(4).Value = txtSupp2.text
'      .Fields.Item(5).Value = txtSupp3.text
'      .Fields.Item(6).Value = txtDtFlgDate.text
'      .Fields.Item(7).Value = txtDtFlgDt2.text
'      .Fields.Item(8).Value = txtDtFlgDt3.text
'      .Fields.Item(9).Value = txtDtFlgDesc.text
'      .Fields.Item(10).Value = txtDtFlgDesc2.text
'      .Fields.Item(11).Value = txtDtFlgDesc3.text
'
'      .Update
'   End With
'
'   rstSQL.Close
'   conConn.Close
'   Set rstSQL = Nothing
'   Set conConn = Nothing
'
'   ShowMsgInTaskBar "The data has been saved successfully."
'End Sub

Private Sub cmdUnitStatus_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "USTA"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoStatus.Refresh
End Sub

Private Sub cmdUploadImageAdd_Click()
   If MsgBox("Do you want to add new image?", vbQuestion + vbYesNo, "Image Attachment") = vbNo Then Exit Sub
   IMAGE_FILE_NAME_ = AddNewImage(imgPropertyPicture, "Property", txtPropertyID.text, lblImageName)
   ShowMsgInTaskBar "Image has been uploaded successfully."
End Sub

Private Sub cmdUsage_Click()
   Dim sSQLQuery_ As String

   frmSecondaryCode.PRIMARY_CODE_SHOW = "UUSE"
   Load frmSecondaryCode
   frmSecondaryCode.Show 1

   adoInsUsage.ConnectionString = getConnectionString

   sSQLQuery_ = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UUSE'"

   adoInsUsage.RecordSource = sSQLQuery_
   adoInsUsage.CommandType = adCmdText
   adoInsUsage.Refresh
End Sub

Private Sub cmdUtilitiesAttach_Click()
   Me.Enabled = False
   Load frmAttachment

   If UNIT_INSURANCE_NEW_ENTRY Then
      If INSURANCE_ID = "" Then INSURANCE_ID = UniqueID()
   Else
      INSURANCE_ID = gridSafety.TextMatrix(gridSafety.row, 0)
   End If

   HEALTH_N_SAFETY_ATTACH = False

   frmAttachment.OwnerID = INSURANCE_ID
   frmAttachment.CallerForm = "Property_Insurance"
   frmAttachment.Show
End Sub

Private Sub cmdUtilitiesCancel_Click()
   UtilitiesButtonMode DefaultMode
   fmeProperty.Enabled = True
End Sub

Private Sub cmdUtilitiesEdit_Click()
   If txtPropertyID.text = "" Then
       Exit Sub
   End If

   UtilitiesButtonMode EditMode
   UNIT_UTILITIES_NEW_ENTRY = False
   fmeProperty.Enabled = False
End Sub

Private Sub cmdUtilitiesNew_Click()
   UtilitiesButtonMode NewEntryMode

   UNIT_UTILITIES_NEW_ENTRY = True

   cboUtilitiesType.SetFocus
   fmeProperty.Enabled = False
End Sub

Private Sub cmdUtilitiesSave_Click()
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   If SavePropertyUtilities(adoConn) Then
      ShowMsgInTaskBar "The Utilities Information have been saved successfully."
      PopulateUtilities adoConn
   Else
      ShowMsgInTaskBar "Could not save Utilities Information", , "N"
   End If

   UtilitiesButtonMode DefaultMode

   adoConn.Close
   Set adoConn = Nothing
   fmeProperty.Enabled = True
End Sub

Private Sub Command1_Click()

End Sub

'Private Sub dtpDateCompleted_Change()
'   'Added By Samrat. 16/01/2006
'   TextBoxChangeDate dtpDateCompleted
'End Sub
'
'Private Sub dtpDateCompleted_KeyPress(KeyAscii As Integer)
'   'Added By Samrat. 16/01/2006
'   TextBoxKeyPrsDate dtpDateCompleted, KeyAscii
'End Sub
'
'Private Sub dtpDateCompleted_LostFocus()
'   'Added By Asif. 13/01/2006
'   TextBoxFormatDate dtpDateCompleted
'End Sub
'
'Private Sub dtpRemindDate_Change()
'   'Added By Samrat. 16/01/2006
'   TextBoxChangeDate dtpRemindDate
'End Sub
'
'Private Sub dtpRemindDate_KeyPress(KeyAscii As Integer)
'   'Added By Samrat. 16/01/2006
'   TextBoxKeyPrsDate dtpRemindDate, KeyAscii
'End Sub
'
'Private Sub dtpRemindDate_LostFocus()
'   'Added By Asif. 13/01/2006
'   TextBoxFormatDate dtpRemindDate
'End Sub
'
'Private Sub dtpReportedDate_Change()
'   'Added By Samrat. 16/01/2006
'   TextBoxChangeDate dtpReportedDate
'End Sub
'
'Private Sub dtpReportedDate_KeyPress(KeyAscii As Integer)
'   'Added By Samrat. 16/01/2006
'   TextBoxKeyPrsDate dtpReportedDate, KeyAscii
'End Sub
'
'Private Sub dtpReportedDate_LostFocus()
'   'Added By Asif. 13/01/2006
'   TextBoxFormatDate dtpReportedDate
'End Sub

Private Sub flxLandlordGrid_Click()
   cmdDeleteLandlord.Enabled = True
End Sub

'Private Sub fmeProperty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''   fmeProperty.MousePointer = vbArrow
'End Sub

Private Sub Form_Activate()
   If LOAD_PROPERTY_PROPERTYID <> "" Then
      txtClientList.text = CLIENT_NAME
      LoadProertyByPropertyID
      'Me.Caption = cboClientID.Column(1) + " - " + txtPropertyName.text
   End If
   txtSearchClientID.Locked = False
   txtSearchClientName.Locked = False
End Sub

Private Sub Form_Load()
   

'   Me.Top = 0
'   Me.Left = 0
   Me.Height = 9195 '8355
   Me.Width = 12675
   Me.BackColor = MODULEBACKCOLOR
   tabProperty.BackColor = MODULEBACKCOLOR
   fmeProperty.BackColor = MODULEBACKCOLOR

   Me.Caption = "Properties"
   DSN_ALARM_ = "WD_ALARM"
   ComponentEnableModeProperty Me, DefaultMode

   tabProperty.Tab = 0

   txtSearchProperty.Enabled = True
   NEWMODE_ = False
   SEARCHPropertyMODE_ = True
   tabProperty.Enabled = False
'   txtClientID.Enabled = True

   '' Populate the codes
   PopulateCodes

   '' Button Modes''
   MaintenanceHistoryButtonMode DefaultMode
   PropertyAnalysisButtonMode DefaultMode
   HealthSafetyButtonMode DefaultMode
   InsuranceButtonMode DefaultMode
   UtilitiesButtonMode DefaultMode

   ''Set the grids
   ConfigFlxLandlordGrid

   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
    
   SetGridPropertyAnalysisHeader adoConn
   setGridUtilities adoConn
   ConfigureGridSafety
   Dim szSQL As String
   Dim adoRST As New ADODB.Recordset
   szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   If Not adoRST.EOF Then
        txtClientList.Tag = adoRST.Fields("CLIENTID").Value
        txtClientList.text = adoRST.Fields("CLIENTNAME").Value
        
   End If
   adoRST.Close
   adoConn.Close
   Set adoConn = Nothing
    If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
       Call WheelHook(Me.hWnd)
    End If
    cmdAnalysisNew.Enabled = False
    cmdAnalysisEdit.Enabled = False
    cmdAnalysisSave.Enabled = False
    cmdDeleteAnalysis.Enabled = False
    cmdAnalysisCancel.Enabled = False
    cmdSupplier.Enabled = False
   'MousePointer = vbDefault
End Sub

Private Sub ConfigFlxLandlordGrid()
   Dim i As Integer

   flxLandlordGrid.Cols = 9
   flxLandlordGrid.RowHeight(0) = 0
   flxLandlordGrid.Rows = 2
   flxLandlordGrid.RowHeight(1) = 240
   For i = 0 To flxLandlordGrid.Cols - 3
      flxLandlordGrid.ColWidth(i) = lblLandlord(i + 1).Left - lblLandlord(i).Left
   Next i
   flxLandlordGrid.ColWidth(flxLandlordGrid.Cols - 2) = flxLandlordGrid.Left + flxLandlordGrid.Width - _
                                                        lblLandlord(i).Left - 125
   flxLandlordGrid.ColWidth(8) = 0
End Sub

Private Sub loadflxLandlordGrid(ByVal conLandlord As ADODB.Connection)
   Dim rstLandlord As New ADODB.Recordset
   Dim szSQL As String, i As Integer
   
   flxLandlordGrid.Rows = 1
'again I am sticking with this SQL because PropertyLandlord table contains relationship
'   szSQL = "SELECT Landlord.LandlordID, Landlord.LandlordName, Landlord.LandlordAddressLine1, " & _
'               "Landlord.LandlordAddressLine2, Landlord.LandlordAddressLine3, " & _
'               "Landlord.LandlordPostCode, Landlord.LandlordHomeTel, " & _
'               "Landlord.LandlordMobile, Landlord.AcBalance " & _
'           "FROM PropertyLandlord, Landlord " & _
'           "WHERE PropertyLandlord.LandlordID = Landlord.LandlordID AND " & _
'               "PropertyLandlord.PropertyID = '" & txtPropertyID.text & "';"

'Debug.Print szSQL
'modfied by anol 27 aug 2015
 szSQL = "SELECT Supplier.SupplierID, Supplier.SupplierName, Supplier.SupplierAddressLine1, " & _
               "Supplier.SupplierAddressLine2, Supplier.SupplierAddressLine3, " & _
               "Supplier.SupplierPostCode, Supplier.SupplierHomeTel, " & _
               "Supplier.SupplierMobile, Supplier.AcBalance " & _
           "FROM PropertyLandlord, Supplier " & _
           "WHERE PropertyLandlord.LandlordID = Supplier.SupplierID AND Supplier.Type='LLORD' AND " & _
               "PropertyLandlord.PropertyID = '" & txtPropertyID.text & "';"
   rstLandlord.Open szSQL, conLandlord, adOpenStatic, adLockReadOnly

   i = 1
   While Not rstLandlord.EOF
      If Not rstLandlord.EOF Then flxLandlordGrid.AddItem ""
      flxLandlordGrid.TextMatrix(i, 0) = IIf(IsNull(rstLandlord!SupplierName), "", rstLandlord!SupplierName)
      flxLandlordGrid.TextMatrix(i, 1) = IIf(IsNull(rstLandlord!SupplierAddressLine1), "", rstLandlord!SupplierAddressLine1)
      flxLandlordGrid.TextMatrix(i, 2) = IIf(IsNull(rstLandlord!SupplierAddressLine2), "", rstLandlord!SupplierAddressLine2)
      flxLandlordGrid.TextMatrix(i, 3) = IIf(IsNull(rstLandlord!SupplierAddressLine3), "", rstLandlord!SupplierAddressLine3)
      flxLandlordGrid.TextMatrix(i, 4) = IIf(IsNull(rstLandlord!SupplierPostCode), "", rstLandlord!SupplierPostCode)
      flxLandlordGrid.TextMatrix(i, 5) = IIf(IsNull(rstLandlord!SupplierHomeTel), "", rstLandlord!SupplierHomeTel)
      flxLandlordGrid.TextMatrix(i, 6) = IIf(IsNull(rstLandlord!SupplierMobile), "", rstLandlord!SupplierMobile)
      flxLandlordGrid.TextMatrix(i, 7) = IIf(IsNull(rstLandlord!AcBalance), "", rstLandlord!AcBalance)
      flxLandlordGrid.TextMatrix(i, 8) = IIf(IsNull(rstLandlord!SupplierID), "", rstLandlord!SupplierID)
      i = i + 1
      rstLandlord.MoveNext
   Wend
   rstLandlord.Close
   Set rstLandlord = Nothing
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''   Me.MousePointer = vbArrow
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call WheelUnHook(Me.hWnd)
   LOAD_PROPERTY_PROPERTYID = ""
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub gridInsurance_Click()
   InsuranceButtonMode GridRowOnSelection
End Sub

Private Sub gridInsurance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridInsurance.ToolTipText = gridInsurance.TextMatrix(gridInsurance.MouseRow, gridInsurance.MouseCol)
End Sub

Private Sub gridInsurance_RowColChange()
   populateControl Me, gridInsurance
End Sub

Private Sub gridMaintenanceHistory_Click()
   If (gridMaintenanceHistory.row > 0 And gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) <> "") Then
      cmdEditMHistory.Enabled = True
   Else
      cmdEditMHistory.Enabled = False
   End If
End Sub

Private Sub gridMaintenanceHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridMaintenanceHistory.ToolTipText = gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.MouseRow, gridMaintenanceHistory.MouseCol)
End Sub

Private Sub gridMaintenanceHistory_RowColChange()
   populateControl Me, gridMaintenanceHistory
   If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 0) = "JOB" Then
      If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 25) = "True" Then
         cmdEmailJS_PO.Enabled = True
         cmdPrintJobSheet.Enabled = True
      Else
         If gridMaintenanceHistory.TextMatrix(gridMaintenanceHistory.row, 23) = "True" Then
            cmdEmailJS_PO.Enabled = True
            cmdPrintJobSheet.Enabled = True
         Else
            cmdEmailJS_PO.Enabled = False
            cmdPrintJobSheet.Enabled = False
         End If
      End If
   End If
End Sub

Private Sub gridPropertyAnalysis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridPropertyAnalysis.ToolTipText = gridPropertyAnalysis.TextMatrix(gridPropertyAnalysis.MouseRow, gridPropertyAnalysis.MouseCol)
End Sub

Private Sub gridSafety_Click()
   HealthSafetyButtonMode GridRowOnSelection
End Sub

Private Sub gridSafety_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridSafety.ToolTipText = gridSafety.TextMatrix(gridSafety.MouseRow, gridSafety.MouseCol)
End Sub

Private Sub gridSafety_RowColChange()
   populateControl Me, gridSafety
End Sub

Private Sub LoadProertyByPropertyID()
   SEARCHPropertyMODE_ = False

   '' LOAD MAIN Property INFORMATION

   fmeLoading.Visible = True
   fmeLoading.Refresh
   'crash after this line
   PopulatePropertyInformation LOAD_PROPERTY_PROPERTYID

   IMAGE_FILE_NAME_ = ImageLoader(imgPropertyPicture, txtPropertyID.text, "Property", lblImageName)

   '' LOAD Property DETAIL INFORMATION

   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString

   PopulateGridPropertyAnalysis adoConn

   LoadGridMaintenanceHistory adoConn
   PopulateHealthSafety adoConn
   PopulateInsurance adoConn
   PopulateUtilities adoConn

   fmeLoading.Visible = False
   adoConn.Close
   Set adoConn = Nothing

   ' SET OTHERS
   fmePropertyLookup.Visible = False
   SEARCHPropertyMODE_ = True
   tabProperty.Enabled = True

   gridPropertyAnalysis.row = 0
   gridPropertyAnalysis.col = 0
   gridPropertyAnalysis.SetFocus
End Sub
Private Sub TotalareofProperty(strPropertyID As String, adoConn As ADODB.Connection)
    'Fixed by anol 20160916
   
    Dim rsCheck As New ADODB.Recordset
    Dim rRow As Double
    rsCheck.Open "Select Sum(U.TotalArea)as SA from Units U where PropertyID='" & strPropertyID & "'", adoConn, adOpenKeyset, adLockReadOnly
    rRow = 0
    If Not rsCheck.EOF Then
        rRow = IIf(IsNull(rsCheck.Fields("SA").Value), 0, rsCheck.Fields("SA").Value)
    End If
    adoConn.Execute "Update Property set TotalArea=" & rRow & " where PropertyID='" & strPropertyID & "'"
     'flxFinancialYear.TextMatrix(flxFinancialYear.row, 6)
    rsCheck.Close
    
    'End of addition
End Sub

Private Sub gridPropertyLookup_Click()
    SEARCHPropertyMODE_ = False
    
    '    LOAD MAIN Property INFORMATION
    fmeProperty.Enabled = True
    tabProperty.Enabled = True
    
    fmeLoading.Visible = True
    fmeLoading.Refresh
    
    '  Supplementary
    '   ClearSupplementary
    txtClientList.Tag = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 5)
    txtClientList.text = gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 6)
    'crash after this line
    PopulatePropertyInformation gridPropertyLookup.TextMatrix(gridPropertyLookup.row, 0)
   
    
    IMAGE_FILE_NAME_ = ImageLoader(imgPropertyPicture, txtPropertyID.text, "Property", lblImageName)
    
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
     ' fixing the total area of an propery is not updated correcly when in putten in the unit
    
    
    
    PopulateGridPropertyAnalysis adoConn
    
    LoadGridMaintenanceHistory adoConn
    PopulateHealthSafety adoConn
    PopulateInsurance adoConn
    PopulateUtilities adoConn
    loadflxLandlordGrid adoConn
    RetrieveMemo adoConn
'    SupplierAccountBalance adoConn
    
    fmeLoading.Visible = False
    adoConn.Close
    Set adoConn = Nothing
    
    '    SET OTHERS
    fmePropertyLookup.Visible = False
    SEARCHPropertyMODE_ = True
    tabProperty.Enabled = True
    
    '   gridPropertyAnalysis.row = 0
    '   gridPropertyAnalysis.col = 0
    gridPropertyAnalysis.Enabled = True
    txtPropertyName.SetFocus
    
    ' Me.Caption = cboClientID.Column(1) + " - " + txtPropertyName.text
End Sub
Private Sub SupplierAccountBalance(adoConn As ADODB.Connection)
   Dim szSQL As String, i As Integer, iIndex As Integer
   Dim adoPayDr As New ADODB.Recordset, adoPayCr As New ADODB.Recordset

   szSQL = "SELECT COUNT(SageAccountNumber) AS X " & _
           "From " & _
            "(" & _
             "SELECT tlbPayment.SageAccountNumber  " & _
             "From tlbPayment " & _
             "GROUP BY tlbPayment.SageAccountNumber" & _
            ");"
   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoPayDr.EOF Then
      adoPayDr.Close
      Set adoPayDr = Nothing
      Exit Sub
   End If

   ReDim szaSupplierBalance(1, adoPayDr.Fields.Item(0).Value) As String
   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Dr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type = 6 OR Type = 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayDr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iIndex = 0
   While Not adoPayDr.EOF
      szaSupplierBalance(0, iIndex) = adoPayDr.Fields.Item("SageAccountNumber").Value
      szaSupplierBalance(1, iIndex) = adoPayDr.Fields.Item("Dr").Value
      iIndex = iIndex + 1
      adoPayDr.MoveNext
   Wend

   adoPayDr.Close

   szSQL = "SELECT SageAccountNumber, SUM(Amount) AS Cr " & _
           "FROM tlbPayment AS Pay " & _
           "WHERE Type <> 6 AND Type <> 24 " & _
           "GROUP BY SageAccountNumber;"

   adoPayCr.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   While Not adoPayCr.EOF
      For i = 0 To iIndex - 1
         If szaSupplierBalance(0, i) = adoPayCr.Fields.Item("SageAccountNumber").Value Then
            Exit For
         End If
      Next i
      If i < iIndex Then
         szaSupplierBalance(1, i) = szaSupplierBalance(1, i) - Val(adoPayCr.Fields.Item("Cr").Value)
      Else
         iIndex = iIndex + 1
         szaSupplierBalance(0, iIndex) = adoPayCr.Fields.Item("Cr").Value
      End If
      adoPayCr.MoveNext
   Wend

   adoPayCr.Close

   Set adoPayDr = Nothing
   Set adoPayCr = Nothing
End Sub
Public Sub RetrieveMemo(ByVal conMemo_ As ADODB.Connection)
   Dim rstMemo_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   On Error Resume Next

   txtMemo.text = ""

   sSQLQuery_ = "SELECT MemoText " & _
                "FROM Property WHERE PropertyID = '" & txtPropertyID.text & "'"
   rstMemo_.Open sSQLQuery_, conMemo_, adOpenStatic, adLockReadOnly

   txtMemo.text = rstMemo_!MemoText 'IIf(IsNull(rstMemo_!MemoText), "<No memo saved>", rstMemo_!MemoText)
   
   Call LoadAttachmentFiles(cmbFiles, txtPropertyID.text, "Property")

   rstMemo_.Close
   Set rstMemo_ = Nothing
End Sub
'
'Private Sub ClearSupplementary()
'   txtSupp1.text = ""
'   txtSupp2.text = ""
'   txtSupp3.text = ""
'   txtDtFlgDate.text = ""
'   txtDtFlgDt2.text = ""
'   txtDtFlgDt3.text = ""
'   txtDtFlgDesc.text = ""
'   txtDtFlgDesc2.text = ""
'   txtDtFlgDesc3.text = ""
'End Sub
Private Sub gridPropertyLookup_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gridPropertyLookup_Click
   End If
End Sub

Private Sub gridUtilities_Click()
   UtilitiesButtonMode GridRowOnSelection
End Sub

Private Sub gridUtilities_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   gridUtilities.ToolTipText = gridUtilities.TextMatrix(gridUtilities.MouseRow, gridUtilities.MouseCol)
End Sub

Private Sub gridUtilities_RowColChange()
   populateControl Me, gridUtilities
End Sub
'
'Private Sub lblSupplementary1_DblClick()
'   txtSuppCaption1.Visible = True
'   txtSuppCaption1.Left = lblSupplementary1.Left
'   txtSuppCaption1.Top = lblSupplementary1.Top
'   txtSuppCaption1.text = lblSupplementary1.Caption
'   txtSuppCaption1.SetFocus
'End Sub
'
'Private Sub lblSupplementary2_Click()
'   txtSuppCaption2.Visible = True
'   txtSuppCaption2.Left = lblSupplementary2.Left
'   txtSuppCaption2.Top = lblSupplementary2.Top
'   txtSuppCaption2.text = lblSupplementary2.Caption
'   txtSuppCaption2.SetFocus
'End Sub
'
'Private Sub lblSupplementary3_DblClick()
'   txtSuppCaption3.Visible = True
'   txtSuppCaption3.Left = lblSupplementary3.Left
'   txtSuppCaption3.Top = lblSupplementary3.Top
'   txtSuppCaption3.text = lblSupplementary3.Caption
'   txtSuppCaption3.SetFocus
'End Sub

Private Sub optAll_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget / Location"
   Label61(3).Caption = "Ref"
   Label61(1).Caption = "Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
End Sub

Private Sub optDiary_Click()
   Dim i As Integer

   Label61(10).Caption = "Location"
   Label61(3).Caption = "Diary No"
   Label61(1).Caption = "Event Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) = "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

Private Sub optJobs_Click()
   Dim i As Integer

   Label61(10).Caption = "Budget"
   Label61(3).Caption = "Job No"
   Label61(1).Caption = "Maintenance Type"
'MsgBox gridMaintenanceHistory.RowHeight(3)
   For i = 1 To gridMaintenanceHistory.Rows - 1
      gridMaintenanceHistory.RowHeight(i) = 240
   Next i
   For i = 1 To gridMaintenanceHistory.Rows - 1
      If gridMaintenanceHistory.TextMatrix(i, 0) <> "JOB" Then
         gridMaintenanceHistory.RowHeight(i) = 0
      Else
         gridMaintenanceHistory.RowHeight(i) = 240
      End If
   Next i
End Sub

'Private Sub tabProperty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   tabProperty.MousePointer = vbArrow
'End Sub

Private Sub txtAnalysisPercentage_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Added By Samrat. 12/10/2006
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtAnalysisPercentage, KA, 4
   KeyAscii = KA
End Sub

Private Sub txtAnalysisQuantity_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Added By Samrat. 12/10/2006
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtAnalysisQuantity, KA
   KeyAscii = KA
End Sub

Private Sub txtAnalysisValue_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'Added By Samrat. 12/10/2006
   Dim KA As Integer
   KA = KeyAscii
   DigitTextKeyPress txtAnalysisValue, KA
   KeyAscii = KA
End Sub
'
'Private Sub txtAnnualPR_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   'Added By Samrat. 12/10/2006
'   Dim KA As Integer
'   KA = KeyAscii
'   DigitTextKeyPress txtAnnualPR, KA
'   KeyAscii = KA
'End Sub

Private Sub txtChargeRate_KeyPress(KeyAscii As Integer)
   DigitTextKeyPress txtChargeRate, KeyAscii
End Sub

Private Sub txtDateChk_Change()
   TextBoxChangeDate txtDateChk
End Sub

Private Sub txtDateChk_GotFocus()
   SelTxtInCtrl txtDateChk
End Sub

Private Sub txtDateChk_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateChk, KeyAscii
End Sub

Private Sub txtDateChk_LostFocus()
   TextBoxFormatDate txtDateChk
End Sub

Private Sub txtDateVacated_Change()
   TextBoxChangeDate txtDateVacated
End Sub

Private Sub txtDateVacated_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDateVacated, KeyAscii
End Sub

Private Sub txtDateVacated_LostFocus()
   TextBoxFormatDate txtDateVacated
End Sub
'
'Private Sub txtDtFlgDate_Change()
'   TextBoxChangeDate txtDtFlgDate
'End Sub
'
'Private Sub txtDtFlgDate_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDate, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDate_LostFocus()
'   TextBoxFormatDate txtDtFlgDate
'End Sub
'
'Private Sub txtDtFlgDt2_Change()
'   TextBoxChangeDate txtDtFlgDt2
'End Sub
'
'Private Sub txtDtFlgDt2_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDt2, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDt2_LostFocus()
'   TextBoxFormatDate txtDtFlgDt2
'End Sub
'
'Private Sub txtDtFlgDt3_Change()
'   TextBoxChangeDate txtDtFlgDt3
'End Sub
'
'Private Sub txtDtFlgDt3_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtDtFlgDt3, KeyAscii
'End Sub
'
'Private Sub txtDtFlgDt3_LostFocus()
'   TextBoxFormatDate txtDtFlgDt3
'End Sub

Private Sub txtExpiryDate_Change()
   TextBoxChangeDate txtExpiryDate
End Sub

Private Sub txtExpiryDate_GotFocus()
   SelTxtInCtrl txtExpiryDate
End Sub

Private Sub txtExpiryDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtExpiryDate, KeyAscii
End Sub

Private Sub txtExpiryDate_LostFocus()
   TextBoxFormatDate txtExpiryDate
End Sub

Private Sub txtNextDueDate_Change()
   TextBoxChangeDate txtNextDueDate
End Sub

Private Sub txtNextDueDate_GotFocus()
   SelTxtInCtrl txtNextDueDate
End Sub

Private Sub txtNextDueDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtNextDueDate, KeyAscii
End Sub

Private Sub txtNextDueDate_LostFocus()
   TextBoxFormatDate txtNextDueDate
End Sub

Private Sub txtPropertyID_LostFocus()
   If txtPropertyID.text <> "" Then txtPropertyID.text = UCase(txtPropertyID.text)

   If NEWMODE_ Then
      If CheckPropertyID Then
         txtPropertyID.text = ""
         txtPropertyID.SetFocus
         ShowMsgInTaskBar "Property ID already exits. Please enter unique property id.", , "N"
      End If
   Else
      If txtPropertyID.text <> "" And txtPropertyID.text <> szEditingPropID Then
         If CheckPropertyID Then
            txtPropertyID.text = ""
            txtPropertyID.SetFocus
            ShowMsgInTaskBar "Property ID already exits. Please enter unique property id.", , "N"
         End If
      End If
   End If
End Sub

Private Sub txtPropertyName_LostFocus()
   If txtPropertyName.text = "" Then Exit Sub

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

   If UBound(szaChoice) > 0 Then
      If szaChoice(4) <> "" Then
         If InStr(szaChoice(4), "P") > 0 Then
'            If NEWMODE_ And txtPropertyID.text = "" Then txtPropertyID.text = GeneratePropertyID
         End If
      End If
   End If
End Sub
'
'Private Sub txtRechargeableAmount_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   'Added By Samrat. 12/10/2006
'   Dim KA As Integer
'   KA = KeyAscii
'   DigitTextKeyPress txtRechargeableAmount, KA
'   KeyAscii = KA
'End Sub

'Private Sub txtPropertySearch_Change()
''    If Not SEARCHPropertyMODE_ Then
''       Exit Sub
''   End If
''   Dim sFilter_ As String
''   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
''                 "ORDER BY PropertyID;"
''   PopulatePropertyLookup sFilter_
'End Sub

'Private Sub txtSearchProperty_Change()
'   If Not SEARCHPropertyMODE_ Then
'       Exit Sub
'   End If
'   Dim sFilter_ As String
'   sFilter_ = "WHERE PropertyID LIKE '" & Trim(txtSearchProperty.text) & "%' " & _
'                 "ORDER BY PropertyID;"
'   PopulatePropertyLookup sFilter_
'   'gridPropertyLookup.Top = 360
'End Sub

Private Sub PrepareList()
   configflxSupplierList flxSupplierList
   LoadAllSupplierFlxGrd
'   UpdateBalance
End Sub
Private Sub configflxSupplierList(conFlxGrid As Control)
   Dim szHeader As String
   
   conFlxGrid.Clear
   conFlxGrid.Cols = 5
   szHeader$ = "|<LandlordID|<LandlordName|<LandlordPostCode|>AccBalance"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 220          'Solid column
   conFlxGrid.ColWidth(1) = 1000       'Supplier ID
   conFlxGrid.ColWidth(2) = 3000       'Supplier Name
   conFlxGrid.ColWidth(3) = 0          'Post Code
   conFlxGrid.ColWidth(4) = 1100       'Account Balance
   conFlxGrid.Rows = 2
'
   'conFlxGrid.RowHeightMin = 300
   conFlxGrid.RowHeight(0) = 0
End Sub
Private Sub LoadAllSupplierFlxGrd()
   Dim adoConn As New ADODB.Connection
   Dim rstRst As New ADODB.Recordset
   Dim szSQL As String

   'On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   adoConn.Open getConnectionString

   szSQL = "SELECT SupplierID, SupplierName, SupplierPostCode " & _
           "FROM Supplier " & _
           "WHERE TYPE = 'LLORD' " & _
           "ORDER BY SupplierName;"

   rstRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRst.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1
   flxSupplierList.TextMatrix(iRow, 1) = ""
   flxSupplierList.TextMatrix(iRow, 2) = ""
   flxSupplierList.AddItem ""
   iRow = 2
   While Not rstRst.EOF
      flxSupplierList.TextMatrix(iRow, 1) = rstRst!SupplierID
      flxSupplierList.TextMatrix(iRow, 2) = rstRst!SupplierName
      flxSupplierList.TextMatrix(iRow, 3) = IIf(IsNull(rstRst!SupplierPostCode), "", rstRst!SupplierPostCode)
      rstRst.MoveNext
      If Not rstRst.EOF Then flxSupplierList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstRst.Close
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   adoConn.Close
   Set rstRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub UpdateBalance()
   Dim i As Integer, j As Integer

   For i = 1 To flxSupplierList.Rows - 1
      For j = 0 To UBound(szaSupplierBalance, 2) - 1
         If flxSupplierList.TextMatrix(i, 1) = szaSupplierBalance(0, j) Then
            flxSupplierList.TextMatrix(i, 4) = Format(szaSupplierBalance(1, j), "0.00")
            Exit For
         End If
      Next j
      If j = UBound(szaSupplierBalance, 2) Then flxSupplierList.TextMatrix(i, 4) = "0.00"
   Next i
End Sub
Public Sub PopulateCodes()
   Dim sSQLQuery As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   ' Analysis Type
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'ATYP'"

   populateCombo adoConn, sSQLQuery, cboAnalysisType

   ' Select option
   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'SOPT'"

   populateCombo adoConn, sSQLQuery, cboAnalysisOption
'
'
'   ' Safety Type
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'STYP'"
'
'   populateCombo adoConn, sSQLQuery, cboSafetyType

   adoSafetyType.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'STYP'"

   adoSafetyType.RecordSource = sSQLQuery
   adoSafetyType.CommandType = adCmdText
   adoSafetyType.Refresh

   adoSafetyStatus.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'SSTA'"
   adoSafetyStatus.RecordSource = sSQLQuery
   adoSafetyStatus.CommandType = adCmdText
   adoSafetyStatus.Refresh

   adoInspector.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IPT'"
   adoInspector.RecordSource = sSQLQuery
   adoInspector.CommandType = adCmdText
   adoInspector.Refresh

'   ' Safety Status
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'SSTA'"
'
'   populateCombo adoConn, sSQLQuery, cboSafetyStatus
'
   ' Insurance Type
   adoInsurer.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'IRER'"
   adoInsurer.RecordSource = sSQLQuery
   adoInsurer.CommandType = adCmdText
   adoInsurer.Refresh

   adoInsuranceType.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'ITYP'"
   adoInsuranceType.RecordSource = sSQLQuery
   adoInsuranceType.CommandType = adCmdText
   adoInsuranceType.Refresh

   adoInsUsage.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                "FROM SECONDARYCODE " & _
                "WHERE PRIMARYCODE = 'UUSE'"
   adoInsUsage.RecordSource = sSQLQuery
   adoInsUsage.CommandType = adCmdText
   adoInsUsage.Refresh

   ' Clients
'   sSQLQuery = "SELECT CLIENTID, CLIENTNAME " & _
'                 "FROM CLIENT "
'
'   populateCombo adoConn, sSQLQuery, cboClientID
'modified by anol 27 Aug 2015
'   sSQLQuery = "SELECT SupplierID, SupplierName " & _
'                "FROM Supplier where Type='LLORD';"
'
'   populateCombo adoConn, sSQLQuery, cboLandLord

   adoConn.Close
   Set adoConn = Nothing

   adoUtilitiesType.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'UTIL'"

   adoUtilitiesType.RecordSource = sSQLQuery
   adoUtilitiesType.CommandType = adCmdText
   adoUtilitiesType.Refresh

   adoSupplier.ConnectionString = getConnectionString

   sSQLQuery = "SELECT SupplierID, SupplierName " & _
                 "FROM SUPPLIER "

   adoSupplier.RecordSource = sSQLQuery
   adoSupplier.CommandType = adCmdText
   adoSupplier.Refresh

   adoStatus.ConnectionString = getConnectionString

   sSQLQuery = "SELECT CODE, VALUE " & _
                 "FROM SECONDARYCODE " & _
                 "WHERE PRIMARYCODE = 'USTA'"

   adoStatus.RecordSource = sSQLQuery
   adoStatus.CommandType = adCmdText
   adoStatus.Refresh


'   ' Maintenance Type
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'MTYP'"
'
'   populateCombo adoConn, sSQLQuery, cboMaintenanceType
'
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'MNTJOB'"
'   populateCombo adoConn, sSQLQuery, cboTaskOwner
'   populateCombo adoConn, sSQLQuery, cboContact
'
'   ' Safety Type
'   sSQLQuery = "SELECT CODE, VALUE " & _
'                 "FROM SECONDARYCODE " & _
'                 "WHERE PRIMARYCODE = 'STYP'"
'
'   populateCombo adoConn, sSQLQuery, cboSafetyType
'
End Sub

Public Function PopulatePropertyInformation(ByVal sPropertyNumber As String) As Boolean
   Dim conProperty_ As New ADODB.Connection
   Dim rstProperty_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   
   On Error GoTo ErrorHandler
   
   'Set the ado Connections to the dataset
   conProperty_.Open getConnectionString
   
   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT * FROM Property WHERE PropertyID = '" & sPropertyNumber & "'"
   
   rstProperty_.Open sSQLQuery_, conProperty_, adOpenStatic, adLockReadOnly
   
   'MsgBox sPropertyNumber
   If rstProperty_.EOF Or rstProperty_.BOF Then
       ShowMsgInTaskBar "WARNING !! No information found for the specified Property.", , "N"
   End If

   With rstProperty_
      While Not .EOF
         'Resolved by BOSL
         'Modified by anol 08 Dec 2014
         'issue 495
         If Not IsNull(!propertyID) Then
            txtClientList.Tag = !ClientID
         End If
         txtPropertyID.text = !propertyID
         txtPropertyName.text = !PropertyName
         If Not IsNull(!ProAddressLine1) Then
            txtProAddressLine1.text = !ProAddressLine1
         End If
         
         If Not IsNull(!ProAddressLine2) Then
            txtProAddressLine2.text = !ProAddressLine2
         End If
         
         If Not IsNull(!ProAddressLine3) Then
            txtProAddressLine3.text = !ProAddressLine3
         End If
         
         If Not IsNull(!ProAddressLine4) Then
            txtProAddressLine4.text = !ProAddressLine4
         End If
         
         txtProPostCode.text = IIf(IsNull(!PROPOSTCODE), "", !PROPOSTCODE)
         txtAnalysisTotalArea.text = IIf(IsNull(!TotalArea), "0.00", !TotalArea)
         txtManager.text = IIf(IsNull(!Manager), "", !Manager)
         txtContactDetails.text = IIf(IsNull(!ContactDetails), "", !ContactDetails)
         
'         lblSupplementary1.Caption = IIf(IsNull(.Fields.Item("SuppCaption1").Value), "Supplementary 1:", .Fields.Item("SuppCaption1").Value)
'         lblSupplementary2.Caption = IIf(IsNull(.Fields.Item("SuppCaption2").Value), "Supplementary 2:", .Fields.Item("SuppCaption2").Value)
'         lblSupplementary3.Caption = IIf(IsNull(.Fields.Item("SuppCaption3").Value), "Supplementary 3:", .Fields.Item("SuppCaption3").Value)
'         txtSupp1.text = IIf(IsNull(.Fields.Item("SuppText1").Value), "", .Fields.Item("SuppText1").Value)
'         txtSupp2.text = IIf(IsNull(.Fields.Item("SuppText2").Value), "", .Fields.Item("SuppText2").Value)
'         txtSupp3.text = IIf(IsNull(.Fields.Item("SuppText3").Value), "", .Fields.Item("SuppText3").Value)
'         txtDtFlgDate.text = IIf(IsNull(.Fields.Item("DateFlagDt1").Value), "", .Fields.Item("DateFlagDt1").Value)
'         txtDtFlgDt2.text = IIf(IsNull(.Fields.Item("DateFlagDt2").Value), "", .Fields.Item("DateFlagDt2").Value)
'         txtDtFlgDt3.text = IIf(IsNull(.Fields.Item("DateFlagDt3").Value), "", .Fields.Item("DateFlagDt3").Value)
'         txtDtFlgDesc.text = IIf(IsNull(.Fields.Item("DateFlagDescription1").Value), "", .Fields.Item("DateFlagDescription1").Value)
'         txtDtFlgDesc2.text = IIf(IsNull(.Fields.Item("DateFlagDescription2").Value), "", .Fields.Item("DateFlagDescription2").Value)
'         txtDtFlgDesc3.text = IIf(IsNull(.Fields.Item("DateFlagDescription3").Value), "", .Fields.Item("DateFlagDescription3").Value)
'
         
         .MoveNext
      Wend
   
      rstProperty_.Close
   End With
   conProperty_.Close
   Set rstProperty_ = Nothing
   Set conProperty_ = Nothing
   
   Exit Function
ErrorHandler:
   ShowMsgInTaskBar Err.Number & " " & Err.description, , ""
   
   rstProperty_.Close
   conProperty_.Close
   Set rstProperty_ = Nothing
   Set conProperty_ = Nothing
End Function

Public Function SavePropertyInformation() As Boolean
   Dim conProperty_ As New ADODB.Connection
   Dim rstProperty_ As New ADODB.Recordset
   Dim sSQLQuery_ As String
   
   If NEWMODE_ Then
       'Set the ado Connections to the dataset
       conProperty_.Open getConnectionString
   
       sSQLQuery_ = "SELECT ClientID, PropertyID, PropertyName, ProAddressLine1, ProAddressLine2, " & _
                     "ProAddressLine3, ProAddressLine4, ProPostCode, TotalArea " & _
                    "FROM Property " & _
                    "WHERE PropertyID = '" & txtPropertyID.text & "'"
   
       rstProperty_.Open sSQLQuery_, conProperty_, adOpenStatic, adLockReadOnly
   
       If rstProperty_.EOF Or rstProperty_.BOF Then
           If (AddPropertyInformation) Then
               SavePropertyInformation = True
           Else
               SavePropertyInformation = False
           End If
       Else
           If (MsgBox("WARNING ! The Property Number entered already exists. Do you want to update the information", vbYesNo, "Save Property Information") = vbYes) Then
               If UpdatePropertyInformation Then
                   SavePropertyInformation = True
               Else
                   ShowMsgInTaskBar "An error occured while updating the Property Information.", , "N"
                   SavePropertyInformation = False
               End If
           End If
       End If

       rstProperty_.Close
       conProperty_.Close
       Set rstProperty_ = Nothing
       Set conProperty_ = Nothing
       SavePropertyInformation = True
       Exit Function
   Else
      If UpdatePropertyInformation Then
          SavePropertyInformation = True
          Exit Function
      Else
          SavePropertyInformation = False
          Exit Function
      End If
   End If

   Exit Function
Exception:
       ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
       rstProperty_.Close
       conProperty_.Close
       Set rstProperty_ = Nothing
       Set conProperty_ = Nothing
       SavePropertyInformation = False
End Function

Public Function AddPropertyInformation() As Boolean

    Dim conProperty_ As New ADODB.Connection
    Dim rstProperty_ As New ADODB.Recordset
    Dim sSQLQuery_ As String

    On Error GoTo Exception
    'Set the ado Connections to the dataset
    conProperty_.Open getConnectionString

    sSQLQuery_ = "SELECT ClientID, PropertyID, PropertyName, ProAddressLine1, ProAddressLine2, " & _
                  "ProAddressLine3, ProAddressLine4, ProPostCode, TotalArea,CreatedBy,CreatedDate " & _
    "FROM Property"

    rstProperty_.Open sSQLQuery_, conProperty_, adOpenDynamic, adLockOptimistic
    rstProperty_.AddNew

    rstProperty_!CreatedBy = User
    rstProperty_!CreatedDate = Now
    rstProperty_!propertyID = txtPropertyID.text
    rstProperty_!ClientID = txtClientList.Tag
    rstProperty_!PropertyName = txtPropertyName.text
    rstProperty_!ProAddressLine1 = txtProAddressLine1.text
    rstProperty_!ProAddressLine2 = txtProAddressLine2.text
    rstProperty_!ProAddressLine3 = txtProAddressLine3.text
    rstProperty_!ProAddressLine4 = txtProAddressLine4.text
    rstProperty_!PROPOSTCODE = txtProPostCode.text
    
    If Not txtAnalysisTotalArea.text = "" Then
        rstProperty_!TotalArea = CDbl(txtAnalysisTotalArea.text)
    End If
    
    rstProperty_.Update
    rstProperty_.Close
    conProperty_.Close
    Set rstProperty_ = Nothing
    Set conProperty_ = Nothing

    AddPropertyInformation = True
    Exit Function

Exception:
    ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
    rstProperty_.Close
    conProperty_.Close
    Set rstProperty_ = Nothing
    Set conProperty_ = Nothing
    AddPropertyInformation = False
End Function

Public Function UpdatePropertyInformation() As Boolean
   Dim adoPropConn As New ADODB.Connection
   Dim rstProperty_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   On Error GoTo Exception
   'Set the ado Connections to the dataset
   adoPropConn.Open getConnectionString

   If txtPropertyID.text <> szEditingPropID Then _
      UpdatePropertyID adoPropConn, txtPropertyID.text, szEditingPropID

   sSQLQuery_ = "SELECT * FROM Property WHERE PropertyID = '" & txtPropertyID.text & "'"

   rstProperty_.Open sSQLQuery_, adoPropConn, adOpenDynamic, adLockOptimistic

   rstProperty_!ClientID = txtClientList.Tag
   rstProperty_!PropertyName = txtPropertyName.text
   rstProperty_!ProAddressLine1 = txtProAddressLine1.text
   rstProperty_!ProAddressLine2 = txtProAddressLine2.text
   rstProperty_!ProAddressLine3 = txtProAddressLine3.text
   rstProperty_!ProAddressLine4 = txtProAddressLine4.text
   rstProperty_!PROPOSTCODE = txtProPostCode.text
   rstProperty_!Manager = txtManager.text
   rstProperty_!ContactDetails = txtContactDetails.text
   rstProperty_!LastModifiedBy = User
   rstProperty_!LastModifiedDate = Now
    
   If Not txtAnalysisTotalArea.text = "" Then
       rstProperty_!TotalArea = CDbl(txtAnalysisTotalArea.text)
   End If
   
   rstProperty_.Update

   rstProperty_.Close
   adoPropConn.Close
   Set rstProperty_ = Nothing
   Set adoPropConn = Nothing
   UpdatePropertyInformation = True
   Exit Function

Exception:

   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
   rstProperty_.Close
   adoPropConn.Close
   Set rstProperty_ = Nothing
   Set adoPropConn = Nothing
   UpdatePropertyInformation = False
End Function

Private Sub UpdatePropertyID(adoConn As ADODB.Connection, szNewID As String, szOldID As String)
   Dim sSQLQuery As String

   sSQLQuery = "UPDATE Property " & _
               "SET PropertyID = '" & szNewID & "' " & _
               "Where PropertyID = '" & szOldID & "'"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE AttachedFile " & _
               "SET OwnerID ='" & szNewID & "' " & _
               "WHERE OwnerID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientGlobalData " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE ClientProAgr " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE DemandTypes " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalData " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalInsurance " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalRC " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE GlobalSC " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)

   sSQLQuery = "UPDATE InterestRates " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE NLPosting " & _
               "SET PROPERTY_ID ='" & szNewID & "' " & _
               "WHERE PROPERTY_ID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyAnalysis " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyInsurance " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyLandlord " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyMaintHistory " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertySafety " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE PropertyUtilities " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchPayment " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchReceipt " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblBatchTransaction " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tblPurInv " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tlbBankPayment " & _
               "SET PropertyID ='" & szNewID & "' " & _
               "WHERE PropertyID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   
   sSQLQuery = "UPDATE tlbPayment " & _
               "SET UnitID ='" & szNewID & "' " & _
               "WHERE UnitID = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)


   sSQLQuery = "UPDATE tlbPaymentSplit " & _
               "SET TRANS ='" & szNewID & "' " & _
               "WHERE TRANS = '" & szOldID & "';"
   adoConn.Execute (sSQLQuery)
   adoConn.Execute "Update Units Set PropertyID ='" & szNewID & "' where PropertyID='" & szOldID & "'"

End Sub

Public Sub MaintenanceHistoryButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
    Case ComponentMode.DefaultMode
        cmdNewMHistory.Enabled = True
        cmdEditMHistory.Enabled = False
        gridMaintenanceHistory.Enabled = True
        
    Case ComponentMode.GridRowOnSelection
        cmdNewMHistory.Enabled = True
        cmdEditMHistory.Enabled = True
        gridMaintenanceHistory.Enabled = True
    
    Case ComponentMode.NewEntryMode
        cmdNewMHistory.Enabled = False
        cmdEditMHistory.Enabled = False
        gridMaintenanceHistory.Enabled = False

    Case ComponentMode.EditMode
         cmdNewMHistory.Enabled = False
         cmdEditMHistory.Enabled = False
         gridMaintenanceHistory.Enabled = False
   End Select
End Sub

Public Sub PropertyAnalysisButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
        Case ComponentMode.DefaultMode
            cmdAnalysisNew.Enabled = True
            cmdAnalysisEdit.Enabled = True
            cmdAnalysisSave.Enabled = False
            cmdAnalysisCancel.Enabled = False
            
            gridPropertyAnalysis.Enabled = True
            cboAnalysisType.Enabled = True
            cboAnalysisType.Locked = True
            cmdAnalysis.Enabled = True
            txtAnalysisDescription.Enabled = True
            txtAnalysisDescription.Locked = True
            cboAnalysisOption.Enabled = True
            cboAnalysisOption.Locked = True
            txtAnalysisValue.Enabled = True
            txtAnalysisValue.Locked = True
            txtAnalysisQuantity.Enabled = True
            txtAnalysisQuantity.Locked = True
            txtAnalysisPercentage.Enabled = True
            txtAnalysisPercentage.Locked = True
            'added by anol 02 nov 2014
            txtAnalysisValue1.Enabled = True
            txtAnalysisValue1.Locked = True
            
            txtAnalysisReference.Enabled = True
            txtAnalysisReference.Locked = True
            txtAnalysisValue1.text = ""
            txtAnalysisReference.text = ""
            'End of modification
            
            cboAnalysisType.ListIndex = -1
            cboAnalysisOption.ListIndex = -1
            txtAnalysisDescription.text = ""
            txtAnalysisValue.text = ""
            txtAnalysisQuantity.text = ""
            txtAnalysisPercentage.text = ""
              
        Case ComponentMode.GridRowOnSelection
            cmdAnalysisNew.Enabled = True
            cmdAnalysisEdit.Enabled = True
            cmdAnalysisSave.Enabled = False
            cmdAnalysisCancel.Enabled = False
            
            gridPropertyAnalysis.Enabled = True
        
        Case ComponentMode.NewEntryMode
            cmdAnalysisNew.Enabled = False
            cmdAnalysisEdit.Enabled = False
            cmdAnalysisSave.Enabled = True
            cmdAnalysisCancel.Enabled = True
            gridPropertyAnalysis.Enabled = False
            cboAnalysisType.Locked = False
            cmdAnalysis.Enabled = True
            txtAnalysisDescription.Locked = False
            txtAnalysisDescription.text = ""
            cboAnalysisOption.Locked = False
            txtAnalysisValue.Locked = False
            txtAnalysisValue.text = ""
            txtAnalysisQuantity.Locked = False
            txtAnalysisQuantity.text = ""
            txtAnalysisPercentage.Locked = False
            txtAnalysisPercentage.text = ""
            'added by anol 02 nov 2014
            txtAnalysisValue1.Locked = False
            txtAnalysisValue1.text = ""
            txtAnalysisReference.Enabled = True
            txtAnalysisReference.Locked = False
            txtAnalysisReference.text = ""
            
        Case ComponentMode.EditMode
            cmdAnalysisNew.Enabled = False
            cmdAnalysisEdit.Enabled = False
            cmdAnalysisSave.Enabled = True
            cmdAnalysisCancel.Enabled = True
            gridPropertyAnalysis.Enabled = False
            cboAnalysisType.Locked = False
            cmdAnalysis.Enabled = False
            txtAnalysisDescription.Locked = False
            cboAnalysisOption.Locked = False
            txtAnalysisValue.Locked = False
            txtAnalysisQuantity.Locked = False
            txtAnalysisPercentage.Locked = False
            'added by anol 02 nov 2014
            txtAnalysisValue1.Locked = False
            txtAnalysisReference.Locked = False
    End Select

End Sub

Public Sub HealthSafetyButtonMode(ByVal mode As ComponentMode)
    Dim ctrl As Control
    Select Case mode
    
        Case ComponentMode.DefaultMode
            cmdSafetyNew.Enabled = True
            cmdSafetyEdit.Enabled = False
            cmdSafetySave.Enabled = False
            cmdSafetyCancel.Enabled = False
            cmdAttachment.Enabled = False

            gridSafety.Enabled = True

            cboSafetyType.Enabled = False
            cboSafetyType.text = ""
            cmdSafety.Enabled = False
            txtRef.Enabled = False
            txtRef.text = ""
            txtDateChk.Enabled = False
            txtDateChk.text = ""
            txtNextDueDate.Enabled = False
            txtNextDueDate.text = ""
            cboSchedule.Enabled = False
            cboSchedule.text = ""
            cboInspectedBy.Enabled = False
            cboInspectedBy.text = ""
            txtSafetyTelephone.Enabled = False
            txtSafetyTelephone.text = ""
            txtComment.Enabled = False
            txtComment.text = ""
            chkCertificate.Enabled = False
            chkCertificate.Value = 0
            chkAlarm.Enabled = False
            chkAlarm.Value = 0
        
        Case ComponentMode.GridRowOnSelection
            cmdSafetyNew.Enabled = True
            cmdSafetyEdit.Enabled = True
            cmdSafetySave.Enabled = False
            cmdSafetyCancel.Enabled = False
            cmdAttachment.Enabled = False

            gridSafety.Enabled = True

        Case ComponentMode.NewEntryMode
            cmdSafetyNew.Enabled = False
            cmdSafetyEdit.Enabled = False
            cmdSafetySave.Enabled = True
            cmdSafetyCancel.Enabled = True
            cmdAttachment.Enabled = True

            gridSafety.Enabled = False

            cboSafetyType.Enabled = True
            cboSafetyType.text = ""
            cmdSafety.Enabled = True
            txtRef.Enabled = True
            txtRef.text = ""
            txtNextDueDate.Enabled = True
            txtNextDueDate.text = ""
            txtDateChk.Enabled = True
            txtDateChk.text = "" 'Format(Date, "dd/mm/yyyy")
            txtNextDueDate.text = ""
            cboSchedule.Enabled = True
            cboSchedule.text = ""
            cboInspectedBy.Enabled = True
            cboInspectedBy.text = ""
            txtSafetyTelephone.Enabled = True
            txtSafetyTelephone.text = ""
            txtComment.Enabled = True
            txtComment.text = ""
            chkCertificate.Enabled = True
            chkAlarm.Enabled = True
            chkCertificate.Value = 0
            chkAlarm.Value = 0

        Case ComponentMode.EditMode
            cmdSafetyNew.Enabled = False
            cmdSafetyEdit.Enabled = False
            cmdSafetySave.Enabled = True
            cmdSafetyCancel.Enabled = True
            cmdAttachment.Enabled = True

            gridSafety.Enabled = False

            cboSafetyType.Enabled = True
            cmdSafety.Enabled = True
            txtRef.Enabled = True
            txtNextDueDate.Enabled = True
            txtDateChk.Enabled = True
            cboSchedule.Enabled = True
            cboInspectedBy.Enabled = True
            txtSafetyTelephone.Enabled = True
            txtComment.Enabled = True
            chkCertificate.Enabled = True
            chkAlarm.Enabled = True
    End Select
End Sub

Public Sub InsuranceButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control
   Select Case mode
      Case ComponentMode.DefaultMode
         cmdInsuranceNew.Enabled = True
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = False
         cmdInsuranceCancel.Enabled = False

         gridInsurance.Enabled = True
         fraInsurance.Enabled = False

         cboInsurer.text = ""
         cboInsuranceType.text = ""
         txtPolicyNo.text = ""
         txtSumInsured.text = ""
         txtAnnualPR.text = ""
         txtStartDate.text = ""
         txtExpiryDate.text = ""
         cboUsage.text = ""
         txtTelephone.text = ""
         txtComments.text = ""

         cboInsurer.Enabled = True
         cboInsuranceType.Enabled = True
         txtPolicyNo.Enabled = True
         txtSumInsured.Enabled = True
         txtAnnualPR.Enabled = True
         txtStartDate.Enabled = True
         txtExpiryDate.Enabled = True
         cboUsage.Enabled = True
         txtTelephone.Enabled = True
         txtComments.Enabled = True

      Case ComponentMode.GridRowOnSelection
         cmdInsuranceNew.Enabled = True
         cmdInsuranceEdit.Enabled = True
         cmdInsuranceSave.Enabled = False
         cmdInsuranceCancel.Enabled = False

         gridInsurance.Enabled = True

      Case ComponentMode.NewEntryMode
         cmdInsuranceNew.Enabled = False
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = True
         cmdInsuranceCancel.Enabled = True

         gridInsurance.Enabled = False
         fraInsurance.Enabled = True

         cboInsurer.text = ""
         cboInsuranceType.text = ""
         txtPolicyNo.text = ""
         txtSumInsured.text = ""
         txtAnnualPR.text = ""
         txtStartDate.text = ""
         txtExpiryDate.text = ""
         cboUsage.text = ""
         txtTelephone.text = ""
         txtComments.text = ""

      Case ComponentMode.EditMode
         cmdInsuranceNew.Enabled = False
         cmdInsuranceEdit.Enabled = False
         cmdInsuranceSave.Enabled = True
         cmdInsuranceCancel.Enabled = True

         gridInsurance.Enabled = False
         fraInsurance.Enabled = True
   End Select
End Sub

Public Sub UtilitiesButtonMode(ByVal mode As ComponentMode)
   Dim ctrl As Control

   Select Case mode
      Case ComponentMode.DefaultMode
         cmdUtilitiesNew.Enabled = True
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = False
         cmdUtilitiesCancel.Enabled = False

         gridUtilities.Enabled = True

         cboUtilitiesType.Enabled = False
         cboUtilitiesType.text = ""
         cboAuthority_Supplier.Enabled = False
         cboAuthority_Supplier.text = ""
         txtUtilitiesReference.Enabled = False
         txtUtilitiesReference.text = ""
         cboUnitUtilityStatus.Enabled = False
         cboUnitUtilityStatus.text = ""
         txtUnitUtilityStDt.Enabled = False
         txtUnitUtilityStDt.text = ""
         txtDateVacated.Enabled = False
         txtDateVacated.text = ""
         txtChargeRate.Enabled = False
         txtChargeRate.text = ""
         txtUnitUtilityIniReading.Enabled = False
         txtUnitUtilityIniReading.text = ""
         txtFinalReading.Enabled = False
         txtFinalReading.text = ""
         txtUnitUtilityCom.Enabled = False
         txtUnitUtilityCom.text = ""

      Case ComponentMode.GridRowOnSelection
         cmdUtilitiesNew.Enabled = True
         cmdUtilitiesEdit.Enabled = True
         cmdUtilitiesSave.Enabled = False
         cmdUtilitiesCancel.Enabled = False

         gridUtilities.Enabled = True

      Case ComponentMode.NewEntryMode
         cmdUtilitiesNew.Enabled = False
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = True
         cmdUtilitiesCancel.Enabled = True

         gridUtilities.Enabled = False

         cboUtilitiesType.Enabled = True
         cboUtilitiesType.text = ""
         cboAuthority_Supplier.Enabled = True
         cboAuthority_Supplier.text = ""
         txtUtilitiesReference.Enabled = True
         txtUtilitiesReference.text = ""
         cboUnitUtilityStatus.Enabled = True
         cboUnitUtilityStatus.text = ""
         txtUnitUtilityStDt.Enabled = True
         txtUnitUtilityStDt.text = ""
         txtDateVacated.Enabled = True
         txtDateVacated.text = ""
         txtChargeRate.Enabled = True
         txtChargeRate.text = ""
         txtUnitUtilityIniReading.Enabled = True
         txtUnitUtilityIniReading.text = ""
         txtFinalReading.Enabled = True
         txtFinalReading.text = ""
         txtUnitUtilityCom.Enabled = True
         txtUnitUtilityCom.text = ""

      Case ComponentMode.EditMode
         cmdUtilitiesNew.Enabled = False
         cmdUtilitiesEdit.Enabled = False
         cmdUtilitiesSave.Enabled = True
         cmdUtilitiesCancel.Enabled = True

         gridUtilities.Enabled = False

         cboUtilitiesType.Enabled = True
         cboAuthority_Supplier.Enabled = True
         txtUtilitiesReference.Enabled = True
         cboUnitUtilityStatus.Enabled = True
         txtUnitUtilityStDt.Enabled = True
         txtDateVacated.Enabled = True
         txtChargeRate.Enabled = True
         txtUnitUtilityIniReading.Enabled = True
         txtFinalReading.Enabled = True
         txtUnitUtilityCom.Enabled = True
    End Select
End Sub
'
'Public Sub PropertyMemoButtonMode(ByVal mode As ComponentMode)
'    Dim ctrl As Control
'
'    Select Case mode
'        Case ComponentMode.DefaultMode
'            cmdPropertyMemoEdit.Enabled = True
'            cmdPropertyMemoSave.Enabled = False
'            cmdPropertyMemoCancel.Enabled = False
'
'            txtPropertyMemo.Enabled = False
'
'        Case ComponentMode.GridRowOnSelection
'            cmdPropertyMemoEdit.Enabled = True
'            cmdPropertyMemoSave.Enabled = False
'            cmdPropertyMemoCancel.Enabled = False
'
'            txtPropertyMemo.Enabled = False
'
'        Case ComponentMode.NewEntryMode
'            cmdPropertyMemoEdit.Enabled = False
'            cmdPropertyMemoSave.Enabled = True
'            cmdPropertyMemoCancel.Enabled = True
'
'            txtPropertyMemo.Enabled = True
'
'        Case ComponentMode.EditMode
'            cmdPropertyMemoEdit.Enabled = False
'            cmdPropertyMemoSave.Enabled = True
'            cmdPropertyMemoCancel.Enabled = True
'
'            txtPropertyMemo.Enabled = True
'    End Select
'End Sub

Public Sub ConfigGridMaintenanceHistory(ByVal rstMHistory_ As ADODB.Recordset)
   Dim iColumn As Integer
   Dim oColumn As ADODB.Field

'  Configure the grid
   gridMaintenanceHistory.Clear
   gridMaintenanceHistory.Rows = 2
   gridMaintenanceHistory.Cols = rstMHistory_.Fields.Count + 1

   For iColumn = 1 To 10
      gridMaintenanceHistory.ColWidth(iColumn - 1) = Label61(iColumn).Left - Label61(iColumn - 1).Left
   Next iColumn
   gridMaintenanceHistory.ColWidth(iColumn) = gridMaintenanceHistory.Width + gridMaintenanceHistory.Left - Label61(iColumn - 1).Left - 70
   gridMaintenanceHistory.ColWidth(6) = 0
   gridMaintenanceHistory.ColWidth(11) = 900
   For iColumn = 12 To rstMHistory_.Fields.Count
      gridMaintenanceHistory.ColWidth(iColumn) = 0
   Next iColumn

   iColumn = 0
   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.RowHeight(0) = 0
   For Each oColumn In rstMHistory_.Fields
      gridMaintenanceHistory.TextMatrix(0, iColumn) = oColumn.Name
      gridMaintenanceHistory.col = iColumn
      gridMaintenanceHistory.CellFontBold = True
      iColumn = iColumn + 1
   Next oColumn
End Sub

Public Sub SetGridPropertyAnalysisHeader(ByVal conPropertyAnalysis As ADODB.Connection)
   'resolved by BOSL
   'issue  494
   'Modified by anol 08 Dec 2014
   gridPropertyAnalysis.Clear
   gridPropertyAnalysis.Rows = 1
   gridPropertyAnalysis.Cols = 9
   gridPropertyAnalysis.ColWidth(0) = 0
   gridPropertyAnalysis.ColWidth(1) = 2080
   gridPropertyAnalysis.ColWidth(2) = 0
   gridPropertyAnalysis.ColWidth(3) = 1525
   gridPropertyAnalysis.ColWidth(4) = 1305
   gridPropertyAnalysis.ColWidth(5) = 1260
   gridPropertyAnalysis.ColWidth(6) = 1320
   gridPropertyAnalysis.ColWidth(7) = 1260
   gridPropertyAnalysis.ColWidth(8) = 3125
End Sub
'
'Public Sub SetGridHealthSafety(ByVal conSafety As ADODB.Connection)
'   Dim rstSafety As New ADODB.Recordset
'   Dim sSQLQuery_ As String
'
'   'On Error Resume Next
'   'Set the ADO Connections to the dataset
'
'    sSQLQuery_ = "SELECT PropertySafetyID, SafetyType, SafetyInspection, " & _
'    "NextDueDate, SafetyStatus, InspectedBy, SafetyTelephone, Certificate " & _
'    "FROM PropertySafety " & _
'    "WHERE PropertyID = '" & txtPropertyID.text & "' "
'
'   rstSafety.Open sSQLQuery_, conSafety, adOpenStatic, adLockReadOnly
'
'   Dim iRow As Integer, iColumn As Integer
'   iRow = 1
'
'   gridSafety.Clear
'   gridSafety.Rows = 2
'   gridSafety.Cols = 8
'
'   gridSafety.ColWidth(0) = 0
'   For iColumn = 2 To gridSafety.Cols - 1
'      gridSafety.ColWidth(iColumn - 1) = Label41(iColumn - 1).Left - Label41(iColumn - 2).Left
'   Next iColumn
'   gridSafety.ColWidth(iColumn - 1) = gridSafety.Width + gridSafety.Left - Label41(iColumn - 2).Left - 70
'
'   Dim oColumn As ADODB.Field
'
'   iColumn = 0
'
'   gridSafety.Cols = rstSafety.Fields.Count
'   gridSafety.Row = 0
'   gridSafety.RowHeight(0) = 0
'
'   For Each oColumn In rstSafety.Fields
'        gridSafety.TextMatrix(0, iColumn) = oColumn.Name
'        gridSafety.Col = iColumn
'        gridSafety.CellFontBold = True
'        iColumn = iColumn + 1
'   Next oColumn
'
'   'SetMaintenanceHistoryControl
'
'   rstSafety.Close
'   Set rstSafety = Nothing
'End Sub

Public Sub SetGridInsurance(ByVal conInsurance As ADODB.Connection)
'   Dim rstInsurance As New ADODB.Recordset
'   Dim sSQLQuery_ As String
'
'   'On Error Resume Next
'   'Set the ADO Connections to the dataset
'
'     sSQLQuery_ = "SELECT PropertyInsuranceID, Insurer, InsuranceType, " & _
'      "PolicyNo, SumInsured, AnnualPR, MonthlyPR, ExpiryDate, " & _
'      "RechargeableAmount, RechargeableFreq, InsContact, Telephone " & _
'      "FROM PropertyInsurance " & _
'      "WHERE PropertyID = '" & txtPropertyID.text & "' "
'
'         'txtProAddressLine1.text = strSQLQuery_
'
'   rstInsurance.Open sSQLQuery_, conInsurance, adOpenStatic, adLockReadOnly
'
'   Dim iRow As Integer, iColumn As Integer, oColumn As ADODB.Field
'   iRow = 1
'
'   gridInsurance.Clear
'   gridInsurance.Rows = 2
'   gridInsurance.Cols = 12
'
'   gridInsurance.ColWidth(0) = 0
'
'   For iColumn = 2 To gridInsurance.Cols - 1
'      gridInsurance.ColWidth(iColumn - 1) = Label4(iColumn - 1).Left - Label4(iColumn - 2).Left
'   Next iColumn
'   gridInsurance.ColWidth(iColumn - 1) = gridInsurance.Width + gridInsurance.Left - Label4(iColumn - 2).Left - 70
'
'   iColumn = 0
'
'   gridInsurance.Cols = rstInsurance.Fields.Count
'   gridInsurance.Row = 0
'   gridInsurance.RowHeight(0) = 0
'
'   For Each oColumn In rstInsurance.Fields
'        gridInsurance.TextMatrix(0, iColumn) = oColumn.Name
'        gridInsurance.Col = iColumn
'        gridInsurance.CellFontBold = True
'        iColumn = iColumn + 1
'   Next oColumn
'
'   rstInsurance.Close
'   Set rstInsurance = Nothing
End Sub

Public Sub setGridUtilities(ByVal conUtilities As ADODB.Connection)
   Dim rstUtilities As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   'Set the ADO Connections to the dataset

   'CREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT PropertyUtilitiesID, Authority_Supplier, UtilitiesType, " & _
      "UtilitiesReference, ChargeRate, UtilitiesTelephone, FinalReading, DateVacated " & _
      "FROM PropertyUtilities " & _
      "WHERE PropertyID = '" & txtPropertyID.text & "' "

   rstUtilities.Open sSQLQuery_, conUtilities, adOpenStatic, adLockReadOnly

   Dim iRow As Integer, iColumn As Integer
   iRow = 1

   gridUtilities.Clear
   gridUtilities.Rows = 2
   gridUtilities.Cols = 8

   gridUtilities.ColWidth(0) = 0
   For iColumn = 2 To gridUtilities.Cols - 1
      gridUtilities.ColWidth(iColumn - 1) = Label82(iColumn - 1).Left - Label82(iColumn - 2).Left
   Next iColumn
   gridUtilities.ColWidth(iColumn - 1) = gridUtilities.Width + gridUtilities.Left - Label82(iColumn - 2).Left - 70

   Dim oColumn As ADODB.Field

   iColumn = 0

   gridUtilities.Cols = rstUtilities.Fields.Count
   gridUtilities.row = 0
   gridUtilities.RowHeight(0) = 0

   For Each oColumn In rstUtilities.Fields
      gridUtilities.TextMatrix(0, iColumn) = oColumn.Name
      gridUtilities.col = iColumn
      gridUtilities.CellFontBold = True
      iColumn = iColumn + 1
   Next oColumn

   rstUtilities.Close
   Set rstUtilities = Nothing
End Sub

Public Sub LoadGridMaintenanceHistory(ByVal conMHistory_ As ADODB.Connection)
   Dim rstMHistory_ As New ADODB.Recordset
   Dim szSQL As String
' Comment out by anol 20161121 view job was not working
'   szSQL = "SELECT IIF(H.RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
'                "H.ReportedDate, H.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, " & _
'                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
'                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
'                "H.Detail, H.ActualCost, H.ReportedBy, H.AssignedIL, " & _
'                "H.ReportedIS, H.RemindTime, H.Urgent, H.MaintenanceType, " & _
'                "H.ReportedFrom, H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed " & _
'           "FROM PropertyMaintHistory AS H, SecondaryCode AS S " & _
'           "WHERE H.PropertyID = '" & txtPropertyID.text & "' AND " & _
'               "S.Code = H.MaintenanceType AND " & _
'               "S.PrimaryCode = 'MTYP' " & _
'           "ORDER BY H.ReportedDate DESC;"
'Debug.Print szSQL
        szSQL = "SELECT IIF(RecordType = 'J', 'JOB', 'DIARY') AS T, S.Value, " & _
                "H.ReportedDate, U.PropertyID & '-' & H.ID AS Ref, H.Job_DiaryName, H.TaskOwner, H.ReportedBy, " & _
                "H.AssignedTo, H.RemindDate, IIF(H.Alarm, 'YES', 'NO'), H.DateCompleted, " & _
                "H.BudgetCost, H.ExpectedStartDate, H.ExpectedCompletionDate, " & _
                "H.Detail, H.ActualCost, H.AssignedIL, H.ReportedIS, " & _
                "H.RemindTime, H.Urgent, H.MaintenanceType, H.ReportedFrom, " & _
                "H.FundID, H.OverrideBudget, H.FYrID, H.BudgetPassed, " & _
                "P.PropertyID, P.ClientID, '', P.PropertyName , '', '',(Select C.ClientName from Client C where C.ClientID=P.ClientID) AS ClientName,(Select FundName from fund where FUNDID=H.FundID) as FundName, ( Select FinancialYear from FinancialYear where FYrID=H.FYrID) as FinancialYear " & _
           "FROM PropertyMaintHistory AS H, SecondaryCode AS S , Units AS U, " & _
                "LeaseDetails AS L, Property AS P " & _
           "WHERE H.PropertyID = '" & txtPropertyID.text & "' AND  U.UnitNumber = L.UnitNumber AND U.PropertyID= H.PropertyID AND  H.PropertyID = P.PropertyID AND H.ReportedBy=L.SageAccountNumber AND " & _
               "S.Code = H.MaintenanceType AND " & _
               "S.PrimaryCode = 'MTYP' " & _
           "ORDER BY H.ReportedDate DESC;"
   rstMHistory_.Open szSQL, conMHistory_, adOpenStatic, adLockReadOnly

   ConfigGridMaintenanceHistory rstMHistory_

   If rstMHistory_.EOF Then
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
      Exit Sub
   Else
      rstMHistory_.Close
      Set rstMHistory_ = Nothing
   End If

   populateGridDefinedHeader conMHistory_, szSQL, gridMaintenanceHistory

   gridMaintenanceHistory.row = 0
   gridMaintenanceHistory.col = 0
End Sub

Public Sub PopulateGridPropertyAnalysis(ByVal conPropertyAnalysis As ADODB.Connection)
   Dim rstPropertyAnalysis_ As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
  
   sSQLQuery_ = "SELECT PropertyAnalysis.PropertyAnalysisID, " & _
      "PropertyAnalysis.AnalysisType, " & _
      "PropertyAnalysis.AnalysisDescription, " & _
      "PropertyAnalysis.AnalysisOption, PropertyAnalysis.AnalysisValue, " & _
      "PropertyAnalysis.AnalysisQuantity, PropertyAnalysis.AnalysisPercentage, " & _
      "SecondaryCode.value,PropertyAnalysis.AnalysisValue1 ,PropertyAnalysis.Reference  " & _
      "FROM PropertyAnalysis, SecondaryCode " & _
      "WHERE PropertyAnalysis.PropertyID = '" & txtPropertyID.text & "' " & _
      "AND SecondaryCode.Code = PropertyAnalysis.AnalysisType " & _
      "AND SecondaryCode.PrimaryCode = 'ATYP'"
'   Debug.Print sSQLQuery_

   rstPropertyAnalysis_.Open sSQLQuery_, conPropertyAnalysis, adOpenStatic, adLockReadOnly
    
   Dim iRow As Integer
   iRow = 1
   SetGridPropertyAnalysisHeader conPropertyAnalysis
   gridPropertyAnalysis.Clear
   'Modified by anol 08 Dec 2014
   gridPropertyAnalysis.Rows = 1
   gridPropertyAnalysis.RowHeight(0) = 0
   If rstPropertyAnalysis_.EOF = True Then
       gridPropertyAnalysis.Rows = 2
   End If
   gridPropertyAnalysis.Cols = 9
   While Not rstPropertyAnalysis_.EOF
      gridPropertyAnalysis.AddItem ""
      gridPropertyAnalysis.TextMatrix(iRow, 0) = rstPropertyAnalysis_!PropertyAnalysisID
      gridPropertyAnalysis.TextMatrix(iRow, 1) = rstPropertyAnalysis_!Value
      gridPropertyAnalysis.TextMatrix(iRow, 2) = rstPropertyAnalysis_!AnalysisDescription
      gridPropertyAnalysis.TextMatrix(iRow, 3) = rstPropertyAnalysis_!AnalysisOption
      gridPropertyAnalysis.TextMatrix(iRow, 4) = rstPropertyAnalysis_!AnalysisValue
      gridPropertyAnalysis.TextMatrix(iRow, 5) = rstPropertyAnalysis_!AnalysisQuantity
      gridPropertyAnalysis.TextMatrix(iRow, 6) = rstPropertyAnalysis_!AnalysisPercentage
      'Modified by BOSL
      'Issue no 494 : Property Form - Property analysis Form
      'Modified by anol 02 Nov 2014
      gridPropertyAnalysis.TextMatrix(iRow, 7) = IIf(IsNull(rstPropertyAnalysis_!AnalysisValue1), "", rstPropertyAnalysis_!AnalysisValue1)
      gridPropertyAnalysis.TextMatrix(iRow, 8) = IIf(IsNull(rstPropertyAnalysis_!Reference), "", rstPropertyAnalysis_!Reference)
            
      rstPropertyAnalysis_.MoveNext
      
      iRow = iRow + 1
   Wend

   rstPropertyAnalysis_.Close
   Set rstPropertyAnalysis_ = Nothing
End Sub

Public Sub ConfigureGridSafety()
   Dim iRow As Integer, szHeader As String
   iRow = 1

   szHeader$ = "UnitSafetyID|SafetyType|Schedule|Ref|DateChk|NextDueDate|" & _
               "InspectedBy|<SafetyTelephone|<Comment|Alarm|Certificate|Attach"

   gridSafety.Clear
   gridSafety.FormatString = szHeader
   gridSafety.Rows = 2
   gridSafety.Cols = 12
   gridSafety.RowHeight(0) = 0

   gridSafety.ColWidth(0) = 0
   gridSafety.ColWidth(1) = Label41(1).Left - Label41(0).Left
   gridSafety.ColWidth(2) = Label41(2).Left - Label41(1).Left
   gridSafety.ColWidth(3) = Label41(3).Left - Label41(2).Left
   gridSafety.ColWidth(4) = Label41(4).Left - Label41(3).Left
   gridSafety.ColWidth(5) = Label41(5).Left - Label41(4).Left
   gridSafety.ColWidth(6) = Label41(6).Left - Label41(5).Left
   gridSafety.ColWidth(7) = Label41(7).Left - Label41(6).Left
   gridSafety.ColWidth(8) = VerticalLabel(0).Left - Label41(7).Left
   gridSafety.ColWidth(9) = VerticalLabel(1).Left - VerticalLabel(0).Left
   gridSafety.ColWidth(10) = VerticalLabel(2).Left - VerticalLabel(1).Left
   gridSafety.ColWidth(11) = gridSafety.Width + gridSafety.Left - VerticalLabel(2).Left - 300

   SetControlStyle gridSafety
End Sub

Public Sub PopulateHealthSafety(ByVal conSafety As ADODB.Connection)
   Dim rstSafety As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next
   sSQLQuery_ = "SELECT U.UnitSafetyID, U.SafetyInspection, U.Attachment, " & _
      "U.NextDueDate, U.SafetyStatus, U.SafetyTelephone, " & _
      "U.Certificate, U.LastDate, U.Comment, " & _
      "U.Alarm, S.Value AS SV, SC.Value AS SS, SE.Value AS InspectedBy " & _
   "FROM ((UnitSafety AS U LEFT JOIN SecondaryCode AS S ON U.SafetyType = S.Code) " & _
      "LEFT JOIN SecondaryCode AS SC ON U.SafetyStatus = SC.Code) " & _
      "LEFT JOIN SecondaryCode AS SE ON U.InspectedBy =  SE.Code " & _
   "WHERE U.UNITNUMBER = '" & txtPropertyID.text & "' AND Module = 'P';"
'Debug.Print sSQLQuery_
   rstSafety.Open sSQLQuery_, conSafety, adOpenDynamic, adLockOptimistic

   Dim iRow As Integer
   iRow = 1

   ConfigureGridSafety

   While Not rstSafety.EOF
      gridSafety.TextMatrix(iRow, 0) = rstSafety!UnitSafetyID
      gridSafety.TextMatrix(iRow, 1) = rstSafety!SV
      gridSafety.TextMatrix(iRow, 2) = IIf(IsNull(rstSafety!SS), "", rstSafety!SS)
      gridSafety.TextMatrix(iRow, 3) = IIf(IsNull(rstSafety!SafetyInspection), "", rstSafety!SafetyInspection)
      gridSafety.TextMatrix(iRow, 4) = IIf(IsNull(rstSafety!LastDate), "", Format(rstSafety!LastDate, "dd/mm/yyyy"))
      gridSafety.TextMatrix(iRow, 5) = IIf(IsNull(rstSafety!NextDueDate), "", rstSafety!NextDueDate)
      gridSafety.TextMatrix(iRow, 6) = IIf(IsNull(rstSafety!InspectedBy), "", rstSafety!InspectedBy)
      gridSafety.TextMatrix(iRow, 7) = IIf(IsNull(rstSafety!SafetyTelephone), "", rstSafety!SafetyTelephone)
      gridSafety.TextMatrix(iRow, 8) = IIf(IsNull(rstSafety!comment), "", rstSafety!comment)
      gridSafety.TextMatrix(iRow, 9) = IIf(IIf(IsNull(rstSafety!Alarm), "N", rstSafety!Alarm) = "Y", "Yes", "No")
      gridSafety.TextMatrix(iRow, 10) = IIf(IIf(IsNull(rstSafety!Certificate), "N", rstSafety!Certificate), "Yes", "No")
      gridSafety.TextMatrix(iRow, 11) = IIf(IIf(IsNull(rstSafety!attachment), "N", rstSafety!attachment) = "Y", "Yes", "No")

      rstSafety.MoveNext
      gridSafety.AddItem ""
      iRow = iRow + 1
   Wend

   rstSafety.Close
   Set rstSafety = Nothing
   gridSafety.row = 0
End Sub

Public Sub PopulateInsurance(ByVal conInsurance As ADODB.Connection)
   Dim rstInsurance As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'On Error Resume Next

   sSQLQuery_ = "SELECT I.*, " & _
                     "SC1.value AS Insu, SC2.value AS InsType, " & _
                     "SC3.value as U " & _
               "FROM ((PropertyInsurance AS I LEFT JOIN SecondaryCode AS SC1 ON " & _
                     "(I.Insurer = SC1.Code AND SC1.PrimaryCode = 'IRER')) " & _
                  "LEFT JOIN SecondaryCode AS SC2 ON (I.InsuranceType = SC2.Code AND SC2.PrimaryCode = 'ITYP')) " & _
                  "LEFT JOIN SecondaryCode AS SC3 ON (I.Usage = SC3.Code AND SC3.PrimaryCode = 'UUSE') " & _
               "WHERE I.PropertyID = '" & txtPropertyID.text & "' AND I.Module = 'P';"

'Debug.Print sSQLQuery_
   rstInsurance.Open sSQLQuery_, conInsurance, adOpenStatic, adLockReadOnly

   Dim iRow As Integer
   ConfigureGridInsurance

   iRow = 1

   While Not rstInsurance.EOF
      gridInsurance.TextMatrix(iRow, 0) = rstInsurance!PropertyInsuranceID
      gridInsurance.TextMatrix(iRow, 1) = IIf(IsNull(rstInsurance!Insu), "", rstInsurance!Insu)
      gridInsurance.TextMatrix(iRow, 2) = IIf(IsNull(rstInsurance!InsType), "", rstInsurance!InsType)
      gridInsurance.TextMatrix(iRow, 3) = rstInsurance!PolicyNo
      gridInsurance.TextMatrix(iRow, 4) = IIf(IsNull(rstInsurance!SumInsured), "0.00", Format(rstInsurance!SumInsured, "0.00"))
      gridInsurance.TextMatrix(iRow, 5) = IIf(IsNull(rstInsurance!AnnualPR), "0.00", Format(rstInsurance!AnnualPR, "0.00"))
      gridInsurance.TextMatrix(iRow, 6) = IIf(IsNull(rstInsurance!StartDate), "", Format(rstInsurance!StartDate, "dd/mm/yyyy"))
      gridInsurance.TextMatrix(iRow, 7) = IIf(IsNull(rstInsurance!ExpiryDate), "", rstInsurance!ExpiryDate)
      gridInsurance.TextMatrix(iRow, 8) = IIf(IsNull(rstInsurance!Telephone), "", rstInsurance!Telephone)
      gridInsurance.TextMatrix(iRow, 9) = IIf(IsNull(rstInsurance!u), "", rstInsurance!u)
      gridInsurance.TextMatrix(iRow, 10) = IIf(IsNull(rstInsurance!Comments), "", rstInsurance!Comments)
      gridInsurance.TextMatrix(iRow, 11) = IIf(IIf(IsNull(rstInsurance!attachment), "N", rstInsurance!attachment) = "Y", "Yes", "No")
      rstInsurance.MoveNext
      gridInsurance.AddItem ""
      iRow = iRow + 1
   Wend

   rstInsurance.Close
   Set rstInsurance = Nothing
   gridInsurance.row = 0
   gridInsurance.col = 0
End Sub

Public Sub ConfigureGridInsurance()
   Dim szHeader As String

   szHeader$ = "PropertyInsuranceID|<Insurer|<InsuranceType|<" & _
               "PolicyNo|>SumInsured|>AnnualPR|<StartDate|<ExpiryDate|<" & _
               "Telephone|<Usage|<Comments|<Attachment"
                  
   gridInsurance.Clear
   gridInsurance.FormatString = szHeader
   gridInsurance.Rows = 2
   gridInsurance.Cols = 12
   gridInsurance.RowHeight(0) = 0

   gridInsurance.ColWidth(0) = 0                                        'ID
   gridInsurance.ColWidth(1) = Label6(1).Left - Label6(0).Left          'Insurer
   gridInsurance.ColWidth(2) = Label6(2).Left - Label6(1).Left          'Ins Type
   gridInsurance.ColWidth(3) = Label6(3).Left - Label6(2).Left          'Policy No
   gridInsurance.ColWidth(4) = Label6(4).Left - Label6(3).Left          'Sum Ins
   gridInsurance.ColWidth(5) = Label6(5).Left - Label6(4).Left          'Annual PR
   gridInsurance.ColWidth(6) = Label6(6).Left - Label6(5).Left          'St Date
   gridInsurance.ColWidth(7) = Label6(7).Left - Label6(6).Left          'Exp Date
   gridInsurance.ColWidth(8) = Label6(8).Left - Label6(7).Left          'Tel
   gridInsurance.ColWidth(9) = Label6(9).Left - Label6(8).Left          'Usage
   gridInsurance.ColWidth(10) = Label6(10).Left - Label6(9).Left        'Comment
   gridInsurance.ColWidth(11) = gridInsurance.Width + gridInsurance.Left - (Label6(10).Left + fraInsurance.Left) - 200  'Attach
End Sub

Public Sub PopulateUtilities(ByVal conUtilities As ADODB.Connection)
   Dim rstUtilities As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim iRow As Integer

   'On Error Resume Next
   sSQLQuery_ = _
         "SELECT U.*, " & _
              "SC1.value AS UT, SC2.value AS US, S.SupplierName " & _
         "FROM ((PropertyUtilities AS U LEFT JOIN SecondaryCode AS SC1 ON U.UtilitiesType = SC1.Code) " & _
              "LEFT JOIN SecondaryCode AS SC2 ON U.Status = SC2.Code) INNER JOIN " & _
              "Supplier AS S ON U.Authority_Supplier = S.SupplierID " & _
         "WHERE U.PropertyID = '" & txtPropertyID.text & "' " & _
              "AND SC1.PrimaryCode = 'UTIL' " & _
              "AND SC2.PrimaryCode = 'USTA';"
'Debug.Print sSQLQuery_
   rstUtilities.Open sSQLQuery_, conUtilities, adOpenStatic, adLockReadOnly

   iRow = 1

   ConfigureGridUtilities

   While Not rstUtilities.EOF
      gridUtilities.TextMatrix(iRow, 0) = IIf(IsNull(rstUtilities!PropertyUtilitiesID), "", rstUtilities!PropertyUtilitiesID)
      gridUtilities.TextMatrix(iRow, 1) = IIf(IsNull(rstUtilities!Occupier), "", rstUtilities!Occupier)
      gridUtilities.TextMatrix(iRow, 2) = IIf(IsNull(rstUtilities!UT), "", rstUtilities!UT)
      gridUtilities.TextMatrix(iRow, 3) = IIf(IsNull(rstUtilities!Authority_Supplier), "", rstUtilities!SupplierName)
      gridUtilities.TextMatrix(iRow, 4) = IIf(IsNull(rstUtilities!UtilitiesReference), "", rstUtilities!UtilitiesReference)
      gridUtilities.TextMatrix(iRow, 5) = IIf(IsNull(rstUtilities!US), "", rstUtilities!US)
      gridUtilities.TextMatrix(iRow, 6) = IIf(IsNull(rstUtilities!StartDate), "", Format(rstUtilities!StartDate, "dd/mm/yyyy"))
      gridUtilities.TextMatrix(iRow, 7) = IIf(IsNull(rstUtilities!DateVacated), "", Format(rstUtilities!DateVacated, "dd/mm/yyyy"))
      gridUtilities.TextMatrix(iRow, 8) = IIf(IsNull(rstUtilities!ChargeRate), "0.00", Format(rstUtilities!ChargeRate, "0.00"))
      gridUtilities.TextMatrix(iRow, 9) = IIf(IsNull(rstUtilities!InitialReading), "", rstUtilities!InitialReading)
      gridUtilities.TextMatrix(iRow, 10) = IIf(IsNull(rstUtilities!FinalReading), "", rstUtilities!FinalReading)
      gridUtilities.TextMatrix(iRow, 11) = IIf(IsNull(rstUtilities!Comments), "", rstUtilities!Comments)

      rstUtilities.MoveNext
      gridUtilities.AddItem ""
      iRow = iRow + 1
   Wend

   rstUtilities.Close
   Set rstUtilities = Nothing

   gridUtilities.row = 0
End Sub

Private Sub ConfigureGridUtilities()
   Dim iRow As Integer, szHeader As String
   iRow = 1

   szHeader$ = "UnitUtilitiesID|Occupier|UtilitiesType|Authority_Supplier|UtilitiesReference|UnitUtilityStatus" & _
               "|UnitUtilityStDt|DateVacated|ChargeRate|UnitUtilityIniReading|FinalReading|UnitUtilityCom"

   gridUtilities.Clear
   gridUtilities.FormatString = szHeader
   gridUtilities.Rows = 2
   gridUtilities.Cols = 12
   gridUtilities.RowHeight(0) = 0

   gridUtilities.ColWidth(0) = 0
   gridUtilities.ColWidth(1) = Label82(1).Left - Label82(0).Left
   gridUtilities.ColWidth(2) = Label82(2).Left - Label82(1).Left
   gridUtilities.ColWidth(3) = Label82(3).Left - Label82(2).Left
   gridUtilities.ColWidth(4) = Label82(4).Left - Label82(3).Left
   gridUtilities.ColWidth(5) = Label82(5).Left - Label82(4).Left
   gridUtilities.ColWidth(6) = Label82(6).Left - Label82(5).Left
   gridUtilities.ColWidth(7) = Label82(7).Left - Label82(6).Left
   gridUtilities.ColWidth(8) = Label82(8).Left - Label82(7).Left
   gridUtilities.ColWidth(9) = Label82(9).Left - Label82(8).Left
   gridUtilities.ColWidth(10) = Label82(10).Left - Label82(9).Left
   gridUtilities.ColWidth(11) = gridUtilities.Width + gridUtilities.Left - Label82(10).Left - 300
End Sub

'
'Public Function SavePropertyMaintenanceHistory(ByVal conMHistory_ as adodb.connection) As Boolean
'    Dim rstMHistory_ As New ADODB.Recordset
'    Dim rstDEL_MHistory As New ADODB.Recordset
'    Dim rstID As New ADODB.Recordset
'    Dim sSQLQuery_ As String
'    Dim sSQLDelete As String
'    Dim sSQLFilter As String
'    Dim iRowIndex As Integer
'    Dim lTableID As Long
'
'    sSQLFilter = ""
'
'    On Error GoTo Exception
'
'    If Not M_HISTORY_NEW_ENTRY_ Then
'        sSQLFilter = "WHERE PropertyID = '" & txtPropertyID.text & "' AND ID = " & txtID.text & ""
'    Else
'        sSQLFilter = ""
'    End If
'
'    sSQLQuery_ = "SELECT * " & _
'    "FROM PropertyMAINTHISTORY " & sSQLFilter
'
'    Set rstMHistory_ = conMHistory_.OpenResultset(sSQLQuery_, rdOpenDynamic, rdConcurRowVer)
'
'    If M_HISTORY_NEW_ENTRY_ Then
'      rstMHistory_.AddNew
'      sSQLQuery_ = "SELECT MAX(ID) AS M_ID FROM PropertyMAINTHISTORY;"
'      Set rstID = conMHistory_.OpenResultset(sSQLQuery_, rdOpenDynamic, rdConcurRowVer)
'      lTableID = IIf(IsNull(rstID!M_ID), 0, rstID!M_ID) + 1
'      rstID.Close
'      Set rstID = Nothing
'   Else
'      rstMHistory_.Edit
'      lTableID = CLng(txtID.text)
'   End If
'
'    rstMHistory_!PROPERTYID = txtPropertyID.text
'    rstMHistory_!MaintenanceType = cboMaintenanceType.Value
'    rstMHistory_!ReportedDate = Format(IIf(dtpReportedDate.text = "", Now, dtpReportedDate.text), "DD MMMM YYYY")
'    rstMHistory_!description = IIf(txtDescription.text = "", "", txtDescription.text)
'    rstMHistory_!EstimateCost = IIf(txtEstimateCost.text = "", "", Format(txtEstimateCost.text, "0.00"))
'    rstMHistory_!TaskOwner = IIf(cboTaskOwner.text = "", "", cboTaskOwner.text)
'    rstMHistory_!Contact = IIf(cboContact.text = "", "", cboContact.text)
'    rstMHistory_!RemindDate = Format(IIf(dtpRemindDate.text = "", Now, dtpRemindDate.text), "DD MMMM YYYY")
'    rstMHistory_!DateCompleted = Format(IIf(dtpDateCompleted.text = "", Now, dtpDateCompleted.text), "DD MMMM YYYY")
'
'    If chkAlarm.Value = 1 Then
'       rstMHistory_!Alarm = True
'
'       If M_HISTORY_NEW_ENTRY_ Then
'          rstMHistory_!REMINDER_ID = NewReminder(Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), "083000", txtDescription.text, "PropertyMAINTHISTORY", CStr(lTableID))
'       Else
'          UpdateReminder rstMHistory_!REMINDER_ID, Format(CDate(rstMHistory_!RemindDate), "YYYYMMDD"), "083000", txtDescription.text
'       End If
'    Else
'       rstMHistory_!Alarm = False
'
'       ClearReminder rstMHistory_!REMINDER_ID
'    End If
'
'    rstMHistory_.Update
'
'    rstMHistory_.Close
'    Set rstMHistory_ = Nothing
'
'    SavePropertyMaintenanceHistory = True
'    Exit Function
'
'Exception:
'    MsgBox ERR.Number & " - " & ERR.description & " - " & ERR.HelpContext & " - " & ERR.Source, vbOKOnly, "Error"
'    rstMHistory_.Close
'
'    Set rstMHistory_ = Nothing
'
'    SavePropertyMaintenanceHistory = False
'End Function

Public Function SavePropertyAnalysis(ByVal conPropertyAnalysis As ADODB.Connection) As Boolean
      Dim rstPropertyAnalysis_ As New ADODB.Recordset
   Dim rstProperty_ As New ADODB.Recordset

   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   If Not Property_ANALYSIS_NEW_ENTRY Then
       sSQLFilter = "WHERE PropertyID = '" & txtPropertyID.text & "' AND PropertyAnalysisID = " & txtPropertyAnalysisID.text & ""
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyANALYSIS " & sSQLFilter

   rstPropertyAnalysis_.Open sSQLQuery_, conPropertyAnalysis, adOpenDynamic, adLockOptimistic

   If Property_ANALYSIS_NEW_ENTRY Then rstPropertyAnalysis_.AddNew

   rstPropertyAnalysis_!propertyID = txtPropertyID.text
   rstPropertyAnalysis_!AnalysisType = IIf(cboAnalysisType.Value <> "", cboAnalysisType.Value, "0")
   rstPropertyAnalysis_!AnalysisDescription = IIf(txtAnalysisDescription.text <> "", txtAnalysisDescription.text, "")
   rstPropertyAnalysis_!AnalysisOption = IIf(cboAnalysisOption.Value <> "", cboAnalysisOption.Value, "")
   rstPropertyAnalysis_!AnalysisValue = IIf(txtAnalysisValue.text <> "", txtAnalysisValue.text, "0")
   rstPropertyAnalysis_!AnalysisQuantity = IIf(txtAnalysisQuantity.text <> "", txtAnalysisQuantity.text, "0")
   rstPropertyAnalysis_!AnalysisPercentage = IIf(txtAnalysisPercentage.text <> "", txtAnalysisPercentage.text, "0")
   'Resolved by BOSL
   'issue 494: Property Form - Property analysis Form
   'Modified by Anol 02 Nov 2014
   rstPropertyAnalysis_!AnalysisValue1 = IIf(txtAnalysisValue1.text <> "", txtAnalysisValue1.text, "0")
   rstPropertyAnalysis_!Reference = IIf(txtAnalysisReference.text <> "", txtAnalysisReference.text, "")
   'end of modification

   rstPropertyAnalysis_.Update

   rstPropertyAnalysis_.Close

   Set rstProperty_ = Nothing
   Set rstPropertyAnalysis_ = Nothing

   SavePropertyAnalysis = True
End Function

Public Function SaveHealthSafety(ByVal conSafety As ADODB.Connection) As Boolean
   Dim rstSafety As New ADODB.Recordset
   Dim sSQLQuery_ As String, sSQLDelete As String
   Dim sSQLFilter As String, iRowIndex As Integer

   sSQLFilter = ""

   'On Error GoTo Exception

   If Not HEALTH_SAFETY_NEW_ENTRY Then
       sSQLFilter = "WHERE UNITNUMBER = '" & txtPropertyID.text & "' AND UnitSafetyID = '" & txtUnitSafetyID.text & "' AND Module = 'P'"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT U.UnitSafetyID, U.Attachment, " & _
                  "U.SafetyInspection, U.UnitNumber, " & _
                  "U.NextDueDate, U.SafetyStatus, " & _
                  "U.InspectedBy, U.SafetyTelephone, " & _
                  "U.Certificate, U.LastDate, U.Comment, " & _
                  "U.Alarm, U.SafetyType, U.Module, U.spare1 " & _
                "FROM UnitSafety AS U " & sSQLFilter

   rstSafety.Open sSQLQuery_, conSafety, adOpenDynamic, adLockOptimistic

   If HEALTH_SAFETY_NEW_ENTRY Then
      rstSafety.AddNew

      rstSafety!UnitSafetyID = UniqueID()

      If chkAlarm.Value = 1 Then
         rstSafety!spare1 = NewReminder(Format(txtNextDueDate.text, "YYYYMMDD"), "010000", _
                           "Health and safety issue for the property no. " & txtPropertyID.text, _
                           "UnitSafety", rstSafety!UnitSafetyID)
         rstSafety!Alarm = "Y"
      End If
   Else
'      rstSafety!UnitSafetyID = HEALTH_SAFETY_ID
      If rstSafety!Alarm = "Y" And chkAlarm.Value = 0 Then _
         ClearReminder rstSafety!spare1
      If rstSafety!Alarm = "Y" And chkAlarm.Value = 1 Then _
         UpdateReminder rstSafety!spare1, Format(txtNextDueDate.text, "YYYYMMDD"), "010000", _
                        "Health and safety issue for the property no. " & txtPropertyID.text
   End If
   rstSafety!UnitNumber = txtPropertyID.text
   rstSafety!SafetyType = cboSafetyType.BoundText
   rstSafety!SafetyInspection = txtRef.text
   rstSafety!NextDueDate = IIf(txtNextDueDate.text = "", Null, Format(txtNextDueDate.text, "dd mmmm yyyy"))
   rstSafety!SafetyStatus = cboSchedule.BoundText
   rstSafety!InspectedBy = cboInspectedBy.BoundText
   rstSafety!SafetyTelephone = txtSafetyTelephone.text
   rstSafety!Certificate = IIf(chkCertificate.Value = 1, True, False)
   rstSafety!LastDate = IIf(txtDateChk.text = "", Format(txtDateChk, "dd mmmm yyyy"), Format(txtDateChk.text, "dd mmmm yyyy"))
   rstSafety!comment = txtComment.text
   rstSafety!Alarm = IIf(chkAlarm.Value = 1, "Y", "N")
   rstSafety!attachment = IIf(HEALTH_N_SAFETY_ATTACH, "Y", "N")
   rstSafety!Module = "P"
   rstSafety.Update

   rstSafety.Close
   Set rstSafety = Nothing

   SaveHealthSafety = True
   Exit Function

Exception:
   'MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
   rstSafety.Close
   conSafety.Close
   Set rstSafety = Nothing
   Set conSafety = Nothing
   SaveHealthSafety = False
End Function

Public Function SavePropertyUtilities(ByVal conUtilities As ADODB.Connection) As Boolean
   If cboUtilitiesType.text = "" Then
      ShowMsgInTaskBar "Please select utility type to save."
      Exit Function
   End If
   If cboAuthority_Supplier.text = "" Then
      ShowMsgInTaskBar "Please select supplier to save."
      Exit Function
   End If
   If cboUnitUtilityStatus.text = "" Then
      ShowMsgInTaskBar "Please select the status of the utility to save."
      Exit Function
   End If

   Dim rstUtilities As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   On Error GoTo Exception

   If Not UNIT_UTILITIES_NEW_ENTRY Then
       sSQLFilter = "WHERE PropertyID = '" & txtPropertyID.text & "' AND " & _
                          "PropertyUtilitiesID = " & txtUnitUtilitiesID.text & " AND " & _
                          "Module = 'P';"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyUtilities " & sSQLFilter

   rstUtilities.Open sSQLQuery_, conUtilities, adOpenDynamic, adLockOptimistic

   If UNIT_UTILITIES_NEW_ENTRY Then rstUtilities.AddNew

   rstUtilities!propertyID = txtPropertyID.text
   rstUtilities!UtilitiesType = cboUtilitiesType.BoundText
   rstUtilities!Authority_Supplier = cboAuthority_Supplier.BoundText
   rstUtilities!UtilitiesReference = txtUtilitiesReference.text
   rstUtilities!Status = cboUnitUtilityStatus.BoundText
   rstUtilities!StartDate = IIf(txtUnitUtilityStDt.text = "", Null, Format(txtUnitUtilityStDt.text, "dd mmmm yyyy"))
   rstUtilities!DateVacated = IIf(txtDateVacated.text = "", Null, Format(txtDateVacated.text, "dd mmmm yyyy"))
   rstUtilities!ChargeRate = IIf(txtChargeRate.text = "", "0", txtChargeRate.text)
   rstUtilities!InitialReading = txtUnitUtilityIniReading.text
   rstUtilities!FinalReading = txtFinalReading.text
   rstUtilities!Comments = txtUnitUtilityCom.text
   rstUtilities!Module = "P"
   rstUtilities.Update

   'Next iRowIndex
   rstUtilities.Close

   Set rstUtilities = Nothing

   SavePropertyUtilities = True
   Exit Function

Exception:

   'MsgBox ERR.Number & " - " & ERR.description, vbOKOnly, "Error"
   rstUtilities.Close

   Set rstUtilities = Nothing

   SavePropertyUtilities = False
End Function

Public Function SavePropertyMemo() As Boolean
    Dim conPropertyMemo_ As New ADODB.Connection
    Dim rstPropertyMemo_ As New ADODB.Recordset
    Dim sSQLQuery_ As String, sSQLFilter As String

    On Error GoTo Exception

    'Set the RDO Connections to the dataset
    conPropertyMemo_.Open getConnectionString

    sSQLFilter = "WHERE PropertyID = '" & txtPropertyID.text & "'"

    sSQLQuery_ = "SELECT * " & _
    "FROM Property " & sSQLFilter

    rstPropertyMemo_.Open sSQLQuery_, conPropertyMemo_, adOpenDynamic, adLockOptimistic

'    rstPropertyMemo_.ed
    If txtMemo.text = "" Then
        rstPropertyMemo_!MemoText = "<No memo saved>"
    Else
        rstPropertyMemo_!MemoText = txtMemo.text
    End If
    rstPropertyMemo_.Update

    rstPropertyMemo_.Close
    conPropertyMemo_.Close
    Set rstPropertyMemo_ = Nothing
    Set conPropertyMemo_ = Nothing
    SavePropertyMemo = True
    Exit Function

Exception:

    ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
    rstPropertyMemo_.Close
    conPropertyMemo_.Close
    Set rstPropertyMemo_ = Nothing
    Set conPropertyMemo_ = Nothing
    SavePropertyMemo = False
End Function


Public Function SavePropertyInsurance(ByVal conInsurance As ADODB.Connection) As Boolean
   If cboInsurer.text = "" Then
       ShowMsgInTaskBar "Please enter the name of the Insurer."
       SavePropertyInsurance = False
       Exit Function
   End If

   If txtPolicyNo.text = "" Then
       ShowMsgInTaskBar "Please enter the Policy No."
       SavePropertyInsurance = False
       Exit Function
   End If

   Dim rstInsurance As New ADODB.Recordset
   Dim rstDEL_MHistory As New ADODB.Recordset
   Dim sSQLQuery_ As String
   Dim sSQLDelete As String
   Dim sSQLFilter As String
   Dim iRowIndex As Integer

   sSQLFilter = ""

   On Error GoTo Exception
   'Set the RDO Connections to the dataset

   If Not UNIT_INSURANCE_NEW_ENTRY Then
       sSQLFilter = "WHERE PropertyID = '" & txtPropertyID.text & "' AND " & _
                        "PropertyInsuranceID = '" & txtPropertyInsuranceID.text & "' AND " & _
                        "Module = 'P';"
   Else
       sSQLFilter = ""
   End If

   sSQLQuery_ = "SELECT * " & _
                "FROM PropertyInsurance " & sSQLFilter
'Debug.Print sSQLQuery_
   rstInsurance.Open sSQLQuery_, conInsurance, adOpenDynamic, adLockOptimistic

   'For iRowIndex = 1 To gridUnitAnalysis.Rows - 2
   If UNIT_INSURANCE_NEW_ENTRY Then rstInsurance.AddNew
   If INSURANCE_ID = "" Then
      rstInsurance!PropertyInsuranceID = UniqueID()
   Else
      rstInsurance!PropertyInsuranceID = INSURANCE_ID
   End If

   rstInsurance!propertyID = txtPropertyID.text
   rstInsurance!Insurer = cboInsurer.BoundText
   rstInsurance!InsuranceType = cboInsuranceType.BoundText
   rstInsurance!PolicyNo = txtPolicyNo.text
   rstInsurance!SumInsured = IIf(txtSumInsured.text = "", "0", txtSumInsured.text)
   rstInsurance!AnnualPR = IIf(txtAnnualPR.text = "", "0", txtAnnualPR.text)
   rstInsurance!StartDate = IIf(txtStartDate.text = "", Null, Format(txtStartDate.text, "dd mmmm yyyy"))
   rstInsurance!ExpiryDate = IIf(txtExpiryDate.text = "", Null, Format(txtExpiryDate.text, "dd mmmm yyyy"))
   rstInsurance!Telephone = txtTelephone.text
   rstInsurance!Usage = cboUsage.BoundText
   rstInsurance!Comments = txtComments.text
   rstInsurance!Module = "P"
   rstInsurance!attachment = IIf(HEALTH_N_SAFETY_ATTACH, "Y", "N")

   rstInsurance.Update
   rstInsurance.Close
   Set rstInsurance = Nothing

   SavePropertyInsurance = True
   Exit Function

Exception:

   ShowMsgInTaskBar Err.Number & " - " & Err.description, , "N"
   rstInsurance.Close

   Set rstInsurance = Nothing

   SavePropertyInsurance = False
End Function

Public Function SetTotalArea()
   Dim iCount, iQuantity As Integer
   Dim dTotalArea, dValue As Double

   Dim conProperty_ As New ADODB.Connection
   Dim rstProperty_ As New ADODB.Recordset

   dTotalArea = 0
   For iCount = 1 To gridPropertyAnalysis.Rows - 2
       dValue = CDbl(gridPropertyAnalysis.TextMatrix(iCount, 4))
       iQuantity = CInt(gridPropertyAnalysis.TextMatrix(iCount, 5))
       dTotalArea = dTotalArea + (dValue * iQuantity)
   Next iCount

   txtAnalysisTotalArea.text = dTotalArea

   Dim sSQLQuery_ As String

   sSQLQuery_ = "SELECT ClientID, PropertyID, PropertyName, ProAddressLine1, ProAddressLine2, " & _
                 "ProAddressLine3, ProAddressLine4, ProPostCode, TotalArea " & _
                "FROM Property WHERE PropertyID = '" & txtPropertyID.text & "'"

   conProperty_.Open getConnectionString

   rstProperty_.Open sSQLQuery_, conProperty_, adOpenDynamic, adLockOptimistic

   rstProperty_!TotalArea = IIf(txtAnalysisTotalArea.text <> "", CDbl(txtAnalysisTotalArea.text), 0)
   rstProperty_.Update

   rstProperty_.Close
   conProperty_.Close
   Set rstProperty_ = Nothing
   Set conProperty_ = Nothing
End Function

Public Function CalculateTotalInsurance(ByVal conProperty_ As ADODB.Connection)
'    Dim iCount, iQuantity As Integer
'    Dim dTotalArea, dValue As Double
'    Dim rstProperty_ As New ADODB.Recordset
'
'    Dim sSQLQuery_ As String
'
'    sSQLQuery_ = "SELECT TOTALSUMINSURED.CONTOTAL, TOTALSUMINSURED.PROTOTAL, SUM_RECHARG.TotalRECHARG FROM " _
'    & "(" _
'    & "SELECT SUM_INSTYPE.PROPERTYID, " _
'    & "SUM(IIF(SUM_INSTYPE.TYPE = 'CON', (TOTAL), '0')) AS CONTOTAL, " _
'    & "SUM(IIF(SUM_INSTYPE.TYPE = 'PRO', (TOTAL), '0')) AS PROTOTAL " _
'    & "From " _
'    & " ( " _
'    & " SELECT PropertyInsurance.PropertyId,PropertyInsurance.InsuranceType AS TYPE, SUM(PropertyInsurance.SumInsured) as Total " _
'    & " From PropertyInsurance " _
'    & " GROUP BY PropertyInsurance.PropertyId,PropertyInsurance.InsuranceType " _
'    & " ) AS SUM_INSTYPE " _
'    & " GROUP BY SUM_INSTYPE.PROPERTYID " _
'    & " ) AS TOTALSUMINSURED, " _
'    & " ( " _
'    & " SELECT  PROPERTYINSURANCE.PROPERTYID, SUM(PropertyInsurance.RechargeableAmount) as TotalRECHARG " _
'    & " From PropertyInsurance " _
'    & " GROUP BY PropertyInsurance.PropertyId " _
'    & " ) AS SUM_RECHARG " _
'    & " Where " _
'    & " TOTALSUMINSURED.PROPERTYID = SUM_RECHARG.PROPERTYID AND " _
'    & " TOTALSUMINSURED.PROPERTYID = '" & txtPropertyID.text & "'"
'
'    'Debug.Print sSQLQuery_
'    rstProperty_.Open sSQLQuery_, conProperty_, adOpenDynamic, adLockOptimistic
'
'    On Error Resume Next
'    txtContentValue.text = IIf(IsNull(rstProperty_!CONTOTAL), "0.00", rstProperty_!CONTOTAL)
'    txtPropertyValue.text = IIf(IsNull(rstProperty_!PROTOTAL), "0.00", rstProperty_!PROTOTAL)
'    txtTotalRechargeable.text = IIf(IsNull(rstProperty_!TotalRECHARG), "0.00", rstProperty_!TotalRECHARG)
'
'    rstProperty_.Close
'    Set rstProperty_ = Nothing
End Function

Public Function ComponentEnableModeProperty(ByVal frmCurrent As Form, ByVal mode As ComponentMode)
   Dim ctrl As Control

   Select Case mode
       Case ComponentMode.DefaultMode
'           For Each ctrl In frmCurrent.Controls
'               Select Case TypeName(ctrl)
'                   Case "TextBox"
'                       ctrl.Enabled = False
'                       ctrl.text = ""
'                   Case "CheckBox"
'                       ctrl.Enabled = False
'                   Case "ComboBox"
'                       ctrl.Enabled = False
'               End Select
'           Next ctrl
            For Each ctrl In frmCurrent.Controls
               Select Case TypeName(ctrl)
                   Case "TextBox"
                       ctrl.Locked = True
                       ctrl.text = ""
                   Case "CheckBox"
                        ctrl.Enabled = False
                   Case "ComboBox"
                        ctrl.Locked = True
               End Select
           Next ctrl
           frmCurrent.Controls("gridPropertyLookup").Visible = False

           frmCurrent.cmdNewProperty.Enabled = True
           frmCurrent.cmdEditProperty.Enabled = True
           frmCurrent.cmdSaveProperty.Enabled = False
           frmCurrent.cmdCancelProperty.Enabled = False
           frmCurrent.cmdCloseProperty.Enabled = True

       Case ComponentMode.SavedMode
'           For Each ctrl In frmCurrent.Controls
'               Select Case TypeName(ctrl)
'                   Case "TextBox"
'                       ctrl.Enabled = False
'                   Case "CheckBox"
'                       ctrl.Enabled = False
'                   Case "ComboBox"
'                       ctrl.Enabled = False
'               End Select
'           Next ctrl
            For Each ctrl In frmCurrent.Controls
               Select Case TypeName(ctrl)
                   Case "TextBox"
                        ctrl.Locked = True
                   Case "CheckBox"
                       ctrl.Enabled = False
                   Case "ComboBox"
                        ctrl.Locked = True
               End Select
           Next ctrl

           frmCurrent.Controls("gridPropertyLookup").Visible = False

           frmCurrent.cmdNewProperty.Enabled = True
           frmCurrent.cmdEditProperty.Enabled = True
           frmCurrent.cmdSaveProperty.Enabled = False
           frmCurrent.cmdCancelProperty.Enabled = False
           frmCurrent.cmdCloseProperty.Enabled = True

       Case ComponentMode.GridRowOnSelection
'           For Each ctrl In frmCurrent.Controls
'               Select Case TypeName(ctrl)
'                   Case "TextBox"
'                       ctrl.Enabled = False
'                       ctrl.text = ""
'                   Case "CheckBox", "DataCombo"
'                       ctrl.Enabled = False
'               End Select
'           Next ctrl
        For Each ctrl In frmCurrent.Controls
               Select Case TypeName(ctrl)
                   Case "TextBox"
                       ctrl.Locked = True
                       ctrl.text = ""
                   Case "CheckBox", "DataCombo"
                       ctrl.Locked = True
               End Select
           Next ctrl
           frmCurrent.Controls("gridPropertyLookup").Visible = False
           frmCurrent.cmdNewProperty.Enabled = True
           frmCurrent.cmdEditProperty.Enabled = True
           frmCurrent.cmdSaveProperty.Enabled = False
           frmCurrent.cmdCancelProperty.Enabled = False
           frmCurrent.cmdCloseProperty.Enabled = True

       Case ComponentMode.NewEntryMode
           For Each ctrl In frmCurrent.Controls
               Select Case TypeName(ctrl)
                   Case "TextBox"
                       If ctrl.Name <> "txtClientList" Then
                        ctrl.Locked = False
                        ctrl.text = ""
                       End If
                   Case "DataCombo"
                       ctrl.Locked = False
               End Select
           Next ctrl
           frmCurrent.Controls("gridPropertyLookup").Visible = False
           frmCurrent.cmdNewProperty.Enabled = False
           frmCurrent.cmdEditProperty.Enabled = False
           frmCurrent.cmdSaveProperty.Enabled = True
           frmCurrent.cmdCancelProperty.Enabled = True
           frmCurrent.cmdCloseProperty.Enabled = False

       Case ComponentMode.EditMode
         For Each ctrl In frmCurrent.Controls
             Select Case TypeName(ctrl)
                 Case "CheckBox", "DataCombo", "ComboBox"
                     ctrl.Enabled = True
                 Case "TextBox"
                     ctrl.Locked = True
                        If ctrl.Name <> "txtClientList" Then
                            ctrl.Locked = False
                       End If
             End Select
         Next ctrl
         frmCurrent.Controls("gridPropertyLookup").Visible = False

         frmCurrent.cmdNewProperty.Enabled = False
         frmCurrent.cmdEditProperty.Enabled = False
         frmCurrent.cmdSaveProperty.Enabled = True
         frmCurrent.cmdCancelProperty.Enabled = True
         frmCurrent.cmdCloseProperty.Enabled = False

       Case ComponentMode.GridLostFocus
         For Each ctrl In frmCurrent.Controls
            Select Case TypeName(ctrl)
               Case "TextBox"
                  ctrl.Locked = True
               Case "CheckBox", "DataCombo"
                  ctrl.Enabled = False
            End Select
         Next ctrl
         frmCurrent.Controls("gridPropertyLookup").Visible = False

         frmCurrent.cmdNewProperty.Enabled = True
         frmCurrent.cmdEditProperty.Enabled = False
         frmCurrent.cmdSaveProperty.Enabled = False
         frmCurrent.cmdCancelProperty.Enabled = False
         frmCurrent.cmdCloseProperty.Enabled = True
         
   End Select
End Function

Private Function CheckPropertyID() As Boolean
   Dim conPropertyID As New ADODB.Connection
   Dim rstPropertyID As New ADODB.Recordset
   Dim sSQLQuery_ As String

   'Set the ado Connections to the dataset
   conPropertyID.Open getConnectionString

   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT PROPERTY.PROPERTYID " & _
                "From Property " & _
                "WHERE PROPERTY.PropertyID = '" & txtPropertyID.text & "';"
'Debug.Print sSQLQuery_
   rstPropertyID.Open sSQLQuery_, conPropertyID, adOpenStatic, adLockReadOnly

   If rstPropertyID.EOF Then
      CheckPropertyID = False
   Else
      CheckPropertyID = True
   End If

   rstPropertyID.Close
   conPropertyID.Close
   Set rstPropertyID = Nothing
   Set conPropertyID = Nothing
End Function

Public Function GeneratePropertyID() As String
   Dim conPropertyID As New ADODB.Connection
   Dim rstPropertyID As New ADODB.Recordset
   Dim sSQLQuery_ As String, sPropertyName As String
   Dim MAX_Property_ As String
   Dim PROPERTY_ID_ As String

   'Set the ado Connections to the dataset
   conPropertyID.Open getConnectionString

   sPropertyName = Replace(txtPropertyName.text, "-", "")
   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
   sSQLQuery_ = "SELECT MAX(RIGHT(PROPERTY.PROPERTYID,2)) + 1 AS  MAX_PROPERTYID " & _
            "From Property " & _
            "WHERE LEFT(PROPERTY.PropertyID,2) = LEFT(TRIM('" & sPropertyName & "'),2)"
'Debug.Print sSQLQuery_
   rstPropertyID.Open sSQLQuery_, conPropertyID, adOpenStatic, adLockReadOnly

   If rstPropertyID.EOF Or rstPropertyID.BOF Then
      MAX_Property_ = "1"
   End If

   While Not rstPropertyID.EOF
      MAX_Property_ = IIf(IsNull(rstPropertyID!MAX_PROPERTYID), "1", rstPropertyID!MAX_PROPERTYID)
      rstPropertyID.MoveNext
   Wend

   GeneratePropertyID = UCase(Left(sPropertyName, 2)) & Lpad(MAX_Property_, "0", 2)
   
   rstPropertyID.Close
   conPropertyID.Close
   Set rstPropertyID = Nothing
   Set conPropertyID = Nothing
End Function

Private Sub txtSearchProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If Not SEARCHPropertyMODE_ Then
'        Exit Sub
'    End If
        
    If KeyAscii = 13 Then
        txtPropertySearch.SetFocus
    End If
End Sub
'
'Private Sub txtSumInsured_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   'Added By Samrat. 12/10/2006
'   Dim KA As Integer
'   KA = KeyAscii
'   DigitTextKeyPress txtSumInsured, KA
'   KeyAscii = KA
'End Sub

Private Sub txtStartDate_Change()
   TextBoxChangeDate txtStartDate
End Sub

Private Sub txtStartDate_GotFocus()
   SelTxtInCtrl txtStartDate
End Sub

Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtStartDate, KeyAscii
End Sub

Private Sub txtStartDate_LostFocus()
   TextBoxFormatDate txtStartDate
End Sub
'
'Private Sub txtSuppCaption1_LostFocus()
'   txtSuppCaption1.Visible = False
'   lblSupplementary1.Caption = IIf(txtSuppCaption1.text = "", lblSupplementary1.Caption, txtSuppCaption1.text)
'End Sub
'
'Private Sub txtSuppCaption2_LostFocus()
'   txtSuppCaption2.Visible = False
'   lblSupplementary2.Caption = IIf(txtSuppCaption2.text = "", lblSupplementary2.Caption, txtSuppCaption2.text)
'End Sub
'
'Private Sub txtSuppCaption3_LostFocus()
'   txtSuppCaption3.Visible = False
'   lblSupplementary3.Caption = IIf(txtSuppCaption3.text = "", lblSupplementary3.Caption, txtSuppCaption3.text)
'End Sub

Private Sub txtUnitUtilityStDt_Change()
   TextBoxChangeDate txtUnitUtilityStDt
End Sub

Private Sub txtUnitUtilityStDt_GotFocus()
   SelTxtInCtrl txtUnitUtilityStDt
End Sub

Private Sub txtUnitUtilityStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtUnitUtilityStDt, KeyAscii
End Sub

Private Sub txtUnitUtilityStDt_LostFocus()
   TextBoxFormatDate txtUnitUtilityStDt
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
