VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPO2PI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Purchase Invoice"
   ClientHeight    =   11895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   23715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11895
   ScaleWidth      =   23715
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   3360
      ScaleHeight     =   2895
      ScaleWidth      =   4815
      TabIndex        =   79
      Top             =   7800
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
         TabIndex        =   80
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   2175
         Index           =   0
         Left            =   15
         TabIndex        =   81
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   82
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
      Height          =   7695
      Index           =   0
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CheckBox chkSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDFDF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   90
         Top             =   2565
         Width           =   255
      End
      Begin VB.TextBox txtUnit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1140
         Width           =   2775
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
         Left            =   8265
         TabIndex        =   44
         Top             =   1155
         Width           =   255
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
         Left            =   8265
         TabIndex        =   43
         Top             =   1455
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
         Left            =   2235
         TabIndex        =   42
         Top             =   1740
         Width           =   255
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
         Left            =   3960
         TabIndex        =   41
         Top             =   1155
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
         Left            =   10605
         TabIndex        =   40
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtPICNTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   4785
         Width           =   975
      End
      Begin VB.TextBox txtVat_ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   10860
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtNet_ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   10200
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   1140
         Width           =   1875
      End
      Begin VB.TextBox txtDetails_ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6000
         MaxLength       =   80
         TabIndex        =   36
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox txtPICNNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   4785
         Width           =   975
      End
      Begin VB.TextBox txtPICNVat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   4785
         Width           =   735
      End
      Begin VB.Frame fraCmds 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   6960
         Width           =   12255
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   0
            Left            =   5761
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   120
            Width           =   1450
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00FFFFFF&
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   0
            Left            =   10665
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   120
            Width           =   1450
         End
         Begin VB.CommandButton cmdSavePI 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4200
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   1450
         End
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   10950
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2040
         Width           =   1120
      End
      Begin VB.Frame fraLay 
         BackColor       =   &H00DFDFDF&
         Height          =   975
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12495
         Begin VB.TextBox txtInv 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3795
            MaxLength       =   20
            TabIndex        =   17
            Top             =   600
            Width           =   2620
         End
         Begin VB.TextBox txtAc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   900
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
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txtTransType 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7260
            TabIndex        =   13
            Top             =   240
            Width           =   920
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
            Left            =   11775
            TabIndex        =   11
            Top             =   620
            Width           =   255
         End
         Begin VB.TextBox txtDueDate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7260
            TabIndex        =   10
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox txtProperty 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9120
            TabIndex        =   12
            Top             =   600
            Width           =   2925
         End
         Begin MSForms.TextBox cmbSC 
            Height          =   285
            Left            =   1500
            TabIndex        =   89
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
            Left            =   4635
            TabIndex        =   28
            Top             =   240
            Width           =   1785
            VariousPropertyBits=   679495709
            BorderStyle     =   1
            Size            =   "3149;503"
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
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reference:"
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
            Index           =   0
            Left            =   3000
            TabIndex        =   26
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trans Type:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
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
            Index           =   0
            Left            =   6480
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A/C:"
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
            Index           =   0
            Left            =   3000
            TabIndex        =   23
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
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
            Index           =   16
            Left            =   8400
            TabIndex        =   22
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date:"
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
            Index           =   19
            Left            =   6480
            TabIndex        =   21
            Top             =   600
            Width           =   705
         End
         Begin MSForms.ComboBox cboClientPI 
            Height          =   285
            Left            =   9120
            TabIndex        =   20
            Top             =   240
            Width           =   2925
            VariousPropertyBits=   1753237531
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5159;503"
            TextColumn      =   2
            ColumnCount     =   3
            ListRows        =   20
            cColumnInfo     =   3
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1411;3527;0"
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
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
            Left            =   8400
            TabIndex        =   19
            Top             =   240
            Width           =   465
         End
         Begin MSForms.Label lblPostingDate 
            Height          =   285
            Left            =   8160
            TabIndex        =   18
            Top             =   240
            Width           =   215
            ForeColor       =   8421504
            BackColor       =   16761024
            Caption         =   " P"
            Size            =   "379;503"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit the Line"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4740
         Width           =   1440
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete the Line"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4740
         Width           =   1450
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1120
      End
      Begin VB.TextBox txtNC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
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
         Left            =   2235
         TabIndex        =   4
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkRecover 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDFDF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1740
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1740
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPI 
         Height          =   1545
         Left            =   90
         TabIndex        =   45
         Top             =   3120
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   2725
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxInvoiced 
         Height          =   1320
         Left            =   90
         TabIndex        =   93
         Top             =   5490
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   2328
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoiced:"
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
         Index           =   7
         Left            =   90
         TabIndex        =   94
         Top             =   5265
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Ref:"
         Height          =   195
         Index           =   6
         Left            =   495
         TabIndex        =   92
         Top             =   2595
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSForms.TextBox txtInvoiceRef 
         Height          =   255
         Left            =   1440
         TabIndex        =   91
         Top             =   2565
         Visible         =   0   'False
         Width           =   2745
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4842;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPFName 
         Height          =   285
         Left            =   2490
         TabIndex        =   77
         Top             =   1740
         Width           =   1725
         VariousPropertyBits=   679495709
         BorderStyle     =   1
         Size            =   "3043;503"
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblVatCode 
         Height          =   255
         Index           =   0
         Left            =   10080
         TabIndex        =   76
         Top             =   1485
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
         Left            =   10200
         TabIndex        =   75
         Top             =   1740
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
         Index           =   3
         Left            =   9600
         TabIndex        =   74
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule ID:"
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
         Left            =   5055
         TabIndex        =   73
         Top             =   1440
         Width           =   885
      End
      Begin MSForms.TextBox txtJobNo 
         Height          =   285
         Left            =   6000
         TabIndex        =   72
         Top             =   1140
         Width           =   2535
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4471;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job No:"
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
         Index           =   1
         Left            =   5055
         TabIndex        =   71
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   70
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net:"
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
         Index           =   0
         Left            =   9600
         TabIndex        =   69
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Index           =   0
         Left            =   5055
         TabIndex        =   68
         Top             =   1740
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/C:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   66
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
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
         Index           =   0
         Left            =   9600
         TabIndex        =   65
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No."
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
         Index           =   3
         Left            =   240
         TabIndex        =   64
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
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
         Index           =   4
         Left            =   6840
         TabIndex        =   63
         Top             =   2880
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   5
         Left            =   7695
         TabIndex        =   62
         Top             =   2880
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
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
         Index           =   6
         Left            =   8160
         TabIndex        =   61
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/C"
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
         Index           =   7
         Left            =   690
         TabIndex        =   60
         Top             =   2880
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Left            =   4560
         TabIndex        =   59
         Top             =   2880
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Left            =   5520
         TabIndex        =   58
         Top             =   2880
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
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
         Left            =   1305
         TabIndex        =   57
         Top             =   2880
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net"
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
         Index           =   11
         Left            =   9570
         TabIndex        =   56
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T/C"
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
         Left            =   10200
         TabIndex        =   55
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
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
         Left            =   10695
         TabIndex        =   54
         Top             =   2880
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   14
         Left            =   11505
         TabIndex        =   53
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total to Invoice:"
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
         Index           =   18
         Left            =   7155
         TabIndex        =   52
         Top             =   4785
         Width           =   1125
      End
      Begin MSForms.TextBox txtNCName 
         Height          =   285
         Left            =   2490
         TabIndex        =   51
         Top             =   1440
         Width           =   1725
         VariousPropertyBits=   679495705
         BorderStyle     =   1
         Size            =   "3043;503"
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSchedules 
         Height          =   285
         Left            =   6000
         TabIndex        =   50
         Top             =   1440
         Width           =   2535
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "4471;503"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recoverable:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   2040
         Width           =   915
      End
      Begin MSForms.TextBox txtRecoverable 
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   48
         Top             =   2040
         Width           =   495
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "873;450"
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
         Index           =   5
         Left            =   2280
         TabIndex        =   47
         Top             =   2085
         Width           =   135
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0FFFF&
         Height          =   195
         Index           =   49
         Left            =   120
         TabIndex        =   78
         Top             =   2880
         Width           =   12300
      End
   End
End
Attribute VB_Name = "frmPO2PI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PO_Ref        As String
Public szPropertyID  As String
Public bEditMode     As Boolean        'Is the PO in Edit mode?
Private Const iXflxPI = 18
Private sTextBox     As String
Dim szaSupplierBal()    As String      'Supplier   balance
Dim iDayTerms           As String
Private nTaxCode     As Double             'Tax code for Invoice
Private sVCFound     As Single             'Vat code found either in Supplier (1) or Global Data (2)
Private iCurEditRow  As Integer
Dim sAddChoice As String
Dim iSelected As Integer

Private Enum ComponentMode
   DefaultMode = 0
   NewLine = 1
   EditLine = -1
   GridLostFocus = -2
   GridRowOnSelection = 2
   SavedMode = 3
   RefundMode = -3
   ExpensesMode = 4
End Enum

Private Sub chkSelect_Change()
Debug.Print chkSelect.Value
End Sub

Private Sub chkSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtInvoiceRef.Visible = chkSelect.Value
   Label12(6).Visible = chkSelect.Value
   HLAllGridRowsT flxPI, chkSelect.Value
   
   If chkSelect.Value Then
      iSelected = flxPI.Rows - 1
   Else
      iSelected = 0
   End If
   UpdateTotalPICN
End Sub

Private Sub flxPI_Click()
   Dim X As Single
   
   X = HLGridRowT(flxPI, flxPI.row)
   
   iSelected = iSelected + X
   
   If chkSelect.Value = 1 And X < 0 Then chkSelect.Value = 0
   If chkSelect.Value = 0 And iSelected = flxPI.Rows - 1 Then chkSelect.Value = 1
   If iSelected = 0 Then
      txtInvoiceRef.Visible = False
      Label12(6).Visible = False
   End If
   If iSelected > 0 Then
      txtInvoiceRef.Visible = True
      Label12(6).Visible = True
   End If
   UpdateTotalPICN
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
   flxSupplier(0).SetFocus
End Sub

Private Sub cmdCancel_Click(Index As Integer)
   If MsgBox("Do you want to cancel?" & Chr(13) & "If you wish to save the data you already entered click No", vbQuestion + vbYesNo, "Add Record") = vbNo Then Exit Sub

   PIComponents DefaultMode

   HandleCommandButton "Cancel"
   flxPI.Enabled = True
   flxPI.col = 0
   flxPI.CellBackColor = vbWhite
End Sub

Private Sub cmdClose_Click(Index As Integer)
   If Not cmdEdit.Enabled And cmdSavePI.Visible Then
      If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Prestige") = vbYes Then
         If cmdSavePI.Enabled Then cmdSavePI.SetFocus
         Exit Sub
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

   If flxPI.Rows = 2 And flxPI.row = 1 Then
      ConfigFlxPI
   End If

   If flxPI.Rows > 2 Then
      For iRow = iCurEditRow To flxPI.Rows - 2
         For iCol = 1 To flxPI.Cols - 1
            flxPI.TextMatrix(iRow, iCol) = flxPI.TextMatrix(iRow + 1, iCol)
         Next iCol
      Next iRow

      flxPI.RemoveItem flxPI.Rows - 1
   End If
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
   fraList.Visible = False
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
         PIComponents NewLine
      End With

      UpdateTotalPICN

      cmdEdit.Enabled = True
      If txtProperty.text = "" Then
         cmdNCList.SetFocus
      Else
         cmdUnitList.SetFocus
      End If
   End If

   If Index = 2 Then                   'Clear
      PIComponents EditLine
      flxPI.Enabled = True

      If txtProperty.text = "" Then
         cmdNCList.SetFocus
      Else
         cmdUnitList.SetFocus
      End If

      If txtProperty.text = "" Then
         cmdNCList.SetFocus
      Else
         cmdUnitList.SetFocus
      End If
      cmdEdit.Enabled = True
   End If
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

   If sTextBox = "A/C" Then
'      bTotalPayTyped = False
      txtAc(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 0)
      If cmbSC.text = "Client" Then
         cboClientPI.Value = txtAc(0).text
         cboClientPI.Locked = True
      Else
         cboClientPI.Locked = False
      End If

      txtSupplierName.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
      If txtNC(0).text = "" Then _
         txtNC(0).text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
      
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
      cmdUnitList.SetFocus
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

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, cboClientPI.Value
End Sub
Private Sub RefreshPO(adoConn As ADODB.Connection)
   If IsLoadedAndVisible("frmPO") Then
      frmPO.LoadFlxPurchase adoConn
   End If
End Sub
Private Sub cmdSavePI_Click()
   If iSelected = 0 Then
      ShowMsgInTaskBar "There are no purchase order lines to invoice", "Y", "N"
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim adoPIHeader As New ADODB.Recordset, adoPISplit As New ADODB.Recordset
   Dim szSQL As String, iRow As Integer, uID As String, lT_ID As Long
   Dim lSlNumber As Long, lTemp As Long

   adoConn.Open getConnectionString

'  ***************************************************************************************************
'           SAVING HEADER PART OF THE PURCHASE INVOICE                                               '
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInv"

   With adoPIHeader
      .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

      .AddNew
      uID = UniqueID()
      .Fields.Item("MY_ID").Value = uID

      lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
      .Fields.Item("SlNumber").Value = lSlNumber
      .Fields.Item("SUPP_AC").Value = txtAc(0).text
      .Fields.Item("TRAN_DATE").Value = Format(txtDate.text, "DD/MMMM/YYYY")
      .Fields.Item("TransactionType").Value = 6
      .Fields.Item("INV_NO").Value = txtInvoiceRef.text
      .Fields.Item("TOTAL_AMOUNT").Value = CCur(txtPICNTotal.text)
      .Fields.Item("TTP").Value = CByte(TransactionTakePlace("TTP", "NewPO", adoConn))
      .Fields.Item("History").Value = False
      .Fields.Item("TrfPayment").Value = True
      .Fields.Item("PropertyID").Value = szPropertyID
      .Fields.Item("CL_ID").Value = cboClientPI.Column(0)
      If Len(txtDueDate.text) = 10 Then _
         .Fields.Item("DueDate").Value = Format(txtDueDate.text, "dd mmmm yyyy")
         ''will be handle by anol 04 jan 2015 issue 469
      .Fields.Item("PostingDate").Value = Format(txtDueDate.text, "dd mmmm yyyy") 'Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")
      .Fields.Item("PO").Value = PO_Ref

      .Update
      .Close
   End With

'  ***************************************************************************************************
'           B4 SAVING SPLITS, THE PI IS EXPORTED TO PAYMENT TABLE                                    '
'  ***************************************************************************************************
   szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
   adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
   adoPIHeader.Close

   szSQL = "SELECT * FROM tlbPayment"                                         'Add New Mode

   With adoPIHeader
      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic

      .AddNew
      !TransactionID = lT_ID
      !szTransactionID = !TransactionID
      !Pi = uID

      !Type = 6
      !SageAccountNumber = txtAc(0).text
      !PDate = Format(txtDate.text, "DD MMMM YYYY")
      !dDate = Format(txtDueDate.text, "DD MMMM YYYY")
      !ref = txtInvoiceRef.text
      !ExtRef = !ref
      !amount = CCur(txtPICNTotal.text)
      !OSAmount = !amount
      !PaymentView = True
      !Details = flxPI.TextMatrix(1, 11)
      !unitid = szPropertyID
      !SlNumber = lSlNumber
      !AdjTag = "N"
      'need to be fixed by anol issue 469 04 jan 2014
      !postingDate = Format(txtDate.text, "DD MMMM YYYY") ' Format(lblPostingDate.ToolTipText, "dd mmmm yyyy")

      .Update
      .Close
   End With

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE in the PAYMENT SPLIT TABLE
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tlbPaymentSplit;"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'Add New Records. At least there is one split line.
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(iRow, 0) <> "" Then
         flxPI.row = iRow
         lTemp = flxPI.CellBackColor
         If lTemp = RGB(233, 232, 155) Then
            With adoPISplit
               .AddNew
               .Fields.Item("TransactionID").Value = UniqueID()
               .Fields.Item("PayHeader").Value = lT_ID
               .Fields.Item("FundID").Value = flxPI.TextMatrix(iRow, 8)
               .Fields.Item("Amount").Value = CCur(flxPI.TextMatrix(iRow, 14)) + _
                                              CCur(flxPI.TextMatrix(iRow, 12))
               .Fields.Item("OSAmount").Value = .Fields.Item("Amount").Value
               .Fields.Item("SplitID").Value = flxPI.TextMatrix(iRow, 0)
               .Fields.Item("DueDate").Value = Format(txtDueDate.text, "DD MMMM YYYY")
               .Fields.Item("Description").Value = flxPI.TextMatrix(iRow, 11)
               .Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
               .Fields.Item("NOMINAL_CODE").Value = flxPI.TextMatrix(iRow, 7)
               .Fields.Item("TRANS").Value = flxPI.TextMatrix(iRow, 3)
               .Fields.Item("UNIT_ID").Value = flxPI.TextMatrix(iRow, 21)
               .Fields.Item("ScheduleID").Value = IIf(flxPI.TextMatrix(iRow, 20) = "", Null, _
                                                      flxPI.TextMatrix(iRow, 20))
               .Fields.Item("RecoverablePt").Value = flxPI.TextMatrix(iRow, 22)
               .Fields.Item("AllocTranID").Value = flxPI.TextMatrix(iRow, 23)
   
               .Update
            End With
         End If
      End If
   Next iRow
'
''  User has deleted all split line.
'   If flxPI.TextMatrix(1, 0) = "" Then
'      With adoPISplit
'         .AddNew
'         .Fields.Item("TransactionID").Value = UniqueID()
'         .Fields.Item("PayHeader").Value = lT_ID
'         .Fields.Item("FundID").Value = 0
'         .Fields.Item("Amount").Value = 0
'         .Fields.Item("OSAmount").Value = 0
'         .Fields.Item("SplitID").Value = 1
'         .Fields.Item("DueDate").Value = Format(txtDueDate.text, "DD MMMM YYYY")
'         .Fields.Item("Description").Value = "ALL SPLIT DELETED"
'         .Fields.Item("NOMINAL_CODE").Value = "0000"
'         .Fields.Item("ScheduleID").Value = 0
'         .Update
'      End With
'   End If
'
   adoPISplit.Close

'  ***************************************************************************************************
'           SAVING SPLITS OF THE PURCHASE INVOICE
'  ***************************************************************************************************
   szSQL = "SELECT * FROM tblPurInvSRec"
   adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic

'Add New Records. At least there is only one split line
   For iRow = 1 To flxPI.Rows - 1
      If flxPI.TextMatrix(iRow, 0) <> "" Then
         flxPI.row = iRow
         lTemp = flxPI.CellBackColor
         If lTemp = RGB(233, 232, 155) Then
            With adoPISplit
               .AddNew
               .Fields.Item("MY_ID").Value = UniqueID()
               .Fields.Item("ParentID").Value = uID
               .Fields.Item("TRAN_ID").Value = flxPI.TextMatrix(iRow, 0)
               .Fields.Item("TRANS").Value = flxPI.TextMatrix(iRow, 3)
               .Fields.Item("UNIT_ID").Value = flxPI.TextMatrix(iRow, 21)
               .Fields.Item("NOMINAL_CODE").Value = flxPI.TextMatrix(iRow, 7)
               .Fields.Item("DEPT_ID").Value = flxPI.TextMatrix(iRow, 8)
               .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
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
               .Fields.Item("PoPiCross").Value = flxPI.TextMatrix(iRow, 23)            'PI split line CrossRef PO split line
   
               If flxPI.TextMatrix(iRow, 9) <> "" Then
                  UpdateJobActualCost .Fields.Item("JOB_ID").Value, .Fields.Item("NET_AMOUNT").Value, adoConn
               End If
               adoConn.Execute "UPDATE tblPurInvSRec " & _
                               "SET Invoiced = TRUE, PoPiCross = '" & .Fields.Item("MY_ID").Value & "' " & _
                               "WHERE MY_ID = '" & flxPI.TextMatrix(iRow, 23) & "';"
   
               .Update
            End With
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
'         .Fields.Item("TRANS").Value = szPropertyID
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
   adoPISplit.Close

'*****  System will check the header amount with the split total  ***************************
   Call SiPi_Check(adoConn, "PI", "25876")
'--------------------------------------------------------------------------------------------
'  Export Transactions to Nominal Ledger (NLPosting table)
   Export_PInPC_2_NL adoConn
'--------------------------------------------------------------------------------------

   'adoConn.Close

   Set adoPISplit = Nothing
   Set adoPIHeader = Nothing
 'added by anol 08 Dec 2015
   RefreshPO adoConn
   ShowMsgInTaskBar "Data has been saved successfully."
   
   If adoConn.State = adStateOpen Then adoConn.Close

   Set adoConn = Nothing

   Unload Me

Exit Sub



'*************************************************************************************************************************************************************************
'   If iPIEdit = 0 Then Exit Sub
'   If MsgBox("Do you wish to create Purchase Invoices from the selected Purchase Orders?", vbQuestion + vbYesNo, "Purchase Order") = -vbNo Then Exit Sub
'
'   Dim adoPO      As New ADODB.Recordset
'   Dim adoPI      As New ADODB.Recordset
'   Dim adoPOsp    As New ADODB.Recordset
'   Dim adoPIsp    As New ADODB.Recordset
'   Dim adoConn    As New ADODB.Connection
'   Dim szSQL      As String
'   Dim iRec       As Integer
'   Dim i          As Integer
'   Dim sPI        As String
'   Dim lT_ID      As Long
'   Dim dDueDt     As Date
'
'   adoConn.Open getConnectionString
'
'   adoPI.Open "SELECT * FROM tblPurInv;", adoConn, adOpenDynamic, adLockOptimistic
'   iRec = adoPI.Fields.count
'   szSQL = "SELECT * FROM tblPurInv WHERE MY_ID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "';"
'   adoPO.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   adoPI.AddNew
'   For i = 0 To iRec - 1
'      adoPI.Fields.Item(adoPI.Fields.Item(i).Name).Value = adoPO.Fields.Item(adoPI.Fields.Item(i).Name).Value
'   Next i
'   adoPI.Fields.Item("PO").Value = adoPI.Fields.Item("MY_ID").Value
'   adoPI.Fields.Item("MY_ID").Value = UniqueID()
'   sPI = adoPI.Fields.Item("MY_ID").Value
'   adoPI.Fields.Item("SlNumber").Value = SlNumber("PI", "tblPurInv", adoConn)
'   adoPI.Fields.Item("TransactionType").Value = 6
'   adoPO.Close
'
''  B4 saving the header update the tlbPayment table
'   szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
'   adoPO.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   lT_ID = CLng(IIf(IsNull(adoPO!TID), 1, adoPO!TID + 1))
'   adoPO.Close
'   szSQL = "SELECT * FROM tlbPayment;"                                         'Add New Mode
'   With adoPO
'      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'      .AddNew
'      !TransactionID = lT_ID
'      !szTransactionID = !TransactionID
'      !Pi = adoPI.Fields.Item("MY_ID").Value
'      !Type = 6
'      !SageAccountNumber = adoPI.Fields.Item("SUPP_AC").Value
'      !PDate = adoPI.Fields.Item("TRAN_DATE").Value
'      !dDate = adoPI.Fields.Item("DueDate").Value
'      dDueDt = !dDate
'      !Ref = adoPI.Fields.Item("INV_NO").Value
'      !ExtRef = !Ref
'      !Amount = adoPI.Fields.Item("TOTAL_AMOUNT").Value
'      !OSAmount = !Amount
'      !PaymentView = True
'      !Details = "Purchase Invoice"
'      !unitid = adoPI.Fields.Item("PropertyID").Value
'      !SlNumber = adoPI.Fields.Item("SlNumber").Value
'      !AdjTag = "N"
'      !PostingDate = adoPI.Fields.Item("PostingDate").Value
'
'      .Update
'      .Close
'   End With
'
'   adoPI.Update
'   adoPI.Close
'
'   adoPIsp.Open "SELECT * FROM tblPurInvSRec;", adoConn, adOpenDynamic, adLockOptimistic
'   iRec = adoPIsp.Fields.count
'   szSQL = "SELECT * FROM tblPurInvSRec WHERE ParentID = '" & flxPurchase.TextMatrix(flxPurchase.row, 0) & "';"
'   adoPOsp.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'   szSQL = "SELECT * FROM tlbPaymentSplit"
'   adoPI.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'   While Not adoPOsp.EOF
'      adoPIsp.AddNew
'      For i = 0 To iRec - 1
'         adoPIsp.Fields.Item(adoPIsp.Fields.Item(i).Name).Value = adoPOsp.Fields.Item(adoPIsp.Fields.Item(i).Name).Value
'      Next i
'      adoPIsp.Fields.Item("MY_ID").Value = UniqueID()
'      adoPIsp.Fields.Item("ParentID").Value = sPI
'
'      adoPI.AddNew
'      adoPI.Fields.Item("TransactionID").Value = UniqueID()
'      adoPI.Fields.Item("PayHeader").Value = lT_ID
'      adoPI.Fields.Item("FundID").Value = adoPIsp.Fields.Item("DEPT_ID").Value
'      adoPI.Fields.Item("Amount").Value = adoPIsp.Fields.Item("TOTAL_AMOUNT").Value
'      adoPI.Fields.Item("OSAmount").Value = adoPI.Fields.Item("Amount").Value
'      adoPI.Fields.Item("SplitID").Value = adoPIsp.Fields.Item("TRAN_ID").Value
'      adoPI.Fields.Item("DueDate").Value = Format(dDueDt, "DD MMMM YYYY")
'      adoPI.Fields.Item("Description").Value = adoPIsp.Fields.Item("DESCRIPTION").Value
'      adoPI.Fields.Item("JobID").Value = adoPIsp.Fields.Item("JOB_ID").Value            'Job No
'      adoPI.Fields.Item("NOMINAL_CODE").Value = adoPIsp.Fields.Item("NOMINAL_CODE").Value
'      adoPI.Fields.Item("TRANS").Value = adoPIsp.Fields.Item("TRANS").Value
'      adoPI.Fields.Item("UNIT_ID").Value = adoPIsp.Fields.Item("UNIT_ID").Value
'      adoPI.Fields.Item("ScheduleID").Value = adoPIsp.Fields.Item("ScheduleID").Value
'      adoPI.Fields.Item("RecoverablePt").Value = adoPIsp.Fields.Item("RecoverablePt").Value
'      adoPI.Fields.Item("AllocTranID").Value = adoPIsp.Fields.Item("MY_ID").Value
'
'      adoPI.Update
'
'      adoPIsp.Update
'      adoPOsp.MoveNext
'   Wend
'   adoPOsp.Close
'   adoPIsp.Close
'   adoPI.Close
'
'   adoConn.Close
'   Set adoConn = Nothing
'
'   ShowMsgInTaskBar "Purchase Invoice has been created", "Y", "P"
End Sub

Private Sub UpdateJobActualCost(szJobId As String, cCost As Currency, adoConn As ADODB.Connection)
   Dim szSQL As String

   If txtTransType.text = "Invoice" Then
      szSQL = "SET ActualCost = ActualCost + " & cCost & " "
   Else
      szSQL = "SET ActualCost = ActualCost - " & cCost & " "
   End If

   szSQL = "UPDATE PropertyMaintHistory " & szSQL & _
           "WHERE ID = '" & szJobId & "';"

   adoConn.Execute szSQL
End Sub

Private Sub Form_Load()
   Me.Width = 12780
   Me.Height = 8280
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   iSelected = 0

   ConfigFlxPI
   ConfigFlxInvoiced
   If Len(txtDate.text) < 10 Then
      txtDate.text = Format(Date, "dd/mm/yyyy")
      lblPostingDate.ToolTipText = txtDate.text
   End If
   
   Dim adoConn As New ADODB.Connection

'   connect to database
   adoConn.Open getConnectionString
   LoadCboClientPI adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub LoadCboClientPI(adoConn As ADODB.Connection)
   Dim szSQL   As String
   Dim adoRst  As New ADODB.Recordset
   
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT " & _
           "FROM   CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

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
   cboClientPI.Column() = Data()
   cboClientPI.ListIndex = 0

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmPO.Enabled = True
End Sub

Private Sub ConfigFlxPI()
   With flxPI
      .Clear
      .Cols = 26
      .Rows = 2
      .RowHeight(0) = 0

      .ColWidth(0) = Label7(7).Left - Label7(3).Left '"TransactionID"
      .ColWidth(1) = 0 'Label7(5).Left - Label7(4).Left '"A/C"
      .ColWidth(2) = 0 'Label7(6).Left - Label7(5).Left '"Date"
      .ColWidth(3) = 0 'Label7(7).Left - Label7(6).Left '"Type"
      .ColWidth(4) = 0 'Label7(8).Left - Label7(7).Left '"Trans"
      .ColWidth(5) = 0 'Label7(9).Left - Label7(8).Left '"Unit ID + Name"
      .ColWidth(6) = 0 'Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
      .ColWidth(7) = Label7(10).Left - Label7(7).Left                    '"N/C"
      .ColWidth(8) = 0                      '"Dept"
      .ColWidth(9) = 0                      '"Job No"
      .ColWidth(10) = 0                     '"Cost Code"
      .ColWidth(11) = Label7(11).Left - Label7(10).Left '"Details"
      .ColWidth(12) = Label7(12).Left - Label7(11).Left '"Net"
      .ColWidth(13) = Label7(13).Left - Label7(12).Left '"T/C"
      .ColWidth(14) = Label7(14).Left - Label7(13).Left '"VAT"
      .ColWidth(15) = .Width - Label7(14).Left - 120 '"Total"
      .ColWidth(16) = 0                     '"Sage"
      .ColWidth(17) = 0           'Stores PI Id hidenly
      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
      .ColWidth(19) = 0           'keep value 0 or 1 for edit
      .ColWidth(20) = 0           'Stores ScheduleId
      .ColWidth(21) = 0           'Stores Unit ID
      .ColWidth(22) = 0           '% Recoverable
      .ColWidth(23) = 0           'ID
      .ColWidth(24) = 0           'FundCode
      .ColWidth(25) = 0           'FundName
      
      
      .ColWidth(0) = Label7(7).Left - Label7(3).Left '"TransactionID"
      .ColWidth(1) = 0 'Label7(5).Left - Label7(4).Left '"A/C"
      .ColWidth(2) = 0 'Label7(6).Left - Label7(5).Left '"Date"
      .ColWidth(3) = 0 'Label7(7).Left - Label7(6).Left '"Type"
      .ColWidth(4) = 0 'Label7(8).Left - Label7(7).Left '"Trans"
      .ColWidth(5) = 0 'Label7(9).Left - Label7(8).Left '"Unit ID + Name"
      .ColWidth(6) = 0 'Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
      .ColWidth(7) = Label7(10).Left - Label7(7).Left + 20                   '"N/C"
      .ColWidth(8) = 0                      '"Dept"
      .ColWidth(9) = 0                      '"Job No"
      .ColWidth(10) = 0                     '"Cost Code"
      .ColWidth(11) = Label7(11).Left - Label7(10).Left + 20 '"Details"
      .ColWidth(12) = Label7(12).Left - Label7(11).Left + 20 '"Net"
      .ColWidth(13) = Label7(13).Left - Label7(12).Left + 20 '"T/C"
      .ColWidth(14) = Label7(14).Left - Label7(13).Left + 20 '"VAT"
      .ColWidth(15) = .Width - Label7(14).Left + 120 '"Total"
      .ColWidth(16) = 0                     '"Sage"
      .ColWidth(17) = 0           'Stores PI Id hidenly
      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
      .ColWidth(19) = 0           'keep value 0 or 1 for edit
      .ColWidth(20) = 0           'Stores ScheduleId
      .ColWidth(21) = 0           'Stores Unit ID
      .ColWidth(22) = 0           '% Recoverable
      .ColWidth(23) = 0           'ID
      .ColWidth(24) = 0           'FundCode
      .ColWidth(25) = 0           'FundName
      
      
      
      .row = 0
   End With

   txtPICNNet.Left = Label7(11).Left
   txtPICNNet.Width = flxPI.ColWidth(12)
   txtPICNVat.Left = Label7(13).Left
   txtPICNVat.Width = flxPI.ColWidth(14)
   txtPICNTotal.Left = Label7(14).Left
   txtPICNTotal.Width = flxPI.ColWidth(15)
End Sub

Private Sub ConfigFlxInvoiced()
   With flxInvoiced
      .Clear
      .Cols = 26
      .Rows = 2
      .RowHeight(0) = 0
      
      .ColWidth(0) = Label7(7).Left - Label7(3).Left '"TransactionID"
      .ColWidth(1) = 0 'Label7(5).Left - Label7(4).Left '"A/C"
      .ColWidth(2) = 0 'Label7(6).Left - Label7(5).Left '"Date"
      .ColWidth(3) = 0 'Label7(7).Left - Label7(6).Left '"Type"
      .ColWidth(4) = 0 'Label7(8).Left - Label7(7).Left '"Trans"
      .ColWidth(5) = 0 'Label7(9).Left - Label7(8).Left '"Unit ID + Name"
      .ColWidth(6) = 0 'Label7(10).Left - Label7(9).Left 'Inv No / Cr. No
      .ColWidth(7) = Label7(10).Left - Label7(7).Left + 20                  '"N/C"
      .ColWidth(8) = 0                      '"Dept"
      .ColWidth(9) = 0                      '"Job No"
      .ColWidth(10) = 0                     '"Cost Code"
      .ColWidth(11) = Label7(11).Left - Label7(10).Left + 20 '"Details"
      .ColWidth(12) = Label7(12).Left - Label7(11).Left + 20 '"Net"
      .ColWidth(13) = Label7(13).Left - Label7(12).Left + 20 '"T/C"
      .ColWidth(14) = Label7(14).Left - Label7(13).Left + 20 '"VAT"
      .ColWidth(15) = .Width - Label7(14).Left + 120 '"Total"
      .ColWidth(16) = 0                     '"Sage"
      .ColWidth(17) = 0           'Stores PI Id hidenly
      .ColWidth(iXflxPI) = 0      'Marked X when row will be selected  iX = 18
      .ColWidth(19) = 0           'keep value 0 or 1 for edit
      .ColWidth(20) = 0           'Stores ScheduleId
      .ColWidth(21) = 0           'Stores Unit ID
      .ColWidth(22) = 0           '% Recoverable
      .ColWidth(23) = 0           'ID
      .ColWidth(24) = 0           'FundCode
      .ColWidth(25) = 0           'FundName
      
      
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
      .row = 0
   End With
'
'   txtPICNNet.Left = Label7(11).Left
'   txtPICNNet.Width = flxPI.ColWidth(12)
'   txtPICNVat.Left = Label7(13).Left
'   txtPICNVat.Width = flxPI.ColWidth(14)
'   txtPICNTotal.Left = Label7(14).Left
'   txtPICNTotal.Width = flxPI.ColWidth(15)
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
End Sub

Private Sub txtDate_GotFocus()
   If txtDate.text = "dd/mm/yyyy" Then
      txtDate.text = ""
      Exit Sub
   End If
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   On Error Resume Next

   If txtDate.text <> "" Then TextBoxFormatDate txtDate
   If txtDate.text <> "" Then txtDueDate.text = txtDate.text
   lblPostingDate.ToolTipText = txtDate.text
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
            .TextMatrix(iRow, 5) = IIf(IsNull(rstRst!PaymentTerms), "", rstRst!PaymentTerms)
            rstRst.MoveNext
            If Not rstRst.EOF Then .AddItem ""
            iRow = iRow + 1
         Wend
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
               .TextMatrix(iRow, 0) = rstRst!clientID
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
              "WHERE N.ClientID = '" & cboClientPI.Value & "' AND " & _
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
           "WHERE ClientID = '" & cboClientPI.Value & "' " & _
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

Private Sub txtNet__GotFocus(Index As Integer)
   txtNet_(Index).SelStart = 0
   txtNet_(Index).SelLength = Len(txtNet_(Index).text)
End Sub

Private Sub txtNet__KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 10 Then txtNet__LostFocus (0)

   DigitTextKeyPress txtNet_(0), KeyAscii
End Sub

Private Sub txtNet__LostFocus(Index As Integer)
   txtVat_(0).text = Format(IIf(txtNet_(0).text = "", 0, Val(txtNet_(0).text)) * (nTaxCode / 100), "0.00")
   txtNet_(0).text = Format(txtNet_(0).text, "0.00")
   txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
   txtTotal.text = Format(txtTotal.text, "0.00")
End Sub

Private Sub txtVat__LostFocus(Index As Integer)
   txtTotal.text = Val(txtVat_(0).text) + Val(txtNet_(0).text)
   txtTotal.text = Format(txtTotal.text, "0.00")
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

Private Sub UpdateTotalPICN()
   Dim i As Integer, lTemp As Long, r As Integer

   txtPICNNet.text = "0"
   txtPICNVat.text = "0"
   txtPICNTotal.text = "0"
   r = flxPI.row

   For i = 1 To flxPI.Rows - 1
      flxPI.row = i
      lTemp = flxPI.CellBackColor
      If lTemp = RGB(233, 232, 155) Then
         txtPICNNet.text = Val(txtPICNNet.text) + Val(flxPI.TextMatrix(i, 12))
         txtPICNVat.text = Val(txtPICNVat.text) + Val(flxPI.TextMatrix(i, 14))
         txtPICNTotal.text = Val(txtPICNTotal.text) + Val(flxPI.TextMatrix(i, 15))
      End If
   Next i

   txtPICNNet.text = Format(txtPICNNet.text, "0.00")
   txtPICNVat.text = Format(txtPICNVat.text, "0.00")
   txtPICNTotal.text = Format(txtPICNTotal.text, "0.00")
   
   flxPI.row = r
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

   Case ComponentMode.NewLine
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

Private Sub HandleCommandButton(szButton As String)
   Select Case szButton
      Case "Save"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdCancel(0).Enabled = False

         flxPI.Enabled = True
         flxPI.col = 0
         flxPI.CellBackColor = vbWhite

         ConfigFlxPI

      Case "Add Invoice"
         cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = True
         cmdCancel(0).Enabled = True

      Case "Edit"
         cmdUpdate(1).Enabled = True
         cmdSavePI.Enabled = False
         cmdCancel(0).Enabled = False

      Case "Cancel"
         cmdUpdate(1).Enabled = False
         cmdSavePI.Enabled = False
         cmdCancel(0).Enabled = False

      Case "Update Record"
         cmdSavePI.Enabled = True
         cmdCancel(0).Enabled = True
         flxPI.Enabled = True
         flxPI.row = 0
   End Select
End Sub

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

