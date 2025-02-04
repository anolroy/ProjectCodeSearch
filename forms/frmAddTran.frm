VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmImpAddTran 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8565
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAddTran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   0
      Left            =   360
      ScaleHeight     =   2535
      ScaleWidth      =   4815
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   4845
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   4935
         TabIndex        =   40
         Top             =   -120
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFund 
            Height          =   1935
            Left            =   30
            TabIndex        =   41
            Top             =   660
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3413
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
         Begin VB.Label lblFlxPayee 
            Caption         =   "EMPTY"
            Height          =   255
            Index           =   0
            Left            =   2115
            TabIndex        =   47
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblPayeeFlxConfigured 
            Caption         =   "NOT"
            Height          =   495
            Index           =   0
            Left            =   1515
            TabIndex        =   46
            Top             =   1680
            Width           =   1095
         End
         Begin MSForms.Label lblSearch0 
            Height          =   195
            Index           =   0
            Left            =   50
            TabIndex        =   45
            Top             =   120
            Width           =   1400
            VariousPropertyBits=   8388627
            Caption         =   "ID"
            Size            =   "2469;344"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label lblSearch0 
            Height          =   195
            Index           =   1
            Left            =   1560
            TabIndex        =   44
            Top             =   135
            Width           =   735
            VariousPropertyBits=   8388627
            Caption         =   "Fund Name"
            Size            =   "1296;353"
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSearch1 
            Height          =   255
            Left            =   30
            TabIndex        =   43
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
         Begin MSForms.TextBox txtSearch2 
            Height          =   255
            Left            =   1350
            TabIndex        =   42
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
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Index           =   0
            Left            =   0
            Top             =   150
            Width           =   4500
         End
      End
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
         Index           =   1
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picAccList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   6345
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtAccTypeSearch 
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
         Left            =   4000
         TabIndex        =   33
         Top             =   300
         Width           =   1935
      End
      Begin VB.CommandButton cmdAccListClose 
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
         TabIndex        =   32
         Top             =   20
         Width           =   255
      End
      Begin VB.TextBox txtAccSearchID 
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
         TabIndex        =   31
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtAccSearchName 
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
         TabIndex        =   30
         Top             =   300
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   25
         Top             =   3240
         Visible         =   0   'False
         Width           =   6015
         Begin MSForms.ComboBox cboSrcClient 
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   29
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
         Begin MSForms.ComboBox cboSrcProp 
            Height          =   315
            Index           =   0
            Left            =   3675
            TabIndex        =   28
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
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   465
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Property:"
            Height          =   195
            Index           =   4
            Left            =   3000
            TabIndex        =   26
            Top             =   0
            Width           =   645
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccList 
         Height          =   1815
         Left            =   45
         TabIndex        =   34
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3201
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   4005
         TabIndex        =   37
         Top             =   75
         Width           =   345
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   36
         Top             =   75
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   75
         Width           =   165
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
         Top             =   80
         Width           =   6015
      End
   End
   Begin VB.PictureBox picAddTran 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5D5D5&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3065
      Left            =   0
      ScaleHeight     =   3030
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   750
         Width           =   735
      End
      Begin VB.CommandButton Command5 
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
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1050
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1350
         Width           =   735
      End
      Begin VB.CommandButton Command6 
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
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1350
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   1650
         Width           =   4695
      End
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
         Left            =   5615
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1350
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   750
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   10
         Left            =   80
         TabIndex        =   23
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   11
         Left            =   80
         TabIndex        =   22
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   12
         Left            =   80
         TabIndex        =   21
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   13
         Left            =   80
         TabIndex        =   20
         Top             =   1350
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Details:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   14
         Left            =   80
         TabIndex        =   19
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   15
         Left            =   80
         TabIndex        =   18
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Type:"
         Height          =   195
         Index           =   29
         Left            =   80
         TabIndex        =   17
         Top             =   1950
         Width           =   960
      End
      Begin MSForms.ComboBox cmbAmtType 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   1950
         Width           =   1815
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3201;503"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans Type:"
         Height          =   195
         Index           =   16
         Left            =   80
         TabIndex        =   15
         Top             =   420
         Width           =   795
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   420
         Width           =   1815
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3201;503"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   3
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705"
      End
      Begin MSForms.CommandButton cmdSPAmtType 
         Height          =   270
         Left            =   2920
         TabIndex        =   13
         Top             =   1950
         Width           =   315
         Caption         =   "; ;"
         Size            =   "556;476"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmImpAddTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdGridUnitLookup_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Activate()
'   Me.Left = .Width / 2 '- frmImpAddTran.Width
'   Me.Top = Me.Height / 2 '- frmImpAddTran.Height / 2
   Me.Height = 3150
   Me.Width = 5985
End Sub

