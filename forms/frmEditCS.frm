VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEditCS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "edit CS"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   24
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Display"
      Height          =   315
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   600
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demand Details: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   14445
      Begin VB.CommandButton Command1 
         Caption         =   "Save Changes"
         Height          =   435
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2400
         Width           =   1800
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   11520
         TabIndex        =   19
         Top             =   315
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   10260
         TabIndex        =   18
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   9105
         TabIndex        =   17
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   7785
         TabIndex        =   16
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   9180
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   8010
         TabIndex        =   14
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   6855
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   5715
         TabIndex        =   12
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2130
         Left            =   75
         TabIndex        =   21
         Top             =   210
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   3757
         _Version        =   393216
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483630
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
      End
   End
   Begin VB.Frame fraDetails 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PI Details: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   14445
      Begin VB.TextBox txtArrayEditDemand 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   5715
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   6855
         TabIndex        =   8
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   8010
         TabIndex        =   7
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtArrayEditDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   9180
         TabIndex        =   6
         Top             =   630
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txtSDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   7785
         TabIndex        =   5
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPayDt 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   9105
         TabIndex        =   4
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPostingDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   10260
         TabIndex        =   3
         Top             =   315
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAEADA&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   11520
         TabIndex        =   2
         Top             =   315
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave1 
         Caption         =   "Save Changes"
         Height          =   315
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   1920
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxChildDemands 
         Height          =   2130
         Left            =   75
         TabIndex        =   10
         Top             =   210
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   3757
         _Version        =   393216
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483630
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
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Cs Number"
      Height          =   255
      Left            =   2040
      TabIndex        =   23
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmEditCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
