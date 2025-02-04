VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.MDIForm frmMMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "Prestige Property Management Program"
   ClientHeight    =   11910
   ClientLeft      =   165
   ClientTop       =   -25320
   ClientWidth     =   20250
   Icon            =   "frmMMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmDisplayTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   5160
   End
   Begin MSComctlLib.StatusBar stbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   11535
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12349
            MinWidth        =   12349
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   2240
            MinWidth        =   2240
            Text            =   "Calculator"
            TextSave        =   "Calculator"
            Key             =   "cal"
            Object.Tag             =   "calcu"
            Object.ToolTipText     =   "Calculator"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   11535
      Left            =   1680
      ScaleHeight     =   11535
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   0
      Width           =   60
   End
   Begin VB.PictureBox picTreeView 
      Align           =   3  'Align Left
      BackColor       =   &H80000009&
      Height          =   11535
      Left            =   0
      ScaleHeight     =   11475
      ScaleWidth      =   1620
      TabIndex        =   0
      Top             =   0
      Width           =   1680
      Begin VB.CommandButton cmdClientBanks 
         Caption         =   "tlbClientBanks"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   5280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   4800
         Visible         =   0   'False
         Width           =   615
      End
      Begin RichTextLib.RichTextBox rtxtMessageDisplay 
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   7080
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16761024
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmMMain.frx":F172
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   1000
         Top             =   1000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":F1F3
               Key             =   ""
               Object.Tag             =   "Client"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":FACD
               Key             =   ""
               Object.Tag             =   "Property"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":103A7
               Key             =   ""
               Object.Tag             =   "Unit"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":10C81
               Key             =   ""
               Object.Tag             =   "Lessee"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":11AD3
               Key             =   ""
               Object.Tag             =   "Tenant"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMMain.frx":11DED
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwLandLord 
         Height          =   9255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   16325
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgList"
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
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   114
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":12107
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":12421
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":12A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":12E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1319E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1378F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":13DAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":14687
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":14B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":15144
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":15705
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":15CBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":15F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":16477
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":16B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":16DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":17269
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":178C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":17EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":187CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":190A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":193B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":195C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":199C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":19FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1A62E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1AC1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1B250
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1B826
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1BC2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1C22E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1DBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1DFF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1E3DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1E994
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1EFDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1F247
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1F5F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1F858
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1FB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":1FF2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":205AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":20B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":211CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":21AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":21DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":22213
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2252D
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":22C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":231D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":234F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":23AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":24102
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2457A
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":24B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":24EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":25564
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":25BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":26181
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":26739
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":26D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":27344
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":27937
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":27EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":27FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2861E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":28AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":29112
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":294C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":29A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":29DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2A165
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2A41F
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2A7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2AAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2AF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2B4AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2B7B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2BB59
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2C1E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2C4C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2CB38
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2D18E
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2D7B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2DA76
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2DE67
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2E29D
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2E8C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2EEC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2F51C
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":2FB3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":30078
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":305F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":30BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":30FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":31220
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":31773
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":31D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":321F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":328A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":32EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":334E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":33B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":33CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":34554
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":349E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":34F65
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":35586
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":35B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":35E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":36435
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":366EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":36D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMMain.frx":3738C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuEditUserNames 
         Caption         =   "Edit &User Names"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Database..."
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuRestoreDB 
         Caption         =   "Restore Database..."
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransactions 
         Caption         =   "Transactions"
         Begin VB.Menu mnuPurchaseInvoice 
            Caption         =   "Purchase Invoice"
         End
         Begin VB.Menu mnuSales 
            Caption         =   "Sales Invoice"
         End
      End
   End
   Begin VB.Menu mnuAccounting 
      Caption         =   "Accounting"
      Begin VB.Menu mnuNominalJournal 
         Caption         =   "Nominal Journal"
      End
      Begin VB.Menu mnuNominalListing 
         Caption         =   "Nominal Listing"
      End
      Begin VB.Menu mnuProfitAndLoss 
         Caption         =   "Profit And Loss"
      End
      Begin VB.Menu mnuTrialBalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu mnuBalanceSheet 
         Caption         =   "Balance Sheet"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuNominalHistoryReport 
         Caption         =   "Nominal History Report"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuchartOfAccounts 
         Caption         =   "Chart Of Accounts"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log &Out"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'   frmMMain
'   Mother Form
'   --------------------------------
'   Created By: Samrat Rahman
'   Published Date: 23/06/2005
'   Legal Copyright: Samrat M Rahman © 23/06/2005
'*****************************************************************************************

Option Explicit

Public iSecCount        As Integer
Public Conn1            As New ADODB.Connection
Public Rst1             As New ADODB.Recordset
Public SystemUserName   As String
Public Leasee1_LesseList_isUptoDate   As Boolean
Public Leasee4_LesseList_isUptoDate   As Boolean
Public frmDemand3_LesseList_isUptoDate   As Boolean
Public frmSupplier_SupplierList_isUptoDate   As Boolean
Public frmSupplier_SupplierListBCL_isUptoDate   As Boolean
Public frmPI_SupplierBalanceByCL_isUptoDate   As Boolean
Public frmPI_SupplierBalance_isUptoDate   As Boolean
Private bLogOFF         As Boolean
Private bFixingDT       As Boolean

Dim g                   As Boolean
Dim iFORM_LOAD          As Integer
Const RibbonVersion     As Boolean = False

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
Dim Theme As Integer
Private bisformLoaded As Boolean
Dim bCascade As Boolean
Dim loopcount As Integer
Public Function IsRibbonVersion() As Boolean
   IsRibbonVersion = RibbonVersion
End Function

Private Sub ACPRibbon1_ButtonClick(ByVal Id As String, ByVal Caption As String)
    Dim strTemp As String
'    Dim bCascade As Boolea
   Select Case Id
      Case "I1"
         Call mnuBackup_Click
      Case "I2"
         Call mnuRestoreDB_Click
      Case "I3"
         Call mnuChangePassword_Click
      Case "I4"
         Call mnuEditUserNames_Click
      Case "I5"
         Call mnuCompanySetup_Click
'      Case "I6"
'         Call Import bank statements
'     case "I7"
'         Call Import Budget
      Case "I8"
         Call mnuThirdParty_Click
      Case "I9"
         Call mnuLogOut_Click
      Case "I10"
         Call mnuExit_Click
      Case "I11"
         Call mnuTreeView_Click
      Case "I12"
         Call mnuRefreshTree_Click
      Case "I13"
         Call mnuLesseePortal_Click
      Case "I14"
         Call mnuSupplierPortal_Click
      Case "I15"
         Dim retval
         retval = Shell(App.Path & "\CCD\CCD\bin\Release\CCD.exe", vbNormalFocus)
      Case "I16"
         retval = Shell(App.Path & "\MaintenanceDashboard\CCD\bin\Release\MaintenanceDashboard.exe", vbNormalFocus)
      Case "I17"
         Call mnuLSummary_Click
      Case "I18"
         Call cmdDemands_Click
      Case "I19"
         Call cmdDemands_Click
         frmDemands3.tabDmdRcpt.Tab = 2
      Case "I20"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "D" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "D"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
     Case "I122"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "D" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "MGR"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
     Case "I123"
            LoadForm frmRetentionMaster
      Case "I21"
         If TOTAL_SUPPLIERS > 0 Then
            Call cmdPurEx_Click
         Else
            MsgBox "You must first create your suppliers before you can" & Chr(13) & _
                   "enter purchase invoices and expenses in the system.", vbInformation + vbOKOnly, "No supplier found"
         End If
      Case "I22"
         If TOTAL_SUPPLIERS > 0 Then
            Call cmdPurEx_Click
            frmPurchaseExpense.tabPurExp.Tab = 1
         Else
            MsgBox "You must first create your suppliers before you can" & Chr(13) & _
                   "enter purchase invoices and expenses in the system.", vbInformation + vbOKOnly, "No supplier found"
         End If
      Case "I23"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "PnE" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "PnE"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
      Case "I24"
         Call cmdCashBook_Click
      Case "I25"
         Call mnuCR_Click
'      Case "I26"
'         call Re-run Batch
      Case "I27"
         Call mnuBatchPayment__Click
      Case "I28"
         Call mnuBPP_Click
      Case "I29"
         Call mnuPayPro_Click
      Case "I30"
         Call mnuViewPayEmails_Click
      Case "I31"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "CB" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "CB"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
      Case "I32"
         'issue 496 Nominal Journals
         'Added by anol 13 Nov 2014
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If
         LoadForm frmNJ
'         frmNJ.Show
'         frmNJ.ZOrder 0
      Case "I33"
      Case "I34"
         Call mnuManagementFees_Click
      Case "I35"
         Call mnuRentPayable_Click
      Case "I36"
         Call mnuAgedDebtors_Click
      Case "I37"
         Call mnuAgedCreditors_Click
      Case "I38"
         Call mnuRCC_Click
      Case "I39"
         Call mnuPCF_Click
      Case "I40"
         Call cmdAgent_Click
      Case "I41"
         Call cmdSupplier_Click
      Case "I42"
         ShowReport App.Path & szReportPath & "\SupplierListingReport.rpt"
      Case "I43"
         Call cmdClient_Click
      Case "I44"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "C" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "C"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
      Case "I45"
         Call cmdShopCentre_Click
      Case "I46"
         Call cmdUnits_Click
      Case "I47"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "U" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "U"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
      Case "I48"
         Call mnuPendingJobs_Click
      Case "I49"
         Call mnuJBCR_Click
      Case "I50"
         Call mnuInsSchedule_Click
      Case "I51"
         Call cmdtenants_Click
      Case "I52"
         Call mnuLesseeDetails_Click
      Case "I53"
         Call mnuGeneralLetters_Click
      Case "I54"
         Call mnuReminderLetters_Click
      Case "I55"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "L" Then Unload frmReportsMenu
         End If
         Load frmReportsMenu
         frmReportsMenu.RootName = "L"
         frmReportsMenu.Show
         frmReportsMenu.ZOrder 0
      Case "I56"
         Call cmdLease_Click
      Case "I57"
         Call mnuGLU_Click
      Case "I58"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "LA" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "LA"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
      Case "I59"
         LoadForm frmServiceCharge
'         frmServiceCharge.Show
'         frmServiceCharge.ZOrder 0
      Case "I60"
      'Resolved by BOSL
      'modified by anol 18 Sep 2014
      'issue 473
'         If IsLoadedAndVisible("frmRentBudget") Then
'            Unload frmRentBudget
'         End If
         frmRentBudget.sModule = "RB"              'Rent budget
         Load frmRentBudget
         'frmRentBudget.cboProperty_Click
         frmRentBudget.Caption = "Rent Budget Details "
         frmRentBudget.Show
         frmRentBudget.ZOrder 0
      Case "I61"
       'Resolved by BOSL
      'modified by anol 18 Sep 2014
      'issue 473
'        If IsLoadedAndVisible("frmRentBudget") Then
'            Unload frmRentBudget
'        End If
         frmRentBudget.sModule = "IB"              'Insurance Budget
         Load frmRentBudget
         'frmRentBudget.cboProperty_Click
         frmRentBudget.Caption = "Insurance Budget Details "
         frmRentBudget.Show
         frmRentBudget.ZOrder 0
      Case "I62"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "LB" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "LB"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
      Case "I63"
         Call mnuDemandTypes_Click
      Case "I64"
         Call mnuChargeTypes_Click
      Case "I65"
         Call mnuPayableTypes_Click
      Case "I66"
'         Call mnuElectricityCharge_Click
         LoadForm frmCommonUtilityCharges
'         frmCommonUtilityCharges.Show
'         frmCommonUtilityCharges.ZOrder 0

      Case "I67"
         If Not IsLoadedAndVisible("frmNominalLedger") Then
            LoadForm frmNominalLedger
            frmNominalLedger.Form_Activated = False
            If frmNominalLedger.MiLoading Then
               LoadForm frmNominalLedger
'               frmNominalLedger.Show
'               frmNominalLedger.ZOrder 0
            Else
               Unload frmNominalLedger
            End If
         Else
             frmNominalLedger.ZOrder 0
         End If
      Case "I68"
         Call mnuChartAccountsReport_Click
'      Case "I69"
'         call Financial Yearend
      Case "I70"
         Call mnuSCY_Click
      Case "I71"
         Call mnuBank_Click
      Case "I72"
         Call mnuBankTransactions_Click
'      Case "I73"
'         call Auto Reconciliation
      Case "I74"
         Call mnuProfitLoss_Click
      Case "I75"
         Call mnuBS_Click
      Case "I76"
'         Call mnuBAE_Click
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "FBR" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "FBR"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
      Case "I77"
         Call mnuTrialBalance_Click
      Case "I78"
         Call mnuVATDetailReport_Click
      Case "I79"
         Call mnuVATSummaryReport_Click
      Case "I80"
         Call mnuNHR_Click
      Case "I81"
         Call cmdGlobal_Click
      Case "I82"
         Call mnuCodes_Click
      Case "I83"
         Call mnuFund_Click
      Case "I84"
         Call mnuSchedule_Click
      Case "I85"                       'Not shown on the menu
         Call mnuControlCode_Click
      Case "I86"
         Call mnuVat_Click
      Case "I87"                                'Report Category
         Call mnuReportCategories_Click
      Case "I88"
         Call mnuGFY_Click
      Case "I89"
         LoadForm frmFinancialYear_Closing
'         frmFinancialYear_Closing.Show
'         frmFinancialYear_Closing.ZOrder 0
      Case "I90"
         Call mmuLetters_Click
      Case "I91"
         Call mnuReminderTemplates_Click
      Case "I92"
         Call mnuBACSEmailTemaplate_Click
      Case "I93"
         Call mnuDemandEmailTemplate_Click
'      Case "I94"                               'Change Lessee ID
'         Call mnuChangeLesseeID_Click
'      Case "I95"                                'Change Supplier ID
'         Call mnuChangeSupplierID_Click
'      Case "I96"                                'Change Client ID
'         Call mnuChangeClientID_Click
'      Case "I97"                                'Change Property ID
'         Call mnuChangePropertyID_Click
'      Case "I98"                                'Change Unit ID
'         Call mnuChangeUnitID_Click
      Case "I99"                                '
         LoadForm frmHelp
'         frmHelp.Show
'         frmHelp.ZOrder 0
      Case "I100"
         Call mnuAbout_Click
      Case "I101"
         Call mnuSetting_Click
      Case "I102"
         Call mnuRAS_Click
      Case "I103"
         LoadForm frmAuditTrail
'         frmAuditTrail.Show
'         frmAuditTrail.ZOrder 0
      Case "I104"
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "CID" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "CID"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
      Case "I105"
         If Not IsLoadedAndVisible("frmBudgetView") Then
            LoadForm frmBudgetView
'            frmBudgetView.Show
'            frmBudgetView.ZOrder 0
         End If
      Case "I106"
         If Not IsLoadedAndVisible("frmBudgetAlert") Then
            LoadForm frmBudgetAlert
'            frmBudgetAlert.Show
'            frmBudgetAlert.ZOrder 0
         End If
       Case "I108"
         'issue 496 Upload Receipts
         'Added by anol 13 Nov 2014
         
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If
         
      Case "I109"
         Load frmUsageHistory
         frmUsageHistory.Show
         frmUsageHistory.ZOrder 0
      Case "I110"
         LoadForm frmImportUsage
         frmImportUsage.Show
         frmImportUsage.ZOrder 0
      Case "I111"
         LoadForm frmMaintenance
         frmMaintenance.Show
         frmMaintenance.ZOrder 0
      Case "I114"
         LoadForm frmSendBackup4Support
'         frmSendBackup4Support.Show
'         frmSendBackup4Support.ZOrder 0
      Case "I115"
         LoadForm frmPO
'         frmPO.Show
'         frmPO.ZOrder 0
      Case "I116"
      'added by anol 09 July 2015
      'modified by anol 25 Nov 2015
         LoadForm frmLandLord2
'         frmLandLord2.Show
'         frmLandLord2.ZOrder 0
       Case "I117" ''added by anol 10 Aug 2016
         If IsLoadedAndVisible("frmReportsMenu") Then
            If frmReportsMenu.RootName <> "CB2" Then Unload frmReportsMenu
         End If
         LoadForm frmReportsMenu
         frmReportsMenu.RootName = "CB2"
'         frmReportsMenu.Show
'         frmReportsMenu.ZOrder 0
       Case "I118"                                '
         Load frmRollBackBankRecon
         frmRollBackBankRecon.Show
         frmRollBackBankRecon.ZOrder 0
        Case "I119"
            If Not bCascade Then
                frmMMain.Arrange vbTileHorizontal
                'frmMMain.ACPRibbon1.ButtonCenter(1).Enabled = False
                bCascade = True
            Else
               'frmMMain.Arrange vbCascade
               bCascade = False
            End If
         Case "I120"
            Call mnuStatementTemplate_Click
         Case "I121"
            LoadForm frmConsolidatedBankList
         Case "I124"
            Call mnuClientStatementTemplate_Click
         Case "I125"
            Call mnufrmAuditLog_Click
            'frmConsolidatedBankList.Show
         Case "I126"
            mnufrmViewCashBalance_Click
   End Select
'   frmMMain.Arrange vbCascade
'    Call FindForms
End Sub
Public Sub FindForms()
    Dim frx As Form, fCount As Integer
    Dim frmName
    For Each frx In Forms()
      fCount = fCount + 1
        If fCount = 1 Then
            frx.Top = 10
            frx.Left = 10
        End If
    Next
End Sub
Private Sub cmdAgent_Click()
'modified by anol 25 Nov 2015
   Load frmManagingAgent3
   frmManagingAgent3.Show
   frmManagingAgent3.ZOrder 0
End Sub

Private Sub cmdAgent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Managing Agent's detail information"
'   Me.MousePointer = vbArrow
End Sub

Private Sub cmdBankTransactions_Click()
   LoadForm frmBankTransactions
'   frmBankTransactions.Show
'   frmBankTransactions.ZOrder 0
End Sub

Private Sub cmdBankTransactions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Bank Transactions"
'   Me.MousePointer = vbArrow
End Sub

Private Sub cmdCashBook_Click()
   LoadForm frmCashbook
'   frmCashbook.Show
'   frmCashbook.ZOrder 0
End Sub

Private Sub cmdCashBook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Cashbook details"
'   Me.MousePointer = vbArrow
End Sub

Private Sub cmdClient_Click()
    LoadForm frmClientNew4
'    frmClientNew4.Show
'    frmClientNew4.ZOrder 0
End Sub

Private Sub cmdClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Client's detail information"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdClientBanks_Click()
   Dim adoRstS As New ADODB.Recordset
   Dim adoRstD As New ADODB.Recordset
   Dim iSs As Integer
   Dim Id As Integer

   Set Conn1 = New ADODB.Connection
   Conn1.Open getConnectionString
   
   adoRstS.Open "SELECT * FROM tlbClientBanks_", Conn1, adOpenStatic, adLockReadOnly
   adoRstD.Open "SELECT * FROM tlbClientBanks", Conn1, adOpenDynamic, adLockOptimistic
   
   While Not adoRstS.EOF
'Debug.Print adoRstS.Fields.count
'Debug.Print adoRstS.Fields.Item(iSs).Name
'Debug.Print adoRstS.Fields.Item(iSs).Type

      adoRstD.AddNew
      For iSs = 0 To adoRstS.Fields.Count - 1
         If adoRstS.Fields.Item(iSs).Name <> "MY_ID" Then
            For Id = 0 To adoRstD.Fields.Count - 1
               If adoRstD.Fields.Item(Id).Name = adoRstS.Fields.Item(iSs).Name Then
                  adoRstD.Fields.Item(Id).Value = adoRstS.Fields.Item(iSs).Value
                  Exit For
               End If
            Next Id
         End If
      Next iSs
      adoRstS.MoveNext
      adoRstD.Update
'Debug.Print "0"
   Wend
   
   adoRstS.Close
   adoRstD.Close
   Set adoRstS = Nothing
   Set adoRstD = Nothing
   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub cmdDemands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Demands & Receipts: Create Demands & Receipts"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Exit from the program"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdFeesCharges_Click()
'   LoadForm frmFeesCharges
'   frmFeesCharges.Show
'   frmFeesCharges.ZOrder 0
End Sub

Private Sub cmdFeesCharges_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Client fees and charges"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdGlobal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Global Data of the system"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdLease_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Lease details"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdNominalLedger_Click()
'   LoadForm frmNominalLedger1
'   frmNominalLedger1.Show
'   frmNominalLedger1.ZOrder 0
End Sub

Private Sub cmdPurEx_Click()
Rem by anol 20190408 issue 749 Locking records
'   If IsLoadedAndVisible("frmBPPreForm") Or _
'      IsLoadedAndVisible("frmBatchPayment") Then
'      MsgBox "Please close the batch payment before open the payment module.", vbInformation + vbOKOnly, "Purchase and Expenses"
'      Exit Sub
'   End If
   'issue 496 Purchase and expenses
   'Added by anol 13 Nov 2014
'         Dim strTemp As String
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If
'   If IsLoadedAndVisible("frmPurchaseExpense") Then
''        Cascading the forms issue 749 by anol 20190418
'        frmPurchaseExpense.Top = IsLoadedAndVisibleCount * 100
'        frmPurchaseExpense.Left = IsLoadedAndVisibleCount * 100
'        frmPurchaseExpense.ZOrder 0
'        Exit Sub
'   End If
    LoadForm frmPurchaseExpense
''   Load frmPurchaseExpense
''   frmPurchaseExpense.Show
''   frmPurchaseExpense.ZOrder 0
End Sub

Private Sub cmdPurEx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Purchases, Payments & Expenses"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdShopCentre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Property details"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdSupplier_Click()
   LoadForm frmSupplier
'   frmSupplier.Show
'   frmSupplier.ZOrder 0
End Sub

Private Sub cmdSupplier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Supplier Details"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdtenants_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Lessee's detail information"
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdTest_Click()
   LoadForm frmDemandTypes
'   frmDemandTypes.Show
'   frmDemandTypes.ZOrder 0
End Sub

Private Sub cmdUnits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Units details"
   Me.MousePointer = vbArrow
End Sub

Private Function FindClientBankErro() As Boolean
   'if there are more than 1 default bank details for a client then mark them not default
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   adoconn.Open getConnectionString
   szSQL = "SELECT Client.ClientID " & _
           "FROM tlbClientBanks INNER JOIN Client ON tlbClientBanks.CLIENT_ID = Client.ClientID " & _
           "GROUP BY Client.ClientID, tlbClientBanks.DEFAULT_AC " & _
           "HAVING (((tlbClientBanks.DEFAULT_AC)=True) AND ((Count(Client.ClientID))>1));"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      FindClientBankErro = True
   Else
      FindClientBankErro = False
      adoRst.Close
      adoconn.Close
      Set adoRst = Nothing
      Set adoconn = Nothing

      Exit Function
   End If

   frmClientNew4.LOAD_CLINT_CLIENTID = adoRst.Fields.Item("ClientID").Value

   While Not adoRst.EOF
      szSQL = "UPDATE tlbClientBanks SET tlbClientBanks.DEFAULT_AC = False " & _
              "WHERE tlbClientBanks.CLIENT_ID = '" & adoRst!ClientID & "';"
      adoconn.Execute szSQL
      adoRst.MoveNext
   Wend

   adoRst.Close
   adoconn.Close
   Set adoRst = Nothing
   Set adoconn = Nothing
End Function

Private Sub StatusMsgBarLoad()
   With rtxtMessageDisplay
      Call SetParent(.hWnd, stbStatusBar.hWnd)
      .Move stbStatusBar.Panels.Item(1).Width + 50, 50, stbStatusBar.Panels.Item(2).Width, .Container.Height - 50
      .SelStart = 0
      .SelLength = Len(.text)
      .Visible = True
   End With
End Sub



Private Sub Command1_Click()
    frmDemands3.Show
End Sub

Private Sub Command2_Click()
    frmRentPayable.Show
End Sub

Private Sub fraCmdButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Label1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
   Me.MousePointer = vbArrow
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub UpdateRecords(Conn1 As ADODB.Connection)
   Dim RstOne   As New ADODB.Recordset
   Dim Rst2    As New ADODB.Recordset
   Dim szSQL   As String
'DoEvents
' following 4/5 SQL are rem due to performance issue 2022-09-12
'   szSQL = "SELECT * " & _
'           "FROM   tlbReceipt AS R, DemandRecords AS D " & _
'           "WHERE  R.DemandRef = D.DemandID AND " & _
'                  "D.TransactionType = 2 AND " & _
'                  "R.SlNumber <> D.DmdSlNo;"
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
''Debug.Print "T: " & Rst2.RecordCount & " " & Rst2.Fields.Item("SlNumber").Value & "->" & Rst2.Fields.Item("DmdSlNo").Value
'      Conn1.Execute "UPDATE tlbReceipt AS R, DemandRecords AS D " & _
'                    "SET R.SlNumber = D.DmdSlNo " & _
'                    "WHERE R.DemandRef = D.DemandID AND " & _
'                          "D.TransactionType = 2;"
'   End If
'   Rst2.Close
'
'   szSQL = "SELECT P.MY_ID, P.PropertyID, S.TRANS " & _
'           "FROM tblPurInv AS P, tblPurInvSRec AS S " & _
'           "WHERE P.MY_ID = S.ParentID And P.PropertyID <> S.TRANS;"
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   While Not Rst2.EOF
'      Conn1.Execute "UPDATE tblPurInvSRec " & _
'                    "SET   TRANS = '" & Rst2.Fields.Item("PropertyID").Value & "' " & _
'                    "WHERE ParentID = '" & Rst2.Fields.Item("MY_ID").Value & "';"
'      Rst2.MoveNext
'   Wend
'   Rst2.Close

'   szSQL = "SELECT * FROM tblPurInv " & _
'           "WHERE  CL_ID = '' OR ISNULL(CL_ID);"
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
'      Conn1.Execute "UPDATE tblPurInv AS I, Property AS P " & _
'                    "SET CL_ID = P.ClientID " & _
'                    "WHERE  I.PropertyID = P.PropertyID;"
'   End If
'   Rst2.Close
'   'correcting incorrect posting for MFee payment 2022-05-20
'   szSQL = "SELECT * FROM NominalLedger " & _
'           "WHERE  CAName = 'Management Fees Control Account (B/S)';"
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
'      Conn1.Execute "UPDATE NominalLedger AS N " & _
'                    "SET CAName = 'Managing Agents control Account (B/S)' " & _
'                    "WHERE  CAName ='Management Fees Control Account (B/S)' ;"
'   End If
'   Rst2.Close
   
   
'   Dim rsPLC As New ADODB.Recordset
'   Dim rsPITransactions As New ADODB.Recordset
'   Debug.Print "New proc s: " & time
'   'Issue 841 2020-5-4 Duplicate entry found in the  NLPosting Table
'   rsPLC.Open "SELECT Code,ClientID FROM NominalLedger WHERE CAName = 'Purchase Ledger Control'", Conn1, adOpenKeyset, adLockReadOnly
'        'after tblPurINv.NLPOST=False I should delete the tlbpayment entries and create the again
'        'The FIND THE PI numbers that are mismatched ,then repost them, Type 6,7
'
'        While Not rsPLC.EOF
'            rsPITransactions.Open "SELECT X.TRANS_ID,X.AMOUNT1,Y.AMOUNT2,X.TRANSACTION_TYPE,X.NOMINAL_CODE,Y.TYPE,Y.NominalCOde,Y.PI FROM " & _
'            "(Select TRANS_ID,Sum(Amount) as AMOUNT1,TRANSACTION_TYPE,NOMINAL_CODE FROM NLPOSTING  where " & _
'            "DELETEFLAG=FALSE AND NOMINAL_CODE='" & rsPLC("Code").Value & "' AND ClientID='" & rsPLC("ClientID").Value & "' AND " & _
'            "(TRANSACTION_TYPE=6 OR TRANSACTION_TYPE=7) GROUP BY TRANS_ID,TRANSACTION_TYPE,NOMINAL_CODE) AS X " & _
'            "INNER Join " & _
'            "(SELECT IIF((R.Type=6) ,(-R.AMOUNT),(R.AMOUNT)) as AMOUNT2 ,R.Type, R.SageAccountNumber, " & _
'            "R.SlNumber,R.NominalCOde,R.PI FROM tlbPayment R WHERE ClientID='" & rsPLC("ClientID").Value & "' AND " & _
'            "(R.Type=6 OR R.Type=7)) AS Y " & _
'            "ON X.TRANS_ID=cstr(Y.SlNumber) AND X.TRANSACTION_TYPE=Y.Type where abs(Y.AMOUNT2)<> abs(X.AMOUNT1)", Conn1, adOpenKeyset, adLockReadOnly
'            While Not rsPITransactions.EOF
'                    Debug.Print rsPITransactions("TRANS_ID").Value & "-" & rsPLC("ClientID").Value
'                    Conn1.Execute "Update tblPurINV SET  tblPurINv.NLPOST=False where tblPurINV.MY_ID='" & rsPITransactions("PI").Value & "'"
'                    Conn1.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & rsPITransactions("TRANS_ID").Value & _
'                                    "' AND TRANSACTION_TYPE=" & rsPITransactions("TYPE").Value & " "
'                rsPITransactions.MoveNext
'            Wend
'            rsPITransactions.Close
'            rsPLC.MoveNext
'        Wend
'     rsPLC.Close
'     Set rsPLC = Nothing
'        Conn1.Close
'        Conn1.Open getConnectionString
'        Export_PInPC_2_NL Conn1
'        Debug.Print "New proc e: " & time
'   szSQL = " SELECT NLPosting.PARENT_RECORD, NLPosting.TRANS_ID, NLPosting.POSTED_DATE, NLPosting.TRANSACTION_TYPE, NLPosting.ClientID, NLPosting.ACCOUNT_NUMBER, NLPosting.AMOUNT, NLPosting.NOMINAL_CODE, NLPosting.TRANSACTION_REF, NLPosting.DeleteFlag," & _
'            " NLPosting.THIS_RECORD, NLPosting.REFERENCE, NLPosting.AMOUNT_TYPE,NLPosting.Transaction_ref FROM NLPosting WHERE (((NLPosting.PARENT_RECORD) In (SELECT [PARENT_RECORD] FROM" & _
'            " [NLPosting] As Tmp GROUP BY [PARENT_RECORD],[TRANS_ID],[POSTED_DATE],[TRANSACTION_TYPE],[ClientID],[ACCOUNT_NUMBER],[AMOUNT],[NOMINAL_CODE],[TRANSACTION_REF],[DeleteFlag] HAVING Count(*)>1  And [TRANS_ID]" & _
'            " = [NLPosting].[TRANS_ID] And [POSTED_DATE] = [NLPosting].[POSTED_DATE] And [TRANSACTION_TYPE] = [NLPosting].[TRANSACTION_TYPE] And" & _
'            " [ClientID] = [NLPosting].[ClientID] And [ACCOUNT_NUMBER] = [NLPosting].[ACCOUNT_NUMBER] And [AMOUNT] = [NLPosting].[AMOUNT]" & _
'            " And [NOMINAL_CODE] = [NLPosting].[NOMINAL_CODE] And [TRANSACTION_REF] = [NLPosting].[TRANSACTION_REF] And [DeleteFlag] =" & _
'            " [NLPosting].[DeleteFlag])))  AND NLPosting.DeleteFlag=false ORDER BY NLPosting.PARENT_RECORD, NLPosting.TRANS_ID, NLPosting.POSTED_DATE, NLPosting.TRANSACTION_TYPE, NLPosting.ClientID," & _
'            " NLPosting.ACCOUNT_NUMBER, NLPosting.AMOUNT, NLPosting.NOMINAL_CODE, NLPosting.TRANSACTION_REF, NLPosting.DeleteFlag;"
'  Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'  Dim X() As String, i As Long: ReDim X(Rst2.RecordCount)
'  i = 0
'  While Not Rst2.EOF
'        X(i) = Rst2("Transaction_ref").Value & "-" & Rst2("TRANSACTION_TYPE").Value
'        i = i + 1
'        Rst2.MoveNext
'  Wend
'  removeDuplicates X
'  Dim temp
'  For i = 0 To UBound(X) - 1
'        'Debug.Print i & ":" & x(i)
'          temp = Split(X(i), "-")
'          adoConn.Execute "Update tblPurINV SET  tblPurINv.NLPOST=False where tblPurINV.SlNumber=" & temp(0) & " AND tblPurINV.TransactionType=" & temp(1) & " "
'          adoConn.Execute "Update NLPOSTING SET  DeleteFlag=true where TRANSACTION_REF='" & temp(0) & _
'                                    "' AND TRANSACTION_TYPE='" & temp(1) & "'"
'         Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'Reposting PI: PI  Serial - type: " & temp(0) & "-" & temp(1) & "' )"
'  Next
'  Rst2.Close
'  Export_PInPC_2_NL adoConn
  'UpdateDatabase3 Conn1 'added by anol 20170316 tlbReceipt.ClientID was not found here
'   added by anol 28 July 2016 FOR CAHSBOOK REPORT WAS NOT SHOWING CORRECTLY
'    Conn1.Execute "Update tlbReceipt,Units,Property SET tlbReceipt.ClientID=Property.ClientID where tlbReceipt.UnitID=Units.UnitNumber AND " & _
'    "Units.PropertyID=Property.PropertyID and tlbReceipt.ClientID is null"
    'FOR SUPPLIEr OS AMOUNT INCORRECT lATER ON I WRItE PRopErTY id ON TH PAyMENT SPLiT aND NOW NOW NEED ThIS SQL
'    Conn1.Execute "Update tlbpayment S,tlbpayment P,PayTransactions PT SET S.UNITID=P.UNITID where PT.FromTran=S.TransactionID and " & _
'            "PT.ToTran=P.TransactionID and (S.UNITID='' OR isnull(S.UNITID))"
''    'added by anol 25 Feb 2017 FOR aged creditors report was not showing correctly issue 308
''    Conn1.Execute "Update tlbPayment,Property SET tlbPayment.ClientID=Property.ClientID where tlbPayment.UnitID=Property.PropertyID AND " & _
''    " (tlbPayment.ClientID is null or tlbPayment.ClientID='')"
''
    'added by anol 08 Feb 2017 issue 304
    'There is an inconsistency between the nominal ledger bank balance and the cash book balance.Solution:
    'Write a routine that runs on “Loading Program” that performs the following task:
    Rem again ny anol 20170918 because I think we disccussed the program should show 0 Valued transaction
'    Conn1.Execute "Update NLPosting N,tlbReceipt R set deleteflag=true where R.SageAccountNumber=N.ACCOUNT_NUMBER AND N.TRANSACTION_REF=cstr(R.SlNumber)" & _
'        " AND R.Amount=0 AND N.REFERENCE=R.Extref AND R.Type=3"
'    Debug.Print "1" & time
    'this line has been remmed for optimization
'    Conn1.Execute "Update NLPosting N,tlbPayment P set deleteflag=true where P.SageAccountNumber=N.ACCOUNT_NUMBER AND N.TRANSACTION_REF=cstr(P.SlNumber) " & _
'    "AND P.Amount=0 AND N.REFERENCE=P.Extref AND P.Type=8"
     'added by anol 07 Feb 2017 issue 253
'     Debug.Print "1" & time
'    Conn1.Execute "Update DemandTypes D,Property P,NominalLedger N  set D.NominalNameforAmount=N.Name where D.PropertyID=P.PropertyID AND " & _
'    "P.ClientID=N.ClientID AND D.NominalCodeforAmount=N.Code AND NominalNameforAmount is null"
   'added by anol 0n 07/ 04/ 2016
   ' fixed by anol 20170320 anol issue 327 procedure was taking long time to run
'   Conn1.Execute "update NLPOSTING set Reference='BACS' where Reference='' AND cstr(PARENT_RECORD) in ( Select cstr(transactionID) from tlbPayment where PayAmtType='BACS' and EXTref='')"
'this line has been remmed for optimization
''   Conn1.Execute "update NLPOSTING A, tlbPayment B set A.Reference='BACS' where A.Reference='' AND cstr(A.PARENT_RECORD) = cstr(B.transactionID) AND B.PayAmtType='BACS' and B.EXTref='';"
''   Conn1.Execute "update tlbPayment  set extref='BACS' where extref='' and PayAmtType='BACS'"
'   Debug.Print "2" & time
    'added by anol 0n 21/ 04/ 2016 NLPOSTING nominal code was not updating due to programming bug
   szSQL = "SELECT * from NLPOSTING where NOMINAL_CODE is null or  NOMINAL_CODE=''"
   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   If Not Rst2.EOF Then
      MsgBox "This database is not up to date. Please contact PCM support quoting the reference: NLPC_BLANK", vbInformation, "Prestige warning"
   End If
   Rst2.Close
   szSQL = "SELECT q.DmdSlNo from (Select DemandID,DmdSlNo from DemandRecords) as q  INNER JOIN" & _
    "(SELECT DemandID,SplitID from DemandSplitRecords GROUP BY  DemandID,SplitID " & _
    "HAVING (  COUNT(splitID)  > 1 ))  as Q1  ON  Q1.DemandID=q.DemandID   order by  q.DmdSlNo Desc"
' szSQL = "SELECT q.DmdSlNo from ((Select DemandID,DmdSlNo from DemandRecords) as q  INNER JOIN" & _
'    "(SELECT DemandID,SplitID from DemandSplitRecords GROUP BY  DemandID,SplitID " & _
'    "HAVING (  COUNT(splitID)  > 1 ))  as Q1  ON  Q1.DemandID=q.DemandID order by  q.DmdSlNo Desc) AS N INNER JOIN tlbReceipt as M ON M.SlNumber=N.DemandID"
   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst2.EOF Then
   'Due to program error Dupplice splitID was generating at demand split table. That was creating from edit demand.
   'That area has been fixed. as well as this procedure is here if in the old  database we still have this problem. FIxed by anol Jun 2106
      MsgBox "This database is not up to date. Please contact PCM support quoting the reference: DMND_SPLIT_DUP", vbInformation, "Prestige warning. DemandID : " & SQL2String(Rst2, 0)
   End If
   Rst2.Close
   Set Rst2 = Nothing
   ' Checking vat amount with total amount
'   szSQL = "SELECT DemandID,TotalAmount,VATAMOUNT,Amount,(TotalAmount-VATAMOUNT-Amount) as TT FROM DemandSplitRecords WHERE 0<> Round((TotalAmount-VATAMOUNT-Amount) ,2)"
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
'      MsgBox "This database is not up to date. Please contact PCM support quoting the reference: Identify_VAT_inconsistency", vbInformation, "Prestige warning. DemandID : " & SQL2String(Rst2, 0)
'   End If
'   Rst2.Close
'   Set Rst2 = Nothing
'issue 133 if the lease end date < current date and override lease end date(OLED) is FALSE, the program should change the "Status" of the lease to
' expired (FALSE) and copy the lease end date to the terminate date column
    ' I am moving this procedure befreo the tree loads issue 405
   ' Conn1.Execute "Update leasedetails set status=false  where OLED=false and EndDate< date() and status=true"
    'Batch Demand Had produced calculation error on it.
    'Conn1.Execute "Update DemandSplitRecords SET TotalAmount=Amount+VATAmount Where A_M='B'"
    
    
    
    'Freqency ID is not writing for some Demand split records
   'Fixed by anol on 03 Jun 2016
   'issue 534
   '(DT.CategoryCode=2 OR DT.CategoryCode=4) AND
   Rem on 2019-06-17
''   Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LServiceCharges LS SET DS.FrequencyID=LS.SCFrequency,DS.ChargingMethod=LS.ChargingMethod,DR.SO=true " & _
''   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=2 AND L.LeaseID=DR.LeaseRef AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
''   "AND isnull(LS.Spare3)"
''   'Here LS.Spare3 is Delete Flag of LService charge
''   Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LRentCharges LS SET DS.FrequencyID=LS.BRFrequency,DS.ChargingMethod=LS.spare1, DR.SO=true " & _
''   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=1 AND L.LeaseID=DR.LeaseRef AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
''   "AND isnull(LS.Spare3)"
''    Conn1.Execute "UPDATE DemandSplitRecords DS,DemandRecords DR,LeaseDetails L,DemandTypes DT,LInsuranceCharges LS SET DS.FrequencyID=LS.InsuranceFrequency,DS.ChargingMethod=LS.ChargingType, DR.SO=true " & _
''   "where LS.LeaseID=L.LeaseID AND DS.TypeOfDemand=DT.ID AND DT.CategoryCode=3 AND L.LeaseID=DR.LeaseRef AND DS.DemandID=DR.DemandID AND A_M='A' and FrequencyID=0 " & _
''   "AND isnull(LS.Spare3)"
   'Issue 440 Invoice number ,receipt and payment number was showing incorrectly in NLhistory
   'added by anol 20170807
   'IN NLPosting table trasation ref was empty so they needs to be updated from the receipt ,payment and bank transacation table.
   'By this order as well
   ' NOw this has been transferred to the help menu
'   Type 3,4,23
''    Conn1.Execute "UPDATE NLPOSTING, tlbReceipt SET TRANSACTION_REF = slNumber " & _
''    "WHERE NLPOSTING.TRANS_ID=cstr(tlbReceipt.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbReceipt.Type AND TRANSACTION_REF is NULL;"
''
'''    Type 24,8,9
''    Conn1.Execute "UPDATE NLPOSTING, tlbPayment SET TRANSACTION_REF = tlbPayment.slNumber ,NLPOSTING.TRANS_ID=tlbPayment.slNumber " & _
''    "WHERE NLPOSTING.PARENT_RECORD=cstr(tlbPayment.TransactionID) AND NLPOSTING.TRansaction_TYpe=tlbPayment.Type AND " & _
''    "NLPOSTING.TRANSACTION_REF is NULL AND  NLPOSTING.TRANS_ID IS NULL;"
''   'Type 6,7,11,12
''    Conn1.Execute "UPDATE NLPOSTING SET TRANSACTION_REF = TRANS_ID where NLPOSTING.TRansaction_TYpe in (11,12,24,8,9,6,7) AND TRANSACTION_REF is NULL"
''
''    'Type 1,2
''    Conn1.Execute "UPDATE NLPOSTING, DemandRecords SET TRANSACTION_REF = DmdSlNo " & _
''    "WHERE NLPOSTING.TRANS_ID=cstr(DemandRecords.DemandID) AND NLPOSTING.TRansaction_TYpe=DemandRecords.TransactionType " & _
''    "AND TRANSACTION_REF is NULL; "
'issue 440 added by anol 20170811
''For Supplier Balance
''IN the tlbPayment ClientID cannot be null, this will cause
' a problem while building a client Balance
''Need to update clientID from from PI where clientID is null


'      Conn1.Execute "Update tlbPayment INNER JOIN tblPurInv ON tlbPayment.PI = tblPurInv.MY_ID set " & _
'                    "tlbPayment.ClientID=tblPurInv.CL_ID WHERE (((tlbPayment.ClientID) Is Null));"
'                   ' issue 655. solved on 2018-10-19 could not find the reason why that msg was showing
'     ' Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN ClientID Text(10) NOT NULL;"
'         Debug.Print "3" & time
      
      '20181012 issue 644
      'In case there as any Client.agent,landlord entity are not in supplier table then insert it into supplier table
      'There is a contra in in landlord if you are saving a landlord it is saving into the supplier table as well as landlord table with an additional prefix 'L-'
      'So you need to ignore those
      'from 2020-01-07 I have decided that no need to update landlord to supplier and agent to supplier because this shall create circular reference and landlord and supplier table
      szSQL = "SELECT ClientID,ClientName,ClientAddressLine1,ClientAddressLine2,ClientAddressLine3,ClientAddressLine4,ClientPostCode,ClientOfficeAddressLine1,ClientOfficeAddressLine2, " & _
      "ClientOfficeAddressLine3,ClientOfficeAddressLine4,ClientOfficePostCode,ClientOfficeEmail,ClientPersonalEmail,'CLIENT' From Client C  LEFT JOIN Supplier S ON  C.ClientID=S.SupplierID  " & _
      "where isnull(S.SupplierID)"
      Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
      If Not Rst2.EOF Then
             Conn1.Execute "Insert into Supplier(SupplierID,SupplierName,SupplierAddressLine1,SupplierAddressLine2,SupplierAddressLine3,SupplierAddressLine4,SupplierPostCode," & _
             "SupplierOfficeAddressLine1,SupplierOfficeAddressLine2,SupplierOfficeAddressLine3,SupplierOfficeAddressLine4,SupplierOfficePostCode,SupplierOfficeEmail," & _
             "SupplierPersonalEmail,TYPE)  " & szSQL
             
      End If
      Rst2.Close
      Set Rst2 = Nothing
      'added by anol 2023-05-15 . Tlbpayment was holding some deleted records. which is not needed to keep.
      Conn1.Execute "Delete from tlbPaymentSplit where description='DELETED PURCHASE TRANSACTIONS'"
      
''      szSQL = "SELECT AgentID,AgentName,AgentAddressLine1,AgentAddressLine2,AgentAddressLine3,AgentAddressLine4,AgentPostCode, " & _
''      "AgentOfficeAddressLine1,AgentOfficeAddressLine2,AgentOfficeAddressLine3,AgentOfficeAddressLine4,AgentOfficePos,AgentOfficeEmail,AgentPersonalEmail,'AGENT' From AGENT C  LEFT JOIN Supplier S ON  C.AgentID=S.SupplierID  " & _
''      "where isnull(S.SupplierID)"
''      Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''      If Not Rst2.EOF Then
''             Conn1.Execute "Insert into Supplier(SupplierID,SupplierName,SupplierAddressLine1,SupplierAddressLine2,SupplierAddressLine3,SupplierAddressLine4,SupplierPostCode," & _
''             "SupplierOfficeAddressLine1,SupplierOfficeAddressLine2,SupplierOfficeAddressLine3,SupplierOfficeAddressLine4,SupplierOfficePostCode,SupplierOfficeEmail," & _
''             "SupplierPersonalEmail,TYPE)  " & szSQL
''
''      End If
''      Rst2.Close
''      Set Rst2 = Nothing
''
''       szSQL = "SELECT LandlordID,LandlordName,LandlordAddressLine1,LandlordAddressLine2,LandlordAddressLine3,LandlordAddressLine4,LandlordPostCode, " & _
''      "LandlordOfficeAddressLine1,LandlordOfficeAddressLine2,LandlordOfficeAddressLine3,LandlordOfficeAddressLine4,LandlordOfficePostCode,LandlordOfficeEmail,LandlordPersonalEmail,'LLORD' From Landlord C  LEFT JOIN Supplier S ON  C.LandlordID=S.SupplierID  " & _
''      "where isnull(S.SupplierID)"
''      Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''      If Not Rst2.EOF Then
''             Conn1.Execute "Insert into Supplier(SupplierID,SupplierName,SupplierAddressLine1,SupplierAddressLine2,SupplierAddressLine3,SupplierAddressLine4,SupplierPostCode," & _
''             "SupplierOfficeAddressLine1,SupplierOfficeAddressLine2,SupplierOfficeAddressLine3,SupplierOfficeAddressLine4,SupplierOfficePostCode,SupplierOfficeEmail," & _
''             "SupplierPersonalEmail,TYPE)  " & szSQL
''
''      End If
''      Rst2.Close
''      Set Rst2 = Nothing
      'issue 645 date 20181015
     ' Conn1.Execute "Update NominalLedger Set Posting =-1 where  CATYPE in('s','r','p','o','I')"
     'issue 659 Vat problem in PI
''      Conn1.Execute "UPDATE Supplier S INNER JOIN tlbVatcode V ON S.vatcode = V.vat_code SET S.vatcode = V.vat_ID"
''      Conn1.Execute "UPDATE tlbPayment P Set P.OSAmount=0 where Type=8 and  P.OSAmount IS NULL"
'        Debug.Print "4" & time
      Conn1.Execute "UPDATE tlbPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' where ServerIPaddress='" & GetIPaddress & "'"
      Conn1.Execute "UPDATE tlbReceipt P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' where ServerIPaddress='" & GetIPaddress & "'"
      Conn1.Execute "UPDATE tlbBankPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' where ServerIPaddress='" & GetIPaddress & "'"
      Conn1.Execute "UPDATE NJ_Header P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' where ServerIPaddress='" & GetIPaddress & "'"
'       Debug.Print "5" & time
      
'      Conn1.Execute "UPDATE tlbPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress=''"
'      Conn1.Execute "UPDATE tlbReceipt P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
'      Conn1.Execute "UPDATE tlbBankPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
'      Conn1.Execute "UPDATE NJ_Header P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
      If DateDiff("d", Date, "27 Nov 2020") >= 0 Then
            Conn1.Execute "Update tlbClientBanks A INNER JOIN Client B ON A.Client_ID=B.ClientID SET FileLoc='U:\BACSFILES\Austin Chambers\PTX\UNPROCESSED BACS FILES' Where B.GroupCode='AC'"
            Conn1.Execute "Update tlbClientBanks A INNER JOIN Client B ON A.Client_ID=B.ClientID SET FileLoc='U:\BACSFILES\Savoy Stewart\PTX\UNPROCESSED BACS FILES' Where B.GroupCode='SS'"
            
            Conn1.Execute "Update tlbClientBanks A INNER JOIN Client B ON A.Client_ID=B.ClientID SET ProcessFileLoc='U:\BACSFILES\Austin Chambers\PTX\INPUT' Where B.GroupCode='AC' AND (ProcessFileLoc IS NULL or ProcessFileLoc='')"
            Conn1.Execute "Update tlbClientBanks A INNER JOIN Client B ON A.Client_ID=B.ClientID SET ProcessFileLoc='U:\BACSFILES\Savoy Stewart\PTX\INPUT' Where B.GroupCode='SS'  AND (ProcessFileLoc IS NULL or ProcessFileLoc='')"
      End If
      'added by anol 2020-12-02 issue 898
      szSQL = "SELECT PrimaryCode,Code from SecondaryCode where  PrimaryCode='DCTG' and Code='5'"
      Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
      If Rst2.EOF Then
             Conn1.Execute "Insert into SecondaryCode values('DCTG','5','Service Charge Year End','Service Charge Year End')"
      End If
      Rst2.Close
      Set Rst2 = Nothing
      'Service Charge Year End
      Dim szFundID As Integer
      'add new fund code 'TENANTDEPOSIT' rentdeposit Here 2021-01-08
      szSQL = "Select FundID as FID from Fund"
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
      If Not Rst1.EOF Then
            szSQL = "Select (max(FundID)+1) as FID from Fund"
            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
            szFundID = 1 'Default value which shall end into a error
            If Not Rst2.EOF Then
                  szFundID = Rst2("FID").Value
            End If
            Rst2.Close
            Set Rst2 = Nothing
            szSQL = "SELECT FundCode from FUND where  FundCode='TENANTDEPOSIT'"
            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
            If Rst2.EOF Then
                  szSQL = "INSERT INTO FUND (FUNDID,FundCode, FundName, CategoryCode) " & _
                          "VALUES (" & szFundID & ",'TENANTDEPOSIT', 'TENANT Deposit', 4);"
                  Conn1.Execute szSQL
                  Conn1.Execute "UPDATE Fund SET szFundID = CSTR(FundID);"
            End If
            Rst2.Close
            Set Rst2 = Nothing
      End If
      Rst1.Close
      If DateDiff("d", Date, "01 May 2021") >= 0 Then
            Conn1.Execute "Update NominalLedger set CAType='' where CAType is null"
      End If
      'This fix was written because when you end a lease status to false automatically by startup update query you need to vacant the unit status occupied to N
      Conn1.Execute "Update Units UT LEFT JOIN (Select L.UnitNumber from  LeaseDetails L,  Units U where L.UnitNumber=U.UnitNumber " & _
                "AND L.Status = True) AS A ON A.UnitNumber=UT.UnitNumber SET UT.Occupied='N' where UT.Occupied='Y' AND isnull(A.UnitNumber)"
'      RstOne.Open "Select * from Units where Occupied='Y'", Conn1, adOpenStatic, adLockReadOnly
'      While Not RstOne.EOF
'            szSQL = "SELECT LeaseDetails.* " & _
'                      "FROM LeaseDetails, UnitS " & _
'                      "WHERE LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
'                        "LeaseDetails.Status = True And " & _
'                        "Units.UnitNumber = '" & szUnitNumber & "';"
'            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'            RstOne.MoveNext
'      Wend
      
   End Sub
Private Sub removeDuplicates(ByRef arrName() As String)
    Dim i As Long, tempArr() As String: ReDim tempArr(UBound(arrName))
    Dim d As New Dictionary, n As Long
    For i = 0 To UBound(arrName)
        If Not d.Exists(arrName(i)) Then
            d.Add arrName(i), arrName(i)
            tempArr(n) = arrName(i): n = n + 1
        End If
    Next
    ReDim Preserve tempArr(n)
    arrName = tempArr
End Sub
Private Sub MDIForm_DblClick()
   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
      '# Repaint Ribbon
'      ACPRibbon1.Refresh
   End If
End Sub

Private Sub ACPRibbon1_CatClick(ByVal Id As String, ByVal Caption As String)
'   MsgBox ID & " " & Caption
End Sub

Private Sub UpdateRibbonData(adoRibConn As ADODB.Connection)
'Written by anol 10 08 2016
    Dim rsAdd As New ADODB.Recordset
    rsAdd.Open "Select * from Items where ID='I117'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "117"
        rsAdd!Id = "I117"
        rsAdd!G_ID = "G31"
        rsAdd!ItemNameL1 = "Report"
        rsAdd!iconnumber = "95"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "3"
        rsAdd.Update
    End If
    rsAdd.Close
    rsAdd.Open "Select * from TreeRoot where RootKey='CB2'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!rootkey = "CB2"
        rsAdd!RootName = "Cashbook"
        rsAdd!Rootvisible = True
        rsAdd!formicon = "7"
        rsAdd!formicon = "7"
        rsAdd.Update
    End If
    rsAdd.Close
    rsAdd.Open "Select * from TreeChildren where ChildKey='CBH'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!childkey = "CBH"
        rsAdd!Parentkey = "CB2"
        rsAdd!childName = "Cashbook History"
        rsAdd!childvisible = True
        rsAdd!iconno = "7"
        rsAdd.Update
    End If
    rsAdd.Close
    'Updating a report tree// adding a new child menu for print labels
     rsAdd.Open "Select * from TreeChildren where ChildKey='LPL'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!childkey = "LPL"
        rsAdd!Parentkey = "L"
        rsAdd!childName = "Print Labels"
        rsAdd!childvisible = True
        rsAdd!iconno = "7"
        rsAdd.Update
    End If
    rsAdd.Close
    'added 20180824 issue 605
    adoRibConn.Execute "Update Items set ItemNameL2='List' where ID='I71'"

'Written by anol 14 05 2019 issue 764
    'adoRibConn.Execute "Update Groups set T_ID='T10' where GroupName='Help'"
    
    'Written by anol 15 05 2019 issue 764

    rsAdd.Open "Select * from Items where ID='I118'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "118"
        rsAdd!Id = "I118"
        rsAdd!G_ID = "G36"
        rsAdd!ItemNameL1 = "Rollback"
        rsAdd!ItemNameL2 = "Bank Reconciliation"
        rsAdd!iconnumber = "1"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "0"
        rsAdd.Update
    End If
    rsAdd.Close
    'added 2019-06-18
    rsAdd.Open "Select * from Items where ID='I119'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "119"
        rsAdd!Id = "I119"
        rsAdd!G_ID = "G36"
        rsAdd!ItemNameL1 = "Tile Windows"
        rsAdd!ItemNameL2 = ""
        rsAdd!iconnumber = "13"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "0"
        rsAdd.Update
    End If
    rsAdd.Close
    'added 2019-08-07 issue 710
    adoRibConn.Execute "Update Items set ItemNameL1='Maintenance',IconNumber=91,G_ID = 'G36' where ID='I99'"
    adoRibConn.Execute "Update TreeRoot set RootName='Property and Unit Reports' where rootkey='U'"
    adoRibConn.Execute "Update Items set DisplayOrder='13' where ID='I117'"
    
    ' added by anol 2020-01-7 issue 820
'    rsAdd.Open "Select * from TreeRoot where rootkey='P'", adoRibConn, adOpenKeyset, adLockOptimistic
'    If rsAdd.EOF Then
'        rsAdd.AddNew
'        rsAdd!rootkey = "P"
'        rsAdd!RootName = "Property"
'        rsAdd!Rootvisible = True
'        rsAdd!Formcaption = ""
'        rsAdd!formicon = 7
'        rsAdd!parenticonNo = 7
'        rsAdd.Update
'    End If
'    rsAdd.Close
    
    rsAdd.Open "Select * from TreeChildren where childkey='PL'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!childkey = "PL"
        rsAdd!Parentkey = "U"
        rsAdd!childName = "Property List Report"
        rsAdd!childvisible = True
        rsAdd!iconno = 7
        rsAdd.Update
    End If
    rsAdd.Close
    'added 2020-08-08
    rsAdd.Open "Select * from Items where ID='I120'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "120"
        rsAdd!Id = "I120"
        rsAdd!G_ID = "G35"
        rsAdd!ItemNameL1 = "Statement"
        rsAdd!ItemNameL2 = "Template"
        rsAdd!iconnumber = "80"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "0"
        rsAdd.Update
    End If
    rsAdd.Close
    'added 2020-08-03
    rsAdd.Open "Select * from Items where ID='I121'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "121"
        rsAdd!Id = "I121"
        rsAdd!G_ID = "G32"
        rsAdd!ItemNameL1 = "Consolidated"
        rsAdd!ItemNameL2 = "Bank List"
        rsAdd!iconnumber = "114"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "0"
        rsAdd.Update
    End If
    rsAdd.Close
    
    
    '********************Addig a new report for CS
        rsAdd.Open "Select * from Items where ID='I122'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "122"
        rsAdd!Id = "I122"
        rsAdd!G_ID = "G14"
        rsAdd!ItemNameL1 = "Fees and Charges "
        rsAdd!ItemNameL2 = "Reports"
        rsAdd!iconnumber = "95"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "3"
        rsAdd.Update
    End If
    rsAdd.Close
    rsAdd.Open "Select * from TreeRoot where RootKey='MGR'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!rootkey = "MGR"
        rsAdd!RootName = "Fees and Charges Reports"
        rsAdd!Rootvisible = True
        rsAdd!formicon = "7"
        rsAdd!formicon = "7"
        rsAdd.Update
    End If
    rsAdd.Close
    rsAdd.Open "Select * from TreeChildren where ChildKey='MGF'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!childkey = "MGF"
        rsAdd!Parentkey = "MGR"
        rsAdd!childName = "Managing Agent's Fees Report"
        rsAdd!childvisible = True
        rsAdd!iconno = "7"
        rsAdd.Update
    End If
    rsAdd.Close
    
     '********************Adding a button for Retention
    rsAdd.Open "Select * from Items where ID='I123'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "123"
        rsAdd!Id = "I123"
        rsAdd!G_ID = "G14"
        rsAdd!ItemNameL1 = "Retentions"
        rsAdd!ItemNameL2 = ""
        rsAdd!iconnumber = "106"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "3"
        rsAdd.Update
    End If
    rsAdd.Close
    
     'added 2022-10-28
    rsAdd.Open "Select * from Items where ID='I124'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "124"
        rsAdd!Id = "I124"
        rsAdd!G_ID = "G35"
        rsAdd!ItemNameL1 = "Client Statement"
        rsAdd!ItemNameL2 = "Template"
        rsAdd!iconnumber = "86"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "0"
        rsAdd.Update
    End If
    rsAdd.Close
    'added 2023-02-20 adding audit log
    rsAdd.Open "Select * from Items where ID='I125'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "125"
        rsAdd!Id = "I125"
        rsAdd!G_ID = "G32"
        rsAdd!ItemNameL1 = "Audit"
        rsAdd!ItemNameL2 = "Log"
        rsAdd!iconnumber = "11"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "9"
        rsAdd.Update
    End If
    rsAdd.Close
    
    'added 2023-02-21 adding View Bank Balances
    rsAdd.Open "Select * from Items where ID='I126'", adoRibConn, adOpenKeyset, adLockOptimistic
    If rsAdd.EOF Then
        rsAdd.AddNew
        rsAdd!X = "126"
        rsAdd!Id = "I126"
        rsAdd!G_ID = "G31"
        rsAdd!ItemNameL1 = "View Bank"
        rsAdd!ItemNameL2 = "Balances"
        rsAdd!iconnumber = "10"
        rsAdd!Display = True
        rsAdd!DisplayOrder = "9"
        rsAdd.Update
    End If
    rsAdd.Close
    
End Sub

Function CheckRolePermission()
'''''''''''''''''''''Modified by Mahboob Change Id 13 Work Item 10 Update Procedure
Dim cn As ADODB.Connection
Dim psql As String
Dim iSQL As String
Dim rsRole As New ADODB.Recordset
Dim rsItem As New ADODB.Recordset
Set cn = New ADODB.Connection

cn.Open getConnectionStringUserAccess
cn.Execute "UPDATE Items SET Display =False"
psql = ""
psql = "Select ItemID from RolePermissions WHERE RoleID=" & roleID & ""
rsRole.Open psql, cn, adOpenDynamic, adLockOptimistic
    If Not rsRole.EOF Then
    Do While Not rsRole.EOF
'            For i = 1 To lvwavailableCompany.ListItems.Count
'                If lvwavailableCompany.ListItems(i).Tag = Val("" & rst.Fields("CompanyID")) Then lvwavailableCompany.ListItems(i).Checked = True
'            Next
iSQL = ""
iSQL = "Select ID from Items WHERE ID='" & rsRole!itemID & "'"
rsItem.Open iSQL, cn, adOpenDynamic, adLockOptimistic
    If Not rsItem.EOF Then
    cn.Execute "UPDATE Items SET Display =TRUE where ID='" & rsRole!itemID & "'"
    End If
    rsItem.Close
            rsRole.MoveNext
        Loop
    End If
'cn.Execute "UPDATE Items SET Display = IIf((SELECT RoleID FROM RolePermissions WHERE RolePermissions.RoleID=" & roleID & " and RolePermissions.ItemID = Items.ID) > 0, True, False)"
rsRole.Close
Set rsRole = Nothing
'rsItem.Close
Set rsItem = Nothing
    cn.Execute "UPDATE Items SET Display =False where ID  in ('" & "I3" & "','" & "I4" & "')"
    cn.Execute "UPDATE Groups SET GroupName ='" & "Company Details" & "' where ID ='" & "G2" & "'"
cn.Close
Set cn = Nothing
End Function


Private Sub LoadingRibbon() 'Loading menu
   Exit Sub
'   Dim adoRibConn As New ADODB.Connection
'   Dim adoRst     As New ADODB.Recordset
'   Dim szItem     As String
'   ''''''''Modified by Mahboob Change Id 13 Work Item 11 Update Procedure
'   CheckRolePermission
'   'End of modification
''   adoRibConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Common Files\RibbonLayOut.mdb;Persist Security Info=False;"
'   adoRibConn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\Common Files\RibbonLayOut.mdb;"
'   Call UpdateRibbonData(adoRibConn)
'   Theme = 1
'
'   '# SET Theme
'   ACPRibbon1.Theme = Theme    ' 0 - Black
'                               ' 1 - Blue
'                               ' 2 - Silver
'
'   '# OPTIONAL - Load Background for Form.
'   Me.Picture = ACPRibbon1.LoadBackground
'
'   '# OPTIONAL - Load Background for Form
''   Me.BackColor = ACPRibbon1.BackColor
'
'   '# Set ImageList to use for icons
'   ACPRibbon1.ImageList = ImageList1
'
'   '# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
'   ACPRibbon1.ButtonCenter = False
'
'   '# Add Tabs ---   ID - Caption         :     Group Header, tabs
'   adoRst.Open "SELECT * FROM Tabs WHERE Display;", adoRibConn, adOpenStatic, adLockReadOnly
'
'   While Not adoRst.EOF
''                             ID                   ,         Caption
'      ACPRibbon1.AddTab adoRst.Fields.Item(0).Value, adoRst.Fields.Item(1).Value
'      adoRst.MoveNext
'   Wend
'   adoRst.Close
'
'   adoRst.Open "SELECT * FROM Groups WHERE Display;", adoRibConn, adOpenStatic, adLockReadOnly
'
'   '# Add Cats ---        :     Group
'   While Not adoRst.EOF
''                             ID                   ,         Tab                ,         Caption            ,  CatLink
'      ACPRibbon1.AddCat adoRst.Fields.Item(0).Value, adoRst.Fields.Item(1).Value, adoRst.Fields.Item(2).Value, adoRst.Fields.Item(4).Value
'      adoRst.MoveNext
'   Wend
'   adoRst.Close
'
'   adoRst.Open "SELECT * FROM Items WHERE Display ORDER BY DisplayOrder;", adoRibConn, adOpenStatic, adLockReadOnly
'
''# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip      :     Items
'   While Not adoRst.EOF
'      If IsNull(adoRst.Fields.Item("ItemNameL2").Value) Then
'         szItem = adoRst.Fields.Item("ItemNameL1").Value
'      Else
'         szItem = adoRst.Fields.Item("ItemNameL1").Value & vbNewLine & adoRst.Fields.Item("ItemNameL2").Value
'      End If
''                                   ID                   ,         Cat                     , Caption,           Icon                              , More Arrow,          ToolTip Msg
'      ACPRibbon1.AddButton adoRst.Fields.Item("ID").Value, adoRst.Fields.Item("G_ID").Value, szItem, CInt(adoRst.Fields.Item("IconNumber").Value), False, IIf(IsNull(adoRst.Fields.Item("ToolTip").Value), "", adoRst.Fields.Item("ToolTip").Value)
'      adoRst.MoveNext
'   Wend
'   adoRst.Close
'
'   '# Repaint Ribbon
'   ACPRibbon1.Refresh
'
''   adoRst.Close
'   Set adoRst = Nothing
'
'   adoRibConn.Close
'   Set adoRibConn = Nothing
End Sub

Private Sub MDIForm_Activate()

   On Error GoTo Err
   If iFORM_LOAD = -1 Then End

   tvwLandLord.Top = 0
   tvwLandLord.Left = 0
   tvwLandLord.Height = picTreeView.Height
   tvwLandLord.Width = picTreeView.Width

   stbStatusBar.Panels(3).text = "User: " & SystemUserName
   stbStatusBar.Panels(4).text = Format(Now, "dddd   DD-MMMM-YYYY")     'dddd Day
   stbStatusBar.Panels(4).AutoSize = sbrContents
   stbStatusBar.Panels(6).AutoSize = sbrContents

   stbStatusBar.Panels(1).Width = Me.Width - _
                                  stbStatusBar.Panels(2).Width - stbStatusBar.Panels(3).Width - _
                                  stbStatusBar.Panels(4).Width - stbStatusBar.Panels(5).Width - _
                                  stbStatusBar.Panels(6).Width - stbStatusBar.Panels(7).Width
   bLogOFF = False

   If FindClientBankErro Then
      ShowMsgInTaskBar "Problems have been found in the client's Bank details. Please setup the default bank details for each clients.", , "N"
      cmdClient_Click
   End If
   Exit Sub
Err:
End Sub
Private Function UpdateSparetable5(Conn1 As ADODB.Connection) As Integer
   On Error GoTo Err
   Dim Rst1 As New ADODB.Recordset
   Rst1.Open "Select Code from Sparetable5", Conn1, adOpenStatic, adLockReadOnly
   If Rst1("Code").DefinedSize <> 255 Then
        Rst1.Close
        Conn1.Execute "ALTER Table SpareTable5 ALTER COLUMN Code Text(255);"
        UpdateSparetable5 = 1
   Else
        Rst1.Close
        GoTo ADD_ClientOfficeAddressLine4
   End If
   Exit Function
Err:
   Conn1.Execute "ALTER Table SpareTable5 add COLUMN Code Text(255);"
   Conn1.Execute "ALTER Table SpareTable5 add COLUMN clientID Text(10);"
   Conn1.Execute "ALTER Table SpareTable5 add COLUMN CC Text(255);"
   UpdateSparetable5 = 1
   Exit Function
ADD_ClientOfficeAddressLine4:
    On Error GoTo MOD_ClientOfficeAddressLine4
    Rst1.Open "Select ClientOfficeAddressLine4 from Client", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_AgentAddressLine4
MOD_ClientOfficeAddressLine4:
     Conn1.Execute "ALTER Table Client add COLUMN ClientOfficeAddressLine4 Text(255);"
     UpdateSparetable5 = 1
     Exit Function
ADD_AgentAddressLine4:
    On Error GoTo MOD_AgentAddressLine4
    Rst1.Open "Select AgentAddressLine4 from Agent", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_AgentOfficeAddressLine4
MOD_AgentAddressLine4:
     Conn1.Execute "ALTER Table Agent add COLUMN AgentAddressLine4 Text(255);"
    UpdateSparetable5 = 1
    Exit Function
ADD_AgentOfficeAddressLine4:
    On Error GoTo MOD_AgentOfficeAddressLine4
    Rst1.Open "Select AgentOfficeAddressLine4 from Agent", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_LandlordAddressLine4
MOD_AgentOfficeAddressLine4:
    Conn1.Execute "ALTER Table Agent add COLUMN AgentOfficeAddressLine4 Text(255);"
    UpdateSparetable5 = 1
    Exit Function
ADD_LandlordAddressLine4:
    On Error GoTo MOD_LandlordAddressLine4
    Rst1.Open "Select LandlordAddressLine4 from Landlord", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_LandlordOfficeAddressLine4
MOD_LandlordAddressLine4:
    Conn1.Execute "ALTER Table LandLord add COLUMN LandlordAddressLine4 Text(255);"
    UpdateSparetable5 = 1
    Exit Function
ADD_LandlordOfficeAddressLine4:
    On Error GoTo MOD_LandlordOfficeAddressLine4
    Rst1.Open "Select LandlordOfficeAddressLine4 from Landlord", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Mod_SageAccountNumber
    Exit Function
MOD_LandlordOfficeAddressLine4:
    Conn1.Execute "ALTER Table LandLord add COLUMN LandlordOfficeAddressLine4 Text(255);"
    UpdateSparetable5 = 1
    Exit Function
Mod_SageAccountNumber:
   Rst1.Open "Select SageAccountNumber from tenants", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("SageAccountNumber").DefinedSize = 8 Then
        Rst1.Close
        Conn1.Execute "ALTER TABLE tenants ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   LeaseDetails    ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   DemandRecords    ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   DemandRecPreview     ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   tlbChildDemandRecord     ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   tlbDRCurrentPrint    ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   tlbLetterReports     ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   tlbReceipt   ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   Units    ALTER COLUMN SageAccountNumber text(30)"
        Conn1.Execute "ALTER TABLE   PropertyMaintHistory     ALTER COLUMN ReportedBy text(30)"
        Conn1.Execute "ALTER TABLE   NLPosting    ALTER COLUMN ACCOUNT_NUMBER text(30)"
        UpdateSparetable5 = 1
        Exit Function
   Else
        Rst1.Close
        GoTo ADD_TLS_ShoppingCenter
   End If
ADD_TLS_ShoppingCenter:
    'added by anol 20191130
     On Error GoTo MOD_TLS_ShoppingCenter
     Rst1.Open "Select TLS from ShoppingCentre;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_SpareTable5_CC
     Exit Function
MOD_TLS_ShoppingCenter:
    Conn1.Execute "ALTER TABLE ShoppingCentre ADD TLS BIT;"
    Conn1.Execute "Update ShoppingCentre Set TLS=false"
    UpdateSparetable5 = 1
    Exit Function
ADD_SpareTable5_CC:
   Rst1.Open "Select CC from Sparetable5", Conn1, adOpenStatic, adLockReadOnly
   If Rst1("CC").DefinedSize <> 255 Then
        Rst1.Close
        Conn1.Execute "ALTER Table SpareTable5 ALTER COLUMN CC Text(255);"
        UpdateSparetable5 = 1
   Else
        Rst1.Close
'        GoTo ADD_ClientOfficeAddressLine4
   End If
   Exit Function

End Function
Private Sub AllUpdateFunctions(Conn1 As ADODB.Connection)
    'written by anol 20170423
    'Loading of the software is taking a long time. SO I have decided to run this update procdures in the background
'     end of loading
'   DoEvents
'   frmMMain.rtxtMessageDisplay.text = "Running update database and checking in background.. "
'GoTo writeErrors
' Debug.Print "UpdateDatabase start :" & time
   'Call UpdateSparetable5(Conn1)
   If App.Major <= 5 Then
      Do
         iFORM_LOAD = UpdateSparetable5(Conn1)

         If iFORM_LOAD = -1 Then
'            Conn1.Close
'            Set Conn1 = Nothing

            Exit Sub
         End If

      Loop While iFORM_LOAD <> 0
   End If
'Debug.Print "UpdateDatabase End :" & time
'Debug.Print time '18:56:37
   If Not DemandMigration(Conn1) Then
'      Conn1.Close
'      Set Conn1 = Nothing
      Exit Sub
   End If
'Debug.Print time
   bFixingDT = False
 
' Debug.Print time '18:56:42'after fixing 21:53:21
'Debug.Print "UpdateDatabase1 start :" & time
'   If App.Major <= 5 Then
'      Do
'         iFORM_LOAD = UpdateDatabase1(Conn1)
'
'         If iFORM_LOAD = -1 Then
''            Conn1.Close
''            Set Conn1 = Nothing
'
'            Exit Sub
'         End If
'      Loop While iFORM_LOAD <> 0
'   End If
'Debug.Print "UpdateDatabase1 End :" & time
'Debug.Print "Loop"
'Debug.Print "UpdateDatabase1 start :" & time
'    If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
'    Else
          Debug.Print time & " UpdateDatabase4"
          Call UpdateDatabase4(Conn1) 'now execution time 17 sec
          Debug.Print time & " UpdateDatabase4"
        'Debug.Print "UpdateDatabase1 End :" & time
        'Debug.Print time '18:56:58 =16'after fixing 21:53:32=11 saved 5 sec
           Debug.Print "UpdateDatabase2 start :" & time
           If RibbonVersion Then
              Do
                 iFORM_LOAD = UpdateDatabase2(Conn1) 'now execution time 7 sec
        
                 If iFORM_LOAD = -1 Then
        '            Conn1.Close
        '            Set Conn1 = Nothing
        
                    Exit Sub
                 End If
              Loop While iFORM_LOAD <> 0
           End If
          Debug.Print "UpdateDatabase2 End :" & time
           'Debug.Print time '18:56:58
           Debug.Print "UpdateRecords start :" & time
           Call UpdateRecords(Conn1) 'Execution time 13 sec
           Debug.Print "UpdateRecords End :" & time
           EmailArchiving
'   End If
'   ShowMsgInTaskBar "Database Checking Completed ...", "Y", "N"
End Sub
Private Sub LoadTreeviewVer2(Conn1 As ADODB.Connection)
    Debug.Print "Treeview loading start:" & time
    Dim nodX As Node
    Dim TreeArray() As String
    Dim NodeCount As Integer
    'Dim Conn1 As New ADODB.Connection
    Dim rsClient As New ADODB.Recordset
    Dim rsProperty As New ADODB.Recordset
    Dim rsUnit As New ADODB.Recordset
    Dim rsTenant As New ADODB.Recordset
    Dim szStr As String
    Dim i As Integer
    Dim K As Integer
    tvwLandLord.ImageList = imgList
    rsClient.Open "SELECT ClientID,ClientName FROM Client Order by ClientName", Conn1, adOpenStatic, adLockReadOnly
    NodeCount = rsClient.RecordCount
   
    szStr = "SELECT ClientID,PROPERTY.PROPERTYID,PROPERTY.PROPERTYNAME " & _
         "FROM PROPERTY "
    rsProperty.Open szStr, Conn1, adOpenStatic, adLockReadOnly
    NodeCount = NodeCount + rsProperty.RecordCount
    
    szStr = "SELECT UNITS.PROPERTYID,UNITS.UNITNUMBER , UNITS.UNITNAME " & _
         "FROM UNITS " & _
         " order by UNITS.UNITNUMBER ;"
    rsUnit.Open szStr, Conn1, adOpenStatic, adLockReadOnly
    NodeCount = NodeCount + rsUnit.RecordCount
'    rsUnit.Close
    
    szStr = "SELECT LeaseDetails.UnitNumber,Tenants.SageAccountNumber, Tenants.Name, LeaseDetails.LeaseID,LeaseDetails.HeadLease,OLED,EndDate " & _
         "FROM UNITS, LeaseDetails, Tenants " & _
         "WHERE Units.UnitNumber=LeaseDetails.UnitNumber AND " & _
             "Units.Occupied = 'Y' And " & _
             "Tenants.SageAccountNumber = LeaseDetails.SageAccountNumber And " & _
             "LeaseDetails.Status = True ;"
    rsTenant.Open szStr, Conn1, adOpenStatic, adLockReadOnly
    While Not rsTenant.EOF
           'this part is making space at array for sub lease
          If Len(IIf(IsNull(rsTenant("HeadLease").Value), "", rsTenant("HeadLease").Value)) > 0 Then
            NodeCount = NodeCount + 1
          End If
          rsTenant.MoveNext
    Wend
    If rsTenant.RecordCount > 0 Then
        rsTenant.MoveFirst
    End If
    NodeCount = NodeCount + rsTenant.RecordCount * 2 + 1
   
    Debug.Print NodeCount
    ReDim TreeArray(NodeCount, 3)
    'This is a two dimensional array
    'col 0 shall contain Relative
    'col 1 shall contain Key
    'col 2 shall contain Text
    'col 3 shall contain Imagelist index
    
    While Not rsClient.EOF
          'Here I am addding CLIENT
          TreeArray(i, 0) = ""
          TreeArray(i, 1) = rsClient("ClientID").Value & "@" & "CLIENT" + " / " + rsClient("ClientID").Value
          TreeArray(i, 2) = rsClient("ClientName").Value
          TreeArray(i, 3) = 1
          rsClient.MoveNext
          i = i + 1
    Wend
    
    
    While Not rsProperty.EOF
          'Here I am addding Property
          TreeArray(i, 0) = rsProperty("ClientID").Value & "@" & "CLIENT" + " / " + rsProperty("ClientID").Value
          TreeArray(i, 1) = rsProperty("PROPERTYID").Value & "@" & "PROPERTY"
          TreeArray(i, 2) = rsProperty("PROPERTYNAME").Value
          TreeArray(i, 3) = 2
          rsProperty.MoveNext
          i = i + 1
    Wend
     While Not rsUnit.EOF
        'Here I am addding UNITS
          TreeArray(i, 0) = rsUnit("PROPERTYID").Value & "@" & "PROPERTY"
          TreeArray(i, 1) = rsUnit("UNITNUMBER").Value & "@" & "UNITS"
          TreeArray(i, 2) = rsUnit("UNITNAME").Value
          TreeArray(i, 3) = 3
          rsUnit.MoveNext
          i = i + 1
    Wend
    While Not rsTenant.EOF
          'Here I am addding tenant
          TreeArray(i, 0) = rsTenant("UnitNumber").Value & "@" & "UNITS"
          TreeArray(i, 1) = rsTenant("UnitNumber").Value & "$" & rsTenant("SageAccountNumber").Value & "@" & "TENANT"
          TreeArray(i, 2) = rsTenant("Name").Value
          TreeArray(i, 3) = 4
          i = i + 1
          'Here I am addding Lease
          TreeArray(i, 0) = rsTenant("UnitNumber").Value & "$" & rsTenant("SageAccountNumber").Value & "@" & "TENANT"
          TreeArray(i, 1) = rsTenant("LeaseID").Value & "@" & "LEASE"
          TreeArray(i, 2) = "LEASE"
          TreeArray(i, 3) = 6
          i = i + 1
          'Here I am addding Sub Lease
          If Len(IIf(IsNull(rsTenant("HeadLease").Value), "", rsTenant("HeadLease").Value)) > 0 Then
                ReDim Preserve TreeArray(NodeCount, 3)
                TreeArray(i, 0) = rsTenant("LeaseID").Value & "@" & "LEASE"
                TreeArray(i, 1) = rsTenant("HeadLease").Value
                TreeArray(i, 2) = "SUB-LEASE"
                TreeArray(i, 3) = 6
                i = i + 1
          End If
          rsTenant.MoveNext
          
    Wend
    rsClient.Close
    rsUnit.Close
    rsProperty.Close
    rsTenant.Close
    tvwLandLord.Nodes.Clear
    For K = 0 To i - 1
         If TreeArray(K, 0) = "" Then
             Set nodX = tvwLandLord.Nodes.Add(, , TreeArray(K, 1), TreeArray(K, 2), CInt(TreeArray(K, 3)), Int(TreeArray(K, 3)))
         Else
             Set nodX = tvwLandLord.Nodes.Add(TreeArray(K, 0), tvwChild, TreeArray(K, 1), TreeArray(K, 2), CInt(TreeArray(K, 3)), CInt(TreeArray(K, 3)))
         End If
    Next K
   Debug.Print "Treeview loading End:" & time
End Sub
Private Sub MDIForm_Load()

    If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
      Dim adoconn As New ADODB.Connection
      adoconn.Open getConnectionString
      adoconn.Execute "Update shoppingcentre set pws='',SMTP='',UNAME=''"
      adoconn.Close
   End If
    'Cls
    'Tree Loading  13 sec
    'UpdateDatabase1 17 sec
    'UpdateDatabase2 7 sec
    'Update Records  13  sec Subtotal 37 Sec
    '
    'Total 50 Sec

    'Form loading start:08:34:52
    'Tree drawing start:08:34:53
    'Tree drawing End:08:35:05
    'AllUpdateFunctions start :08:35:05
    'UpdateDatabase2 start :08:35:23
    'UpdateDatabase2 End :08:35:29
    'UpdateRecords start :08:35:29
    'UpdateRecords End :08:35:42
    'AllUpdateFunctions End :08:35:42
    'Form loading Ends:08:35:42

   'Debug.Print "Form loading start:" & time
   Dim szTemp  As String

   Me.Move (Screen.Width - Width) / 2, 0
   Me.Caption = "Prestige Property Management Program - " & gCurrentShopCentreName & " - " & "V" & App.Major & "." & App.Minor & "." & App.Revision
   tmDisplayTimer.Enabled = False

   szDataBaseUpdateStatus = "Database is not upto date." & Chr(13) & "Please relogin into the system."

   ExpTrans3rdParty = ExpMMSLicence

   Conn1.Open getConnectionString

   If RibbonVersion Then
'      fraCmdButton.Visible = False
      LoadingRibbon
   Else
'      ACPRibbon1.Visible = False
'      fraCmdButton.Visible = True
      mnuFile.Visible = True
'      mnuView.Visible = True
'      mnuShopCen.Visible = True
'      mnuTools.Visible = True
'      mnuReports.Visible = True
'      mnuHelp.Visible = True
'      mnuAbout.Visible = True

      'Me.Picture = LoadPicture(App.Path + "\Package1\BackGroundImages\London Skyline_.jpg")
   End If

   szReportPath = Value_SecondaryCode("FPATH", "FPATH2", Conn1)
 'load the left client property tree
' Debug.Print time Placed here by anol 20170607
   Conn1.Execute "Update leasedetails set status=false  where OLED=false and EndDate< date() and status=true"
    'Conn1.Execute "Update Units U LEFT JOIN LeaseDetails L ON U.UnitNumber=L.UnitNumber set Occupied='N' where L.Status=true AND "

   'Order by ClientName this clause added by anol 20171016 issue 500
'   Rst1.Open "SELECT ClientID FROM Client Order by ClientName", Conn1, adOpenStatic, adLockReadOnly
'   Debug.Print "Tree drawing start:" & time
'   'Here we are drawing the tree. It takes 12 sec to load the tree
'   While Not Rst1.EOF
'      DrawLandLordTree tvwLandLord, imgList, Rst1!clientID, False, ""
'
'      Rst1.MoveNext
'   Wend
'   Rst1.Close
''   Debug.Print "Tree drawing End:" & time
   'MsgBox "Test"
   Call LoadTreeviewVer2(Conn1)
'   End
   'added by anol 20170426
  ' If UCase(SystemUser) <> "BOSLUSER" And UCase(WS_Name) <> "PCM-DEV2" Then
     Debug.Print "AllUpdateFunctions start :" & time
'    If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
'    Else
        Call AllUpdateFunctions(Conn1)
'    End If
    Debug.Print "AllUpdateFunctions End :" & time
  ' End If
'   If UCase(SystemUser) = "BOSLUSER" And UCase(WS_Name) = "PCM-DEV2" Then
'         Call UpdateRecords(Conn1)
'   End If
''end of loading
'   If App.Major < 5 Then
'      Do
'         iFORM_LOAD = UpdateDatabase
'
'         If iFORM_LOAD = -1 Then
'            Conn1.Close
'            Set Conn1 = Nothing
'
'            Exit Sub
'         End If
'
'      Loop While iFORM_LOAD <> 0
'   End If
''Debug.Print time '18:56:37
'   If Not DemandMigration(Conn1) Then
'      Conn1.Close
'      Set Conn1 = Nothing
'
'      Exit Sub
'   End If
'
'   bFixingDT = False
'' Debug.Print time '18:56:42'after fixing 21:53:21
'   If App.Major <= 5 Then
'      Do
'         iFORM_LOAD = UpdateDatabase1
'
'         If iFORM_LOAD = -1 Then
'            Conn1.Close
'            Set Conn1 = Nothing
'
'            Exit Sub
'         End If
'      Loop While iFORM_LOAD <> 0
'   End If
''Debug.Print time '18:56:58 =16'after fixing 21:53:32=11 saved 5 sec
'   If RibbonVersion Then
'      Do
'         iFORM_LOAD = UpdateDatabase2
'
'         If iFORM_LOAD = -1 Then
'            Conn1.Close
'            Set Conn1 = Nothing
'
'            Exit Sub
'         End If
'      Loop While iFORM_LOAD <> 0
'   End If
'   'Debug.Print time '18:56:58
'
'   UpdateRecords
'
'   EmailArchiving
  
   
   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
      cmdTest.Visible = True
      cmdClientBanks.Visible = True

      Rst1.Open "SELECT UserName FROM UserNames WHERE UserName = 'samrat';", Conn1, adOpenStatic, adLockReadOnly

      If Rst1.EOF Then
         Rst1.Close
         Conn1.Execute "UPDATE Tenants SET spare8 = Email1, spare9 = Email2;"
         Conn1.Execute "UPDATE Tenants SET Email1 = 'developers@pcmuk.net', Email2 = 'developers@pcmuk.net';"
         Conn1.Execute "UPDATE Supplier SET SageSuppAC = SupplierOfficeEmail;"
         Conn1.Execute "UPDATE Supplier SET SupplierOfficeEmail = 'prestige@pcmuk.net';"
         Conn1.Execute "INSERT INTO UserNames (UserName, Password) Values ('samrat', 'pcm');"
         Rst1.Open "SELECT SMTP, UName, Pws, Port FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly

         szTemp = IIf(IsNull(Rst1.Fields.Item("SMTP").Value), "", Rst1.Fields.Item("SMTP").Value) & "#" & _
                  IIf(IsNull(Rst1.Fields.Item("UName").Value), "", Rst1.Fields.Item("UName").Value) & "#" & _
                  IIf(IsNull(Rst1.Fields.Item("Pws").Value), "", Rst1.Fields.Item("Pws").Value) & "#" & _
                  IIf(IsNull(Rst1.Fields.Item("Port").Value), "", Rst1.Fields.Item("Port").Value)

         Conn1.Execute "UPDATE ShoppingCentre " & _
                       "SET SMTP = '192.168.0.2', UName = 'DEVELOPERS', " & _
                           "Pws = 'pcmuk', Port = 25, Field8 = '" & szTemp & "';"
         szSMTPserver = "192.168.0.2"
         szFromEmail = "developers@pcmuk.net"
         szUName = "developers"
         szPws = "pcmuk"
         szPort = 25
         Conn1.Execute "UPDATE tlbClientBanks SET FileLoc_ = FileLoc;"
         Conn1.Execute "UPDATE tlbClientBanks SET FileLoc = 'C:\Samrat\PropertyManagementProgram\Non_Client_Server\Non_SAGE\BlockMng\BACS';"
      End If
      Rst1.Close
   End If

   If iFORM_LOAD = 0 Then _
      szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"

   If LCase(User) = "manager" Then mnuEditUserNames.Enabled = True
   If LCase(User) <> "manager" Then mnuEditUserNames.Enabled = False

   'If Not BatchPaymentLicence Then mnuBatchPayment_.Enabled = False

   TOTAL_SUPPLIERS = 0
   Rst1.Open "SELECT * FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then TOTAL_SUPPLIERS = Rst1.RecordCount
   Rst1.Close
   TOTAL_LESSEES = 0
   Rst1.Open "SELECT * FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then TOTAL_LESSEES = Rst1.RecordCount
   Rst1.Close
   TOTAL_CLIENTS = 0
   Rst1.Open "SELECT * FROM Client;", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then TOTAL_CLIENTS = Rst1.RecordCount
   Rst1.Close

   Conn1.Close
   Set Conn1 = Nothing

'@@@@@@@@@@@@@@@@@@@    TESTING CODE    @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
'      LoadTestCode
'   End If
    
'MissingTable_GlobalRC:
'   MsgBox "The database is not prepared for batch recript. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBatchReceipt & tblBtRptTran"
'   Conn1.Close
'   Set Conn1 = Nothing
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Debug.Print "Form loading Ends:" & time
End Sub

Private Sub EmailArchiving()
'explantion by anol
'Program shall look for this location \Logs\moved_PI.dat , if It does not find this file it will not ask for any question about archive
'if it finds a file moved_PI or moved_SI it shalll append new information about mail/SI in this file
   Dim bSI     As Boolean
   Dim bPI     As Boolean
   Dim szTemp  As String
   Dim szLine  As String
   
   bPI = FileExists(App.Path & "\Logs\moved_PI.dat")
   bSI = FileExists(App.Path & "\Logs\moved_SI.dat")

   If Not bPI Or Not bSI Then
      If MsgBox("Email archiving process is being updated." & Chr(13) & _
                "Please make sure no one is using Prestige." & Chr(13) & _
                "Do you want to proceed now?", vbQuestion + vbYesNo, "Email Archive") = vbYes Then
'        Move Email_PI.dat to server
         If Not bPI Then
            Open DB_PATH & "\AllStuff\Logs\Email_PI.dat" For Append As #1
   '        Read the old log file
            Open App.Path & "\Logs\Email_PI.dat" For Input As #2

   '        Transfer data to the new location
            While Not EOF(2)
               Line Input #2, szLine
               Print #1, szLine
            Wend

            Close #1
            Close #2

   '        Mark the system that archive has been transfered to server
            Open App.Path & "\Logs\moved_PI.dat" For Output As #1
            Close #1
         End If
         If Not bSI Then
            Open DB_PATH & "\AllStuff\Logs\Email_SI.dat" For Append As #1
   '        Read the old log file
            Open App.Path & "\Logs\Email_SI.dat" For Input As #2

   '        Transfer data to the new location
            While Not EOF(2)
               Line Input #2, szLine
               Print #1, szLine
            Wend

            Close #1
            Close #2

   '        Mark the system that archive has been transfered to server
            Open App.Path & "\Logs\moved_SI.dat" For Output As #1
            Close #1
         End If

         MoveTempFolder2Server 'transfer local App.Path & "\Temp\*.pdf file to the Server (i.e where the databse exists)
      End If
   End If
End Sub

Private Function DemandMigration(Conn1 As ADODB.Connection) As Boolean
   Dim szSQL      As String
   'note by anol 2019-08-08 if demand is created and not posted then function shall post in tlbreceipt table
   '
'DoEvents
'   szSQL = "SELECT D.DemandID " & _
'           "FROM   DemandRecords as D, DemandSplitRecords as S " & _
'           "WHERE  D.DemandID = S.DemandID AND " & _
'               "S.TrfReceipt = TRUE AND " & _
'               "D.DemandID NOT IN (" & _
'                  "SELECT DemandRef " & _
'                  "FROM tlbreceipt " & _
'                  "WHERE Type = 1 OR Type = 2)"
'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
    szSQL = "SELECT DemandID from ( SELECT D.DemandID FROM   DemandRecords as D, DemandSplitRecords as S " & _
            "WHERE  D.DemandID = S.DemandID AND   S.TrfReceipt = TRUE ) as X LEFT JOIN tlbreceipt ON tlbreceipt.DemandRef= X.DemandID " & _
            "WHERE  (tlbreceipt.Type = 1 OR tlbreceipt.Type = 2) AND tlbreceipt.DemandRef is NULL"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      
       'keep a log in the error log table
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'DemandMigration called,DmndID: " & Rst1.Fields("DemandID").Value & "' )"
      Rst1.Close
'      szSQL = "UPDATE DemandSplitRecords " & _
'              "SET TrfReceipt = FALSE " & _
'              "WHERE DemandID IN (" & _
'                   "SELECT D.DemandID " & _
'                   "FROM   DemandRecords as D, DemandSplitRecords as S " & _
'                   "WHERE  D.DemandID = S.DemandID AND " & _
'                   "S.TrfReceipt = TRUE AND " & _
'                   "D.DemandID NOT IN (" & _
'                         "SELECT DemandRef " & _
'                         "FROM tlbreceipt " & _
'                         "WHERE Type = 1 OR Type = 2)" & _
'              ");"
'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
        szSQL = "UPDATE DemandSplitRecords  SET TrfReceipt = FALSE WHERE DemandID IN (SELECT DemandID from " & _
                 "( SELECT D.DemandID FROM   DemandRecords as D, DemandSplitRecords as S " & _
                 "WHERE  D.DemandID = S.DemandID AND   S.TrfReceipt = TRUE ) as X LEFT JOIN tlbreceipt ON tlbreceipt.DemandRef= X.DemandID " & _
                 "where ( tlbreceipt.Type = 1 OR tlbreceipt.Type = 2) AND tlbreceipt.DemandRef is NULL );"
            'Debug.Print szSQL
      Conn1.Execute szSQL

      MigrateInvIntoReceipt Conn1
   Else
      Rst1.Close
   End If

   DemandMigration = True
End Function

'  New table has been added in the system #tlbPaymentSplit
'  System will create splits for the payment transactions. In the payment table
'     there are PI and PP transactions. Therefore, system will create splits for
'     both of these transactions.
Private Function UpdateTlbPaymentSplit() As Boolean
   Dim szPP    As String
   Dim szSQL   As String
   Dim adoRst  As New ADODB.Recordset
   Dim adoRst1 As New ADODB.Recordset

   UpdateTlbPaymentSplit = True

   On Error GoTo DoWork
 
'  TrfPayment flag is used to mark the database (not the record) that system has created
'   payment splits  according the  .

   adoRst.Open "SELECT TrfPayment FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly

   adoRst.Close
   Set adoRst = Nothing
   Exit Function

DoWork:

'  FIXED PAYMENT ALLOCATION BUG.
'-----------------------------------------------------------------------------
'  THERE ARE DATA CORRUPTION FOUND IN THE PAYMENT AND ALOCATION TABLE
'  REMOVE THAT REDUNDENT RECORDS WHICH DOES NOT MAKE SENCE
   szSQL = "SELECT P.TransactionID, P.Amount, PT.PP " & _
           "FROM tlbPayment AS P, (" & _
               "SELECT FromTran, SUM(PaymentAmount) AS PP " & _
               "From PayTransactions " & _
               "GROUP BY FromTran" & _
           ") AS PT " & _
           "Where P.TransactionID = PT.FromTran And (P.Amount - P.OSAmount) <> PT.PP;"
'Debug.Print szSQL
   adoRst.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      szSQL = "SELECT * " & _
              "FROM   PayTransactions " & _
              "WHERE  FromTran = " & adoRst.Fields.Item("TransactionID").Value & " AND " & _
                    "PaymentAmount = " & adoRst.Fields.Item("Amount").Value & ";"
      adoRst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      If Not adoRst1.EOF Then
         szSQL = "DELETE * " & _
                 "FROM  PayTransactions " & _
                 "WHERE FromTran = " & adoRst.Fields.Item("TransactionID").Value & " AND " & _
                       "PaymentAmount <> " & adoRst.Fields.Item("Amount").Value & ";"
   'Debug.Print szSQL
         Conn1.Execute szSQL
      End If
      adoRst1.Close

      adoRst.MoveNext
   Wend

   adoRst.Close
'  Recheck -
   szSQL = "SELECT P.TransactionID, P.Amount, PT.PP " & _
           "FROM tlbPayment AS P, ( " & _
               "SELECT FromTran, SUM(PaymentAmount) AS PP " & _
               "From PayTransactions " & _
               "GROUP BY FromTran " & _
           ") AS PT " & _
           "Where P.TransactionID = PT.FromTran And (P.Amount - P.OSAmount) <> PT.PP;"
'Debug.Print szSQL
   adoRst.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   While Not adoRst.EOF
      szSQL = "DELETE * " & _
              "FROM  PayTransactions " & _
              "WHERE FromTran = " & adoRst.Fields.Item("TransactionID").Value & " AND " & _
                    "PaymentAmount > " & adoRst.Fields.Item("Amount").Value & ";"
      Conn1.Execute szSQL

      szSQL = "SELECT SUM(PaymentAmount) AS P " & _
              "FROM   PayTransactions " & _
              "WHERE  FromTran = " & adoRst.Fields.Item("TransactionID").Value & " " & _
              "GROUP BY FromTran"
'Debug.Print szSQL
      adoRst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      If Not adoRst1.EOF Then
         If adoRst1.Fields.Item("P").Value <> adoRst.Fields.Item("Amount").Value Then
            szPP = adoRst.Fields.Item("TransactionID").Value & ", " & szPP
         End If
      End If

      adoRst1.Close
      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing

   If szPP <> "" Then
      MsgBox "Please contact with PCM developers to amend the company data." & Chr(13) & szPP, vbInformation + vbOKOnly, "Purchase Payment Allocation"
      UpdateTlbPaymentSplit = False

      Exit Function
   End If

''  BUG FOUND: PI OS amount in tlbPayment table does not comply with the PayTransactions allocation
''  System found that there is a outstanding amount of a PI in tlbPayment table, however according to
''     PayTransactions allocation there should not be any OS amount of the PI
'
'   szSQL = "SELECT P.TransactionID, P.OSAmount, (P.Amount - PT.T) AS X " & _
'           "FROM tlbPayment AS P, ( " & _
'               "SELECT PT.FromTran, SUM(PaymentAmount) AS T " & _
'               "FROM PayTransactions AS PT " & _
'               "GROUP BY PT.FromTran " & _
'               ") AS PT " & _
'           "Where P.TransactionID = PT.FromTran And P.OSAmount <> P.Amount - PT.T;"
''Debug.Print szSQL
'   adoRst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   While Not adoRst1.EOF
'      Conn1.Execute "UPDATE tlbPayment " & _
'                    "SET OSAmount = " & adoRst1.Fields.Item("X").Value & " " & _
'                    "WHERE TransactionID = " & adoRst1.Fields.Item("TransactionID").Value & ";"
'      adoRst1.MoveNext
'   Wend
'
'   adoRst1.Close
'   Set adoRst1 = Nothing

   Conn1.Execute "ALTER TABLE tblPurInvSRec ADD COLUMN TrfPayment BIT;"

'System is creating PAYMENT SPLITs in the following procedure
   GeneratePaymentSplit Conn1
'------------------------------------------------------------
End Function

Private Sub CreateBankReconSplits()
   Dim rstSrc       As New ADODB.Recordset
   Dim rstDST       As New ADODB.Recordset
   Dim szaTemp()    As String

   rstDST.Open "SELECT * FROM tlbBankReconcilation;", Conn1, adOpenDynamic, adLockOptimistic

   With rstSrc
      .Open "SELECT * FROM tlbPayment WHERE ReconNow <>'';", Conn1, adOpenStatic, adLockReadOnly

      While Not .EOF
         rstDST.AddNew
         rstDST.Fields.Item("MY_ID").Value = UniqueID()
         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("Type").Value
         rstDST.Fields.Item("RefID").Value = .Fields.Item("TransactionID").Value
         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("SageAccountNumber").Value
         rstDST.Fields.Item("UnitID").Value = .Fields.Item("UnitID").Value
         rstDST.Fields.Item("TDate").Value = .Fields.Item("PDate").Value
         rstDST.Fields.Item("DDate").Value = .Fields.Item("DDate").Value
         rstDST.Fields.Item("TRef").Value = .Fields.Item("Ref").Value
         rstDST.Fields.Item("Details").Value = .Fields.Item("Details").Value
         rstDST.Fields.Item("Amount").Value = .Fields.Item("Amount").Value
         rstDST.Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value

         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
         rstDST.Fields.Item("ReconType").Value = szaTemp(1)

         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BankCode").Value
         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NominalCode").Value
         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("ExtRef").Value
         rstDST.Fields.Item("TranMth").Value = .Fields.Item("PayAmtType").Value
         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("SlNumber").Value
         rstDST.Fields.Item("FundID").Value = .Fields.Item("FundID").Value
         rstDST.Fields.Item("Recoverable").Value = .Fields.Item("Recoverable").Value
         rstDST.Update

         .MoveNext
      Wend

      .Close
   End With
   With rstSrc
      .Open "SELECT * FROM tlbReceipt WHERE ReconNow <>'';", Conn1, adOpenStatic, adLockReadOnly

      While Not .EOF
         rstDST.AddNew
         rstDST.Fields.Item("MY_ID").Value = UniqueID()
         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("Type").Value
         rstDST.Fields.Item("RefID").Value = .Fields.Item("TransactionID").Value
         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("SageAccountNumber").Value
         rstDST.Fields.Item("UnitID").Value = .Fields.Item("UnitID").Value
         rstDST.Fields.Item("TDate").Value = .Fields.Item("RDate").Value
         rstDST.Fields.Item("DDate").Value = .Fields.Item("DDate").Value
         rstDST.Fields.Item("TRef").Value = .Fields.Item("Ref").Value
         rstDST.Fields.Item("Details").Value = .Fields.Item("Details").Value
         rstDST.Fields.Item("Amount").Value = .Fields.Item("Amount").Value
         rstDST.Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value

         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
         rstDST.Fields.Item("ReconType").Value = szaTemp(1)

         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BankCode").Value
         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NominalCode").Value
         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("ExtRef").Value
         rstDST.Fields.Item("TranMth").Value = .Fields.Item("RptAmtType").Value
         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("SlNumber").Value
         rstDST.Fields.Item("FundID").Value = .Fields.Item("FundID").Value
         rstDST.Update

         .MoveNext
      Wend

      .Close
   End With
   With rstSrc
      .Open "SELECT * FROM tlbBankPayment WHERE ReconNow <>'';", Conn1, adOpenStatic, adLockReadOnly

      While Not .EOF
         rstDST.AddNew
         rstDST.Fields.Item("MY_ID").Value = UniqueID()
         rstDST.Fields.Item("TransactionType").Value = .Fields.Item("TransactionType").Value
         rstDST.Fields.Item("RefID").Value = .Fields.Item("MY_ID").Value
         rstDST.Fields.Item("AccountNum").Value = .Fields.Item("BANK_AC").Value
         rstDST.Fields.Item("UnitID").Value = .Fields.Item("PropertyID").Value
         rstDST.Fields.Item("TDate").Value = .Fields.Item("TRAN_DATE").Value
         rstDST.Fields.Item("DDate").Value = .Fields.Item("TRAN_DATE").Value
         rstDST.Fields.Item("TRef").Value = .Fields.Item("PROJ_REF").Value
         rstDST.Fields.Item("Details").Value = .Fields.Item("DESCRIPTION").Value
         rstDST.Fields.Item("Amount").Value = .Fields.Item("NET_AMOUNT").Value + .Fields.Item("VAT").Value
         rstDST.Fields.Item("ReconAmount").Value = .Fields.Item("Reconciled").Value

         szaTemp = Split(.Fields.Item("ReconNow").Value, "#")
         rstDST.Fields.Item("ReconDate").Value = CDate(szaTemp(0))
         rstDST.Fields.Item("ReconType").Value = szaTemp(1)

         rstDST.Fields.Item("BankCode").Value = .Fields.Item("BANK_AC").Value
         rstDST.Fields.Item("NominalCode").Value = .Fields.Item("NOMINAL_CODE").Value
         rstDST.Fields.Item("ExtRef").Value = .Fields.Item("PROJ_REF").Value
         rstDST.Fields.Item("SlNumber").Value = .Fields.Item("TRAN_ID").Value
         rstDST.Fields.Item("FundID").Value = .Fields.Item("DEPT_ID").Value
         rstDST.Fields.Item("ClientID").Value = .Fields.Item("ClientID").Value
         rstDST.Update

         .MoveNext
      Wend

      .Close
   End With

   rstDST.Close
   Set rstDST = Nothing
   Set rstSrc = Nothing
End Sub

'Private Function UpdateDatabase1(Conn1 As ADODB.Connection)
'   Dim Rst2       As New ADODB.Recordset
'   Dim Rst3       As New ADODB.Recordset
'   Dim Rst4       As New ADODB.Recordset
'   Dim lSlNumber  As Long
'   Dim i          As Integer
'   Dim iRec       As Integer
'   Dim szEmail    As String
'   Dim szSQL      As String
'   Dim szSQL_     As String
'   Dim szSQL__    As String
'   Dim szSQL___    As String
'   Dim cOS        As Currency
'
''   loopcount = loopcount + 1
'
'
'
''   DoEvents
'   UpdateDatabase1 = 0
'
''   'Resolved by BOSL
'''0000468: Posting dates not implemented correctly
'''added by anol anol 23 Nov 2014
''Modify_TABLE_tlbPayment:
''   On Error GoTo Mod_TABLE_tlbPayment
''
''   Rst1.Open "SELECT PostingDate FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields(0).Type = 135 Then
''      Rst1.Close
''      GoTo NEW_TABLE_MemoDetails
''   Else
''Mod_TABLE_tlbPayment:
''      Rst1.Close
'Debug.Print time
'Debug.Print 134
''      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN PostingDate DateTime;"
'Debug.Print 134
''      GoTo NEW_TABLE_MemoDetails
''   End If
''   UpdateDatabase1 = 1
''   Exit Function
'
' '   Add new column  on 18/08/14 tlbClientBanks
''###############################################################################################################
''Resolved by BOSL
''Issue 458: Batch Payments Crashes with Upgraded Data
''Resolved by anol 18 Aug 2014
''Issue 466: Batch Payments Crashes with Upgraded Data
''Resolved by anol 27 Aug 2014
'ADD_COLUMN_PostingDate:
' On Error GoTo ERR_tblBatchPayment
' Debug.Print 135
' Debug.Print time
'   Rst1.Open "SELECT PostingDate FROM tblBatchPayment;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'Debug.Print 135
'Debug.Print time
'   GoTo ADD_PCB_tlbClientBanks
'ERR_tblBatchPayment:
'Debug.Print time
'Debug.Print 154
'   Conn1.Execute "ALTER TABLE tblBatchPayment ADD COLUMN PostingDate Date;"
'Debug.Print 154
'Debug.Print time
'   UpdateDatabase1 = 1
'   Exit Function
'
'
'
''   Add new column PCB on 29/04/10 tlbClientBanks
''###############################################################################################################
'ADD_PCB_tlbClientBanks:
'   On Error GoTo CHANGE_ADD_PCB_tlbClientBanks
'   Debug.Print 137
'   Debug.Print time
'   Rst1.Open "SELECT PCB FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   Debug.Print 137
'   Debug.Print time
'   GoTo ADD_email_tlbClientBanks
'CHANGE_ADD_PCB_tlbClientBanks:
'Debug.Print time
'Debug.Print 170
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN PCB Currency;"
'Debug.Print 170
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column email on 29/04/XX tlbClientBanks
''###############################################################################################################
'ADD_email_tlbClientBanks:
'   On Error GoTo CHANGE_ADD_email_tlbClientBanks
'   Debug.Print 171
'   Debug.Print time
'   Rst1.Open "SELECT email FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'    Debug.Print time
'    Debug.Print 171
'   GoTo ADDNEW_Memo_Property
'CHANGE_ADD_email_tlbClientBanks:
'Debug.Print time
'Debug.Print 184
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN email TEXT(100);"
'Debug.Print 184
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column Memo on 16/03/10 Property
''###############################################################################################################
'ADDNEW_Memo_Property:
'   On Error GoTo Error_ADDNEW_Memo_Property
'   'Resolved by BOSL
'   'issue 466
'    'Debug.Print
'    Debug.Print time
'Debug.Print 185
'   Rst1.Open "SELECT MemoText FROM Property;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print Err.description
'   Rst1.Close
'Debug.Print time
'Debug.Print 185
'   GoTo ADD_Fund_tlbReceipt
'
'Error_ADDNEW_Memo_Property:
'Debug.Print time
'Debug.Print 202
'   Conn1.Execute "ALTER TABLE Property ADD COLUMN MemoText MEMO"
'Debug.Print 202
''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col - RptAmtType) - DemandRecords"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column FundID on 18/03/10 tlbReceipt
''###############################################################################################################
'ADD_Fund_tlbReceipt:
'   On Error GoTo CHANGE_ADD_Fund_tlbReceipt
'Debug.Print time
'Debug.Print 187
'   Rst1.Open "SELECT FundID FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'Debug.Print time
'Debug.Print 187
'   GoTo Fund_4_RoA
'CHANGE_ADD_Fund_tlbReceipt:
'Debug.Print time
'Debug.Print 217
'   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN FundID Long;"
'Debug.Print 217
'   UpdateDatabase1 = 1
'   Exit Function
'
''***************************************************************************************************************
''                  Check System data - RoA/SRR without Fund 20/03/10 tlbReceipt                                '
''###############################################################################################################
'Fund_4_RoA:
'Debug.Print time
'Debug.Print 188
'   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
'                  "RDate, Details, Amount, BankCode, ExtRef " & _
'             "FROM tlbReceipt " & _
'             "WHERE (Type = 4 OR Type = 23) AND " & _
'                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
'             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print time
'Debug.Print 188
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo FundID_4_PoA
'   End If
'
'   If MsgBox("There are receipts on account and receipt refunds with no fund assigned." & Chr(13) & _
'          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
'          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
'      Load frmFund4RoA_RF
'      frmFund4RoA_RF.Show 1
'   End If
'
'   Rst1.Close
'
''   Add new column FundID on 31/03/10 tlbPayment
''###############################################################################################################
'FundID_4_PoA:
'   On Error GoTo ERROR_Fund_4_PoA
'Debug.Print time
'Debug.Print 190
'   Rst1.Open "SELECT FundID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'Debug.Print time
'Debug.Print 190
'   GoTo DrCr_4_NominalLedger
'ERROR_Fund_4_PoA:
'Debug.Print time
'Debug.Print 256
'   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN FundID Long;"
'Debug.Print 256
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column DrCr on 21/04/10 NominalLedger
''###############################################################################################################
'DrCr_4_NominalLedger:
'   On Error GoTo ERROR_DrCr_4_NominalLedger
'Debug.Print time
'Debug.Print 193
'   Rst1.Open "SELECT DrCr FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'Debug.Print time
'Debug.Print 193
'   GoTo Fund_4_PoA
'ERROR_DrCr_4_NominalLedger:
'Debug.Print time
'Debug.Print 270
'   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN DrCr TEXT(2);"
'Debug.Print 270
'   UpdateDatabase1 = 1
'   Exit Function
'
''***************************************************************************************************************
''                  Check System data - PoA/PPR without Fund 20/03/10 tlbPayment                                '
''###############################################################################################################
'Fund_4_PoA:
'Debug.Print time
'Debug.Print 195
'   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
'                  "PDate, Details, Amount, BankCode, ExtRef " & _
'             "FROM tlbPayment " & _
'             "WHERE (Type = 9 OR Type = 24) AND " & _
'                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
'             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print time 'this one took 5 sec .this is a long time
'Debug.Print 195
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo Fund_4_SR
'   End If
'
'   If MsgBox("There are payment on account and payment refunds with no fund assigned." & Chr(13) & _
'          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
'          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
'      Load frmFund4PoA
'      frmFund4PoA.Show 1
'
'      Rst1.Close
'
'      UpdateDatabase1 = 1
'      Exit Function
'   End If
'
''
''   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
''                  "PDate, Details, Amount, BankCode, ExtRef " & _
''             "FROM tlbPayment " & _
''             "WHERE (Type = 9 OR Type = 24) AND " & _
''                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
''             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.EOF Then
''      Rst1.Close
''      GoTo Fund_4_SR
''   End If
''
''   If MsgBox("There are payment on account and payment refunds with no fund assigned." & Chr(13) & _
''          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
''          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
''      Load frmFund4PoA
''      frmFund4PoA.Show 1
''   End If
''   Rst1.Close
'
'Fund_4_SR:
'Debug.Print time
'Debug.Print 196
'   Rst1.Open "SELECT TransactionID FROM tlbReceipt WHERE (ISNULL(FundID) OR FundID = 0) AND Type = 3;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print time
'Debug.Print 196
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo Receipt_Splits
'   End If
'   Rst1.Close
'
''***************************************************************************************************************
''                  Check System data - System will check receipts' splits in the new table tlbReceiptSplit     '
''                  Existing data will not have any split in the split table.                                   '
''                  If there any receipt found without a split line then system will try to create the split    '
''                  record.
''                                               13/05/2010
''###############################################################################################################
'Receipt_Splits:
'   On Error GoTo RptSpitMissing
'
''  THE FOLLOWING PROCEDURE IS DISABLED FOR CLARKE PROPERTY. I HAVE UPDATED THEIR DATA.
''  THIS PROCEDURE CAUSES THEIR SYSTEM VERY SLOW TO LOAD
''   GenerateReceiptSplit Conn1
'   GoTo ADD_CtrlCode_tlbTransactionTypes
'
'RptSpitMissing:
'   MsgBox "System could not update your database. Please contact with PCM.", vbCritical + vbOKOnly, "Err. Generating Receipt Split"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Add new column CtrlCode on 04/06/2010 tlbTransactionTypes
''###############################################################################################################
'ADD_CtrlCode_tlbTransactionTypes:
'   On Error GoTo CHANGE_ADD_CtrlCode_tlbTransactionTypes
'Debug.Print time
'Debug.Print 197
'   Rst1.Open "SELECT CtrlCode FROM tlbTransactionTypes;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   Debug.Print time
'Debug.Print 197
'   GoTo ADD_Group_tlbTransactionTypes
'
'CHANGE_ADD_CtrlCode_tlbTransactionTypes:
'Debug.Print time
'Debug.Print 364
'   Conn1.Execute "ALTER TABLE tlbTransactionTypes ADD COLUMN CtrlCode TEXT(15);"
'Debug.Print 364
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column Group on 04/02/2014 tlbTransactionTypes
''###############################################################################################################
'ADD_Group_tlbTransactionTypes:
'   On Error GoTo CHANGE_ADD_Group_tlbTransactionTypes
'Debug.Print time
'Debug.Print 198
'   Rst1.Open "SELECT Group FROM tlbTransactionTypes;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   Debug.Print time
'Debug.Print 198
'   GoTo ADD_NEW_TABLE_NLPosting
'
'CHANGE_ADD_Group_tlbTransactionTypes:
'Debug.Print time
'Debug.Print 378
'   Conn1.Execute "ALTER TABLE tlbTransactionTypes ADD COLUMN Group TEXT(50);"
'Debug.Print 378
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 07/06/2010 NLPosting
''###############################################################################################################
'ADD_NEW_TABLE_NLPosting:
'   On Error GoTo CHANGE_ADD_NEW_TABLE_NLPosting
'   Debug.Print time
'   Debug.Print 199
'   Rst1.Open "SELECT * FROM NLPosting;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   Debug.Print time
'    Debug.Print 199
'   GoTo BUGFIX_tlbPayment_OSAmount
'
'CHANGE_ADD_NEW_TABLE_NLPosting:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - NLPosting"
'   UpdateDatabase1 = -1
'   Exit Function
''
'''   Update table 'tlbPayment' data on 16/06/2010 UnitID
'''###############################################################################################################
''UPDATE_tlbPayment_UnitID:
''   On Error GoTo ERROR_UPDATE_tlbPayment_UnitID
''
''   Rst1.Open "SELECT * FROM tlbPayment WHERE ISNULL(UnitID) OR UnitID='';", Conn1, adOpenStatic, adLockReadOnly
''
''   If Not Rst1.EOF Then
''      Rst1.Close
'Debug.Print time
'Debug.Print 405
''      Conn1.Execute "UPDATE tlbPayment, Property SET UnitID = PropertyID " & _
'Debug.print 405
''                    "WHERE ISNULL(UnitID) OR UnitID='';"
''   Else
''      Rst1.Close
''      GoTo BUGFIX_tlbPayment_OSAmount
''   End If
''
''   GoTo BUGFIX_tlbPayment_OSAmount
''
''ERROR_UPDATE_tlbPayment_UnitID:
''   Debug.Print ERR.Number & " " & ERR.description
''   Exit Function
'
''   Bugfix table 'tlbPayment' data on 21/06/2010 OSAmount
''###############################################################################################################
'BUGFIX_tlbPayment_OSAmount:
'Debug.Print time
'Debug.Print 421
'   Rst1.Open "Select * from  tlbPayment WHERE Amount < OSAmount;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print time
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 423
'        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbpayment Amount < OSAmount " & Rst1.Fields("TransactionID").Value & "' )"
'Debug.Print 423
'Debug.Print time
'Debug.Print 424
'        Conn1.Execute "UPDATE tlbPayment " & _
'                 "SET OSAmount = Amount " & _
'                 "WHERE Amount < OSAmount;"
'                 Debug.Print 424
'   End If
'   Rst1.Close
''  THE FOLLOWING SEGMENT OF CODE IS DISABLED FOR CLARKE PROPERTY. I HAVE UPDATED THEIR DATA.
''  THESE CODE CAUSES THEIR SYSTEM VERY SLOW TO LOAD
'''
'''   Update table 'tlbReceiptSplit' data on 01/07/2010
'''###############################################################################################################
''   szsql = "SELECT * FROM tlbReceipt " & _
''             "WHERE TransactionID NOT IN (" & _
''               "SELECT RptHeader FROM tlbReceiptSplit);"
'''Debug.Print szsql
''   Rst1.Open szsql, Conn1, adOpenStatic, adLockReadOnly
''
''   If Not Rst1.EOF Then
''      Rst2.Open "SELECT * FROM tlbReceiptSplit;", Conn1, adOpenDynamic, adLockOptimistic
''
''      While Not Rst1.EOF
''         With Rst2
''            .AddNew
''            .Fields.Item("TransactionID").Value = UniqueID()
''            .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
''            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
''            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
''            .Fields.Item("SplitID").Value = 1
''            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
''            .Fields.Item("Description").Value = Rst1.Fields.Item("Details").Value
''            .Update
''         End With
''         Rst1.MoveNext
''      Wend
''      Rst2.Close
''   End If
''   Rst1.Close
'
''   Fixing data Amount in RptSplit on 08/07/2010
''###############################################################################################################
'Debug.Print time
'Debug.Print 463
'    Rst1.Open "Select * from tlbReceipt AS R, tlbReceiptSplit AS S WHERE S.Amount > R.Amount AND R.Type > 1 AND R.TransactionID = S.RptHeader;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print time
'    If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 465
'        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbReceipt Amount < OSAmount " & Rst1.Fields("TransactionID").Value & "' )"
'Debug.Print 465
'        szSQL = "UPDATE tlbReceipt AS R, tlbReceiptSplit AS S " & _
'             "SET S.Amount = R.Amount " & _
'             "WHERE S.Amount > R.Amount AND " & _
'                  "R.Type > 1 AND " & _
'                  "R.TransactionID = S.RptHeader;"
'Debug.Print time
'Debug.Print 471
'        Conn1.Execute szSQL
'Debug.Print 471
'    End If
'    Rst1.Close
''   Fixing data SlNumber in 'DemandRecords' & 'tlbReceipt' on 05/07/2010
''###############################################################################################################
'   szSQL = "SELECT DMDSLNO, COUNT(DMDSLNO) AS X, TransactionType " & _
'             "From DEMANDRECORDS " & _
'             "GROUP BY DMDSLNO, TransactionType " & _
'             "Having count(DMDSLNO) > 1 " & _
'             "ORDER BY DMDSLNO;"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
''#
'Debug.Print "In SI serial number duplicating!!!!! SAMRAT check it."
''Debug.Print szsql
'      lSlNumber = SlNumber(IIf(Rst1.Fields.Item("TransactionType").Value = 1, "SI", "SC"), "DemandRecords", Conn1)
'Debug.Print time
'Debug.Print 488
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'SI serial number duplicating,dmslno:" & Rst1.Fields.Item(0).Value & "' )"
'Debug.Print 488
'      While Not Rst1.EOF
'         szSQL = "SELECT DMDSLNO " & _
'                   "FROM DEMANDRECORDS " & _
'                   "WHERE DMDSLNO = " & Rst1.Fields.Item(0).Value & " AND " & _
'                     "TransactionType = " & Rst1.Fields.Item("TransactionType").Value & " " & _
'                   "ORDER BY DemandID;"
'         Rst2.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
'         Rst2.MoveNext
'         While Not Rst2.EOF
'            Rst2.Fields.Item(0).Value = lSlNumber
'            Rst2.Update
'            Rst2.MoveNext
'            lSlNumber = lSlNumber + 1
'         Wend
'         Rst2.Close
'         Rst1.MoveNext
'      Wend
'      Rst1.Close
'
'Debug.Print time
'Debug.Print 508
'      Conn1.Execute _
'         "UPDATE tlbReceipt AS R, DemandRecords AS D " & _
'         "SET    R.SlNumber = D.DmdSlNo " & _
'         "WHERE  R.DemandRef = D.DemandID;"
'         Debug.Print 508
'   Else
'      Rst1.Close
'   End If
'
''   Fixing data SlNumber in 'tlbReceipt' on 13/07/2010
''###############################################################################################################
'   szSQL = "UPDATE tlbReceipt AS R, DemandRecords AS D " & _
'             "SET R.SlNumber = D.DmdSlNo " & _
'             "WHERE R.DemandRef = D.DemandID AND " & _
'               "R.SlNumber <> D.DmdSlNo AND R.Type = 1;"
'Debug.Print time
'Debug.Print 522
'   Conn1.Execute szSQL
'Debug.Print 522
'
''   New table on 07/06/2010 tblBtRptTran
''###############################################################################################################
'ADD_NEW_TABLE_tblBtRptTran:
'   On Error GoTo CHANGE_ADD_NEW_TABLE_tblBtRptTran
'
'   Rst1.Open "SELECT * FROM tblBtRptTran;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   GoTo ADD_RptDt_tblBtRptTran
'
'CHANGE_ADD_NEW_TABLE_tblBtRptTran:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBtRptTran"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Add new column RptDt on 19/07/2010 tblBtRptTran
''###############################################################################################################
'ADD_RptDt_tblBtRptTran:
'   On Error GoTo CHANGE_ADD_RptDt_tblBtRptTran
'
'   Rst1.Open "SELECT RptDt FROM tblBtRptTran;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo FundID_4_PayTransactions
'CHANGE_ADD_RptDt_tblBtRptTran:
'Debug.Print time
'Debug.Print 548
'   Conn1.Execute "ALTER TABLE tblBtRptTran ADD COLUMN RptDt TEXT(15);"
'Debug.Print 548
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column FundID on 29/07/10 PayTransactions
''###############################################################################################################
'FundID_4_PayTransactions:
'   On Error GoTo ERROR_Fund_4_PayTransactions
'
'   Rst1.Open "SELECT FundID FROM PayTransactions;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_FundID_RptTransactions
'ERROR_Fund_4_PayTransactions:
'Debug.Print time
'Debug.Print 562
'   Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN FundID Long;"
'Debug.Print 562
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new a columns FundID on 03/08/2010 RptTransactions
''###############################################################################################################
'ADD_FundID_RptTransactions:
'   On Error GoTo CHANGE_ADD_FundID_RptTransactions
'
'   Rst1.Open "SELECT FundID FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_LeaseValue_LeaseDetails
'CHANGE_ADD_FundID_RptTransactions:
'Debug.Print time
'Debug.Print 577
'   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN FundID Long;"
'Debug.Print 577
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns LeaseValue on 11/08/10 LeaseDetails
''###############################################################################################################
'ADD_LeaseValue_LeaseDetails:
'   On Error GoTo CHANGE_ADD_LeaseValue_LeaseDetails
'
'   Rst1.Open "SELECT LeaseValue FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_GPrataDmd_LeaseDetails
'
'CHANGE_ADD_LeaseValue_LeaseDetails:
'Debug.Print time
'Debug.Print 593
'   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN LeaseValue CURRENCY;"
'Debug.Print 593
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns GPrataDmd on 11/08/10 LeaseDetails
''###############################################################################################################
'ADD_GPrataDmd_LeaseDetails:
'   On Error GoTo CHANGE_ADD_GPrataDmd_LeaseDetails
'
'   Rst1.Open "SELECT GPrataDmd FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_SCYE_LServiceCharges
'
'CHANGE_ADD_GPrataDmd_LeaseDetails:
'Debug.Print time
'Debug.Print 610
'   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN GPrataDmd BIT;"
'Debug.Print 610
'Debug.Print time
'Debug.Print 611
'   Conn1.Execute "UPDATE LeaseDetails SET GPrataDmd = TRUE;"
'Debug.Print 611
'
'Debug.Print time
'Debug.Print 613
'   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN FDD TEXT(25);"
'Debug.Print 613
'Debug.Print time
'Debug.Print 614
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN FDD TEXT(25);"
'Debug.Print 614
'Debug.Print time
'Debug.Print 615
'   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN FDD TEXT(25);"
'Debug.Print 615
'Debug.Print time
'Debug.Print 616
'   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN StopAutoDmd BIT;"
'Debug.Print 616
'Debug.Print time
'Debug.Print 617
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN StopAutoDmd BIT;"
'Debug.Print 617
'Debug.Print time
'Debug.Print 618
'   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN StopAutoDmd BIT;"
'Debug.Print 618
'Debug.Print time
'Debug.Print 619
'   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN GPD BIT;"
'Debug.Print 619
'Debug.Print time
'Debug.Print 620
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN GPD BIT;"
'Debug.Print 620
'Debug.Print time
'Debug.Print 621
'   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN GPD BIT;"
'Debug.Print 621
'
'   UpdateFDD Conn1
'   UpdateStopAutoDmd Conn1
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns SCYE on 08/08/2011 LServiceCharges
''###############################################################################################################
'ADD_SCYE_LServiceCharges:
'   On Error GoTo CHANGE_ADD_SCYE_LServiceCharges
'
'   Rst1.Open "SELECT SCYE FROM LServiceCharges;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_DueDate_tblPurInv
'
'CHANGE_ADD_SCYE_LServiceCharges:
'Debug.Print time
'Debug.Print 641
'   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN SCYE BIT;"
'Debug.Print 641
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns DueDate on 11/08/10 tblPurInv
''###############################################################################################################
'ADD_DueDate_tblPurInv:
'   On Error GoTo CHANGE_ADD_DueDate_tblPurInv
'
'   Rst1.Open "SELECT DueDate FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_PropertyID_ClientGlobalData
'
'CHANGE_ADD_DueDate_tblPurInv:
'Debug.Print time
'Debug.Print 657
'   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN DueDate TEXT(20);"
'Debug.Print 657
'   UpdateDatabase1 = 1
'   Exit Function
''
'''   Add a columns JobID on 18/08/10 DemandSplitRecords_tlbReceiptSplit
'''###############################################################################################################
''ADD_JobID_DemandSplitRecords_tlbReceiptSplit:
''   On Error GoTo CHANGE_ADD_JobID_DemandSplitRecords_tlbReceiptSplit
''
''   Rst1.Open "SELECT JobID FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   GoTo ADD_PropertyID_ClientGlobalData
''
''CHANGE_ADD_JobID_DemandSplitRecords_tlbReceiptSplit:
'Debug.Print time
'Debug.Print 673
''   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN JobID TEXT(10);"
'Debug.Print 673
'Debug.Print time
'Debug.Print 674
''   Conn1.Execute "ALTER TABLE tlbReceiptSplit ADD COLUMN JobID TEXT(10);"
'Debug.Print 674
''   UpdateDatabase1 = 1
''   Exit Function
'
''   Add a columns PropertyID on 24/08/10 ClientGlobalData
''###############################################################################################################
'ADD_PropertyID_ClientGlobalData:
'   On Error GoTo CHANGE_ADD_PropertyID_ClientGlobalData
'
'   Rst1.Open "SELECT PropertyID FROM ClientGlobalData;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_RecoverableExp_ChargeTypes
'
'CHANGE_ADD_PropertyID_ClientGlobalData:
'Debug.Print time
'Debug.Print 689
'   Conn1.Execute "ALTER TABLE ClientGlobalData ADD COLUMN PropertyID TEXT(4);"
'Debug.Print 689
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns RecoverableExp on 24/08/10 ChargeTypes
''###############################################################################################################
'ADD_RecoverableExp_ChargeTypes:
'   On Error GoTo CHANGE_ADD_RecoverableExp_ChargeTypes
'
'   Rst1.Open "SELECT RecoverableExp FROM ChargeTypes;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADDNEW_COL_RCCComments_Tenants
'
'CHANGE_ADD_RecoverableExp_ChargeTypes:
'Debug.Print time
'Debug.Print 704
'   Conn1.Execute "ALTER TABLE ChargeTypes ADD COLUMN RecoverableExp TEXT(1);"
'Debug.Print 704
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column RCCComments on 23/09/2010 Tenants
''###############################################################################################################
'ADDNEW_COL_RCCComments_Tenants:
'   On Error GoTo MISSING_ADDNEW_COL_RCCComments_Tenants
'
'   Rst1.Open "SELECT RCCComments FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_COL_RCC_Property
'
'MISSING_ADDNEW_COL_RCCComments_Tenants:
'Debug.Print time
'Debug.Print 720
'   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN RCCComments TEXT(250);"
'Debug.Print 720
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column RCC on 24/09/2010 Property
''###############################################################################################################
'ADDNEW_COL_RCC_Property:
'   On Error GoTo MISSING_ADDNEW_COL_RCC_Property
'
'   Rst1.Open "SELECT RCC FROM Property;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_COL_Comments_Supplier
'
'MISSING_ADDNEW_COL_RCC_Property:
'Debug.Print time
'Debug.Print 736
'   Conn1.Execute "ALTER TABLE Property ADD COLUMN RCC TEXT(1);"
'Debug.Print 736
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column Comments on 06/12/2010 Supplier
''###############################################################################################################
'ADDNEW_COL_Comments_Supplier:
'   On Error GoTo MISSING_ADDNEW_COL_Comments_Supplier
'
'   Rst1.Open "SELECT Comments FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_PCB_ABR_tlbClientBanks
'
'MISSING_ADDNEW_COL_Comments_Supplier:
'Debug.Print time
'Debug.Print 752
'   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN Comments TEXT(200);"
'Debug.Print 752
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column PCB_ABR on 24/12/10 tlbClientBanks
''###############################################################################################################
'ADD_PCB_ABR_tlbClientBanks:
'   On Error GoTo CHANGE_ADD_PCB_ABR_tlbClientBanks
'
'   Rst1.Open "SELECT PCB_ABR FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo CheckData_Multiple_SI_tlbReceipt
'
'CHANGE_ADD_PCB_ABR_tlbClientBanks:
'Debug.Print time
'Debug.Print 767
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN PCB_ABR Currency;"
'Debug.Print 767
'Debug.Print time
'Debug.Print 768
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN SOB Currency;"
'Debug.Print 768
'Debug.Print time
'Debug.Print 769
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN LSD TEXT(50);"
'Debug.Print 769
'   UpdateDatabase1 = 1
'   Exit Function
'
''  Check SI in receipt table on 22/10/2010 tlbReceipt
''  There is a data corraption has arirse. Duplicate SI has been exported to receipt table.
''  System identify them and clear from the db.
''###############################################################################################################
'CheckData_Multiple_SI_tlbReceipt:
'
'   On Error GoTo ERROR_CheckData_Multiple_SI_tlbReceipt
'
'   szSQL = "SELECT X.SageAccountNumber, X.DemandRef " & _
'             "FROM " & _
'               "(" & _
'                 "SELECT COUNT(DemandRef) AS C, SageAccountNumber, DemandRef " & _
'                 "FROM tlbReceipt " & _
'                 "WHERE type = 1 " & _
'                 "GROUP by demandref, SageAccountNumber " & _
'               ") AS X " & _
'             "Where C > 1;"
'
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      Rst1.Close
'      If MsgBox("System will update your data." + Chr(10) + _
'                "Please make sure everyone is out of the system before you click YES.", vbYesNo, "Receipt Record Update") = vbNo Then
'         Exit Function
'      Else
'         szSQL = "SELECT X.DemandRef, X.C " & _
'                   "FROM " & _
'                     "(" & _
'                       "SELECT COUNT(DemandRef) AS C, DemandRef " & _
'                       "FROM tlbReceipt " & _
'                       "WHERE type = 1 " & _
'                       "GROUP by demandref, SageAccountNumber " & _
'                     ") AS X " & _
'                   "Where C > 1;"
'         Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'         While Not Rst1.EOF
'            szSQL = "SELECT * " & _
'                      "FROM tlbReceipt " & _
'                      "WHERE DemandRef = " & Rst1.Fields.Item("DemandRef").Value & " " & _
'                      "ORDER BY TransactionID;"
'
'            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'            For i = 1 To Rst1.Fields.Item("C").Value - 1
'Debug.Print time
'Debug.Print 818
'               Conn1.Execute "DELETE tlbReceiptSplit.* " & _
'                             "FROM tlbReceipt, tlbReceiptSplit " & _
'                             "WHERE tlbReceiptSplit.RptHeader = tlbReceipt.TransactionID AND " & _
'                                 "tlbReceipt.TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"
'
'Debug.Print time
'Debug.Print 823
'               Conn1.Execute "DELETE * " & _
'                             "FROM tlbReceipt " & _
'                             "WHERE TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"
'               Rst2.MoveNext
'            Next i
'            Rst2.Close
'            Rst1.MoveNext
'         Wend
'      End If
'   End If
'
'   Rst1.Close
'
'   GoTo ADD_EmailDmd_Tenants
'
'ERROR_CheckData_Multiple_SI_tlbReceipt:
'   MsgBox "System could not update your database. Please contact with PCM Consulting Ltd.", vbCritical + vbInformation + "Receipt Correction Error"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column EmailDmd on 14/02/2011 Tenants
''###############################################################################################################
'ADD_EmailDmd_Tenants:
'   On Error GoTo CHANGE_ADD_EmailDmd_Tenants
'
'   Rst1.Open "SELECT EmailDmd FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_EmailSt_Tenants
'
'CHANGE_ADD_EmailDmd_Tenants:
'Debug.Print time
'Debug.Print 854
'   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailDmd BIT;"
'Debug.Print 854
'Debug.Print time
'Debug.Print 855
'   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare12;"
'Debug.Print 855
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column EmailSt on 31/03/2011 Tenants
''###############################################################################################################
'ADD_EmailSt_Tenants:
'   On Error GoTo CHANGE_ADD_EmailSt_Tenants
'
'   Rst1.Open "SELECT EmailSt FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_CombEmail_Tenants
'
'CHANGE_ADD_EmailSt_Tenants:
'Debug.Print time
'Debug.Print 870
'   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailSt BIT;"
'Debug.Print 870
'Debug.Print time
'Debug.Print 871
'   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare11;"
'Debug.Print 871
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column CombEmail on 31/03/2011 Tenants
''###############################################################################################################
'ADD_CombEmail_Tenants:
'   On Error GoTo CHANGE_ADD_CombEmail_Tenants
'
'   Rst1.Open "SELECT CombEmail FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_CombEmail_tlbDRCurrentPrint
'
'CHANGE_ADD_CombEmail_Tenants:
'Debug.Print time
'Debug.Print 886
'   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN CombEmail BIT;"
'Debug.Print 886
'Debug.Print time
'Debug.Print 887
'   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare10;"
'Debug.Print 887
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column CombEmail on 01/04/2011 tlbDRCurrentPrint
''###############################################################################################################
'ADD_CombEmail_tlbDRCurrentPrint:
'   On Error GoTo CHANGE_ADD_CombEmail_tlbDRCurrentPrint
'
'   Rst1.Open "SELECT CombEmail FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_UName_ShoppingCentre
'
'CHANGE_ADD_CombEmail_tlbDRCurrentPrint:
'Debug.Print time
'Debug.Print 902
'   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN CombEmail BIT;"
'Debug.Print 902
'Debug.Print time
'Debug.Print 903
'   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint DROP COLUMN spare11;"
'Debug.Print 903
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column UName on 05/04/2011 ShoppingCentre
''###############################################################################################################
'ADD_UName_ShoppingCentre:
'   On Error GoTo CHANGE_ADD_UName_ShoppingCentre
'
'   Rst1.Open "SELECT UName FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo RESIZE_SMTP_ShoppingCentre
'
'CHANGE_ADD_UName_ShoppingCentre:
'Debug.Print time
'Debug.Print 918
'   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN UName TEXT(100);"
'Debug.Print 918
'Debug.Print time
'Debug.Print 919
'   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN Pws TEXT(50);"
'Debug.Print 919
'Debug.Print time
'Debug.Print 920
'   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN Port TEXT(5);"
'Debug.Print 920
'Debug.Print time
'Debug.Print 921
'   Conn1.Execute "ALTER TABLE ShoppingCentre DROP COLUMN Field4;"
'Debug.Print 921
'Debug.Print time
'Debug.Print 922
'   Conn1.Execute "ALTER TABLE ShoppingCentre DROP COLUMN Field5;"
'Debug.Print 922
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Extend the field size SMTP on 05/04/2011 ShoppingCentre
''###############################################################################################################
'RESIZE_SMTP_ShoppingCentre:
'
'   On Error GoTo MissingTable_RESIZE_SMTP_ShoppingCentre
'
'   Rst1.Open "SELECT SMTP FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields.Item("SMTP").DefinedSize = 15 Then
'      Rst1.Close
'      Set Rst1 = Nothing
'
'Debug.Print time
'Debug.Print 938
'      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN SMTP TEXT(100)"
'Debug.Print 938
'   End If
'   Rst1.Close
'   Set Rst1 = Nothing
'
'   GoTo ADDNEW_COL_RecoverablePt_tblPurInvSRec
'
'MissingTable_RESIZE_SMTP_ShoppingCentre:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - SMTP of DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column RecoverablePt on 24/05/2011 tblPurInvSRec
''###############################################################################################################
'ADDNEW_COL_RecoverablePt_tblPurInvSRec:
'   On Error GoTo MISSING_ADDNEW_COL_RecoverablePt_tblPurInvSRec
'
'   Rst1.Open "SELECT RecoverablePt FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_NEW_TABLE_Fund
'
'MISSING_ADDNEW_COL_RecoverablePt_tblPurInvSRec:
'Debug.Print time
'Debug.Print 962
'   Conn1.Execute "ALTER TABLE tblPurInvSRec ADD COLUMN RecoverablePt Single;"
'Debug.Print 962
'Debug.Print time
'Debug.Print 963
'   Conn1.Execute "UPDATE tblPurInvSRec SET RecoverablePt = 100;"
'Debug.Print 963
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 07/06/XXXX Fund
''###############################################################################################################
'ADD_NEW_TABLE_Fund:
'   On Error GoTo CHANGE_ADD_NEW_TABLE_Fund
'
'   Rst1.Open "SELECT * FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   GoTo ADDNEW_COL_CategoryCode_Fund
'
'CHANGE_ADD_NEW_TABLE_Fund:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - Fund"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Add new column CategoryCode on 01/07/2011 Fund
''###############################################################################################################
'ADDNEW_COL_CategoryCode_Fund:
'   On Error GoTo MISSING_ADDNEW_COL_CategoryCode_Fund
'
'   Rst1.Open "SELECT CategoryCode FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_COL_Recoverable_tlbPayment
'
'MISSING_ADDNEW_COL_CategoryCode_Fund:
'Debug.Print time
'Debug.Print 993
'   Conn1.Execute "ALTER TABLE Fund ADD COLUMN CategoryCode BYTE;"
'Debug.Print 993
'Debug.Print time
'Debug.Print 994
'   Conn1.Execute "UPDATE Fund SET CategoryCode = 1  WHERE INSTR(FundName, 'Rent')>0;"
'Debug.Print 994
'Debug.Print time
'Debug.Print 995
'   Conn1.Execute "UPDATE Fund SET CategoryCode = 2  WHERE INSTR(FundName, 'SERVICE')>0;"
'Debug.Print 995
'Debug.Print time
'Debug.Print 996
'   Conn1.Execute "UPDATE Fund SET CategoryCode = 3  WHERE INSTR(FundName, 'INSURANCE')>0;"
'Debug.Print 996
'Debug.Print time
'Debug.Print 997
'   Conn1.Execute "UPDATE Fund SET CategoryCode = 4  WHERE ISNULL(CategoryCode);"
'Debug.Print 997
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column Recoverable on 04/07/11 tlbPayment
''###############################################################################################################
'ADDNEW_COL_Recoverable_tlbPayment:
'   On Error GoTo MISSING_ADDNEW_COL_Recoverable_tlbPayment
'
'   Rst1.Open "SELECT Recoverable FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo TypeIE_4_NominalLedger
'MISSING_ADDNEW_COL_Recoverable_tlbPayment:
'Debug.Print time
'Debug.Print 1011
'   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN Recoverable SINGLE;"
'Debug.Print 1011
'Debug.Print time
'Debug.Print 1012
'   Conn1.Execute "UPDATE tlbPayment SET Recoverable = 0;"
'Debug.Print 1012
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column TypeIE on 23/08/11 NominalLedger
''###############################################################################################################
'TypeIE_4_NominalLedger:
'   On Error GoTo ERROR_TypeIE_4_NominalLedger
'
'   Rst1.Open "SELECT TypeIE FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADDNEW_REC_IE
'ERROR_TypeIE_4_NominalLedger:
'Debug.Print time
'Debug.Print 1026
'   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN TypeIE TEXT(2);"
'Debug.Print 1026
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new record on 23/08/11 PrimarySecondaryCode
''###############################################################################################################
'ADDNEW_REC_IE:
'
'   On Error GoTo MissingRec_REC_IE
'
'   With Rst1
'
'      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'IE';", Conn1, adOpenStatic, adLockReadOnly
'
'      If .EOF Then
'         .Close
'         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         !Code = "IE"
'         !Value = "INC_EXP"
'         .Update
'         .Close
'         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item(0).Value = "IE"
'         .Fields.Item(1).Value = "INC"
'         .Fields.Item(2).Value = "Income"
'         .Fields.Item(3).Value = "Nominal Ledger Type - Income"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "IE"
'         .Fields.Item(1).Value = "EXP"
'         .Fields.Item(2).Value = "Expenditure"
'         .Fields.Item(3).Value = "Nominal Ledger Type - Expenditure"
'         .Update
'         .Close
'      End If
'      .Close
'
'      .Open "SELECT Code FROM SecondaryCode WHERE PrimaryCode = 'IE' AND Value = 'Balance Sheet';", Conn1, adOpenStatic, adLockReadOnly
'
'      If .EOF Then
'         .Close
'         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item(0).Value = "IE"
'         .Fields.Item(1).Value = "BS"
'         .Fields.Item(2).Value = "Balance Sheet"
'         .Fields.Item(3).Value = "Nominal Ledger Type - Balance Sheet"
'         .Update
'      End If
'      .Close
'   End With
'
'   GoTo MODIFY_DEPT_ID_tblPurInvSRec
'
'MissingRec_REC_IE:
''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Modify DATA TYPE DEPT_ID on 24/10/07 tblPurInvSRec
''###############################################################################################################
'MODIFY_DEPT_ID_tblPurInvSRec:
'
'   On Error GoTo CHANGE_MODIFY_DEPT_ID_tblPurInvSRec
'
'   Rst1.Open "SELECT DEPT_ID FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields(0).Type = 3 Then
'      Rst1.Close
'   Else
'      Rst1.Close
'Debug.Print time
'Debug.Print 1099
'      Conn1.Execute "ALTER TABLE tblPurInvSRec ALTER COLUMN DEPT_ID LONG;"
'Debug.Print 1099
'      GoTo CHANGE_MODIFY_DEPT_ID_tblPurInvSRec
'   End If
'   'If Month(Date) = 6 Then 'added by anol 2019-06-12 issue 785
'    'GoTo CHECK_DATA_CORRUPTION_RECEIPT
'   'Else
'    GoTo CHECK_DATA_CORRUPTION_DEMAND
'   'End If
'CHANGE_MODIFY_DEPT_ID_tblPurInvSRec:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check Data corruption in Demand & Receipt table on 10/02/2011
''NOTE:
''     SYSTEM INVESTIGATES THE DATA CORRUPTION.
''     CAUSES - SYSTEM COULD NOT UPDATE SI IN THE RECEIPT TABLE IF THE SI WAS MODIFIED.
''     SOLUTION - WHEN USER WILL GET THIS MSG -> MEANS THERE ARE SOME CORRUPTION.
''                1. DELETE THE SI FROM THE RECEIPT TABLE
''                2. UPDATE THE TrfReceipt = FALSE IN THE DEMAND SPLIT TABLE
''###############################################################################################################
'CHECK_DATA_CORRUPTION_DEMAND:
'   szSQL = "SELECT R.TransactionID, R.Amount, R.DemandRef, S.DT " & _
'             "FROM tlbReceipt AS R, " & _
'                 "(SELECT D.DemandID,  SUM(S.TotalAmount) AS DT " & _
'                  "FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                  "GROUP BY D.DemandID) AS S " & _
'             "WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND " & _
'                  "ROUND(R.Amount, 2) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 2);"
'
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'
'      GoTo CHECK_DATA_CORRUPTION_RECEIPT 'CHECK_BACS_EMAIL_tlbClientBanks
'   End If
''   Debug.Print szSQL
'   MsgBox "DATA ERROR: PCM_1254" & Chr(13) & "PLEASE CONTACT WITH PCM SUPPORT.", vbCritical + vbOKOnly, "PRESTIGE SYSTEM"
'Debug.Print time
'Debug.Print 1138
''   Conn1.Execute "Delete from tlbReceipt where transactionid=" & Rst1("TransactionID").Value & ""
'Debug.Print 1138
'Debug.Print time
'Debug.Print 1139
''   Conn1.Execute "Delete from tlbReceiptSplit where RptHeader=" & Rst1("TransactionID").Value & ""
'Debug.Print 1139
'Debug.Print time
'Debug.Print 1140
''   Conn1.Execute "Update DemandSplitRecords set TrfReceipt = FALSE  where DemandID=" & Rst1("DemandRef").Value & ""
'Debug.Print 1140
''   MigrateInvIntoReceipt Conn1
'   Rst1.Close
'   UpdateDatabase1 = -1
'   Exit Function
'
''  Check Data corruption in the Receipt Split table.
''     Data corrupt when user saves a receipt against a multiple SI splits.
''     System saved duplicated SR splits when there is multiple line of SI.
''Solution:
''  1. Collect the list of TransactionID which are corrupted.
''  2. Remove all splits from RptSplits.
''  3. Get the list of allocation splits from RptTransaction
''  4. Create all splits in tlbReceiptSplit according to RptTransaction's splits
''
''###############################################################################################################
'CHECK_DATA_CORRUPTION_RECEIPT:
''  Check is there any transaction in the receipt table with 0 value in the header but non-0 value in the split.
''     if found then make the split value to 0. otherwise divident by zero will arise.
'   szSQL = "SELECT R.TransactionID " & _
'           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
'                                  "From tlbReceiptSplit " & _
'                                  "GROUP BY RptHeader " & _
'                                 ") AS Q " & _
'           "WHERE R.TransactionID=Q.RptHeader AND " & _
'                 "Q.A > 0 AND R.Amount=0;"
''Debug.Print szSQL
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   While Not Rst1.EOF
'Debug.Print time
'Debug.Print 1170
'      Conn1.Execute "UPDATE tlbReceiptSplit " & _
'                    "SET    Amount = 0, OSAmount = 0 " & _
'                    "WHERE  RptHeader = " & Rst1.Fields.Item(0).Value & ";"
'      Rst1.MoveNext
'   Wend
'
'   Rst1.Close
'
'''  Update SI split in tlbReceiptSplit according to DemandSplit where they dont match and they has not paid full/part.
''   szSQL = "UPDATE tlbReceiptSplit AS RS, tlbReceipt AS R, " & _
''                  "DemandRecords as D, DemandSplitRecords as DS " & _
''           "SET    RS.Amount = DS.TotalAmount, RS.OSAmount = DS.TotalAmount " & _
''           "WHERE  D.DemandID = DS.DemandID AND " & _
''                  "R.TransactionID = RS.RptHeader AND " & _
''                  "R.DemandRef = D.DemandID AND " & _
''                  "DS.TotalAmount = RS.OSAmount AND " & _
''                  "RS.Amount <> DS.TotalAmount AND " & _
''                  "DS.SplitID = RS.SplitID;"
'''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1189
''   Conn1.Execute szSQL
'Debug.Print 1189
''
'''  Update SI split in tlbReceiptSplit according to DemandSplit where they dont match and they has not paid full/part.
''   szSQL = "UPDATE tlbReceiptSplit AS RS, tlbReceipt AS R, " & _
''                  "DemandRecords as D, DemandSplitRecords as DS " & _
''           "SET    RS.Amount = DS.TotalAmount, RS.OSAmount = DS.TotalAmount " & _
''           "WHERE  D.DemandID = DS.DemandID AND " & _
''                  "R.TransactionID = RS.RptHeader AND " & _
''                  "R.DemandRef = D.DemandID AND " & _
''                  "RS.Amount = RS.OSAmount AND " & _
''                  "RS.Amount <> DS.TotalAmount AND " & _
''                  "DS.SplitID = RS.SplitID;"
''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1202
''   Conn1.Execute szSQL
'Debug.Print 1202
''
''Sol 1: Collect the list of TransactionID which are corrupted.
'   szSQL = "SELECT R.TransactionID, Q.A-R.Amount AS D, Q.A / R.Amount AS R " & _
'           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
'                                  "From tlbReceiptSplit " & _
'                                  "GROUP BY RptHeader " & _
'                                 ") AS Q " & _
'           "WHERE R.TransactionID = Q.RptHeader AND " & _
'                 "ROUND(R.Amount, 2) <> ROUND(Q.A, 2) AND " & _
'                 "R.Amount+Q.A>0 AND R.OSAmount + Q.A <> R.Amount;"
''Debug.Print szSQL
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo CHECK_DEMANDS_MISSED_MIGRATION
'   End If
'
'   iRec = Rst1.RecordCount
'   i = 0
'   While Not Rst1.EOF
'      If Val(Rst1.Fields.Item("D").Value) > 0 And _
'            CInt(Rst1.Fields.Item("R").Value) = Val(Rst1.Fields.Item("R").Value) Then
'Debug.Print time
'Debug.Print 1226
'            Conn1.Execute "DELETE * FROM tlbReceiptSplit " & _
'                       "WHERE RptHeader = " & Rst1.Fields.Item("TransactionID").Value & " AND " & _
'                           "SplitID > 1;"
'         i = i + 1
'      End If
'      Rst1.MoveNext
'   Wend
'   Rst1.Close
'
'   If iRec > 0 Then
'Debug.Print time
'Debug.Print 1236
''      Conn1.Execute "DELETE * " & _
'Debug.print 1236
''                    "FROM tlbReceiptSplit AS S " & _
''                    "WHERE S.RptHeader NOT IN " & _
''                         "(SELECT TransactionID FROM  tlbReceipt AS R)"
''issue  381 SQL that is taking long time to complete 20170512 fixed by anol
'Debug.Print time
'Debug.Print 1241
'      Conn1.Execute "DELETE S.*  FROM tlbReceiptSplit AS S LEFT JOIN tlbReceipt AS R ON  S.RptHeader = R.TransactionID  WHERE R.TransactionID is NULL"
'Debug.Print 1241
'   End If
'   If iRec <> i Then
'      MsgBox "Warning: This Company data need to be updated. Please contact with PCM.", vbExclamation + vbOKOnly, "Data Update"
'      GoTo CHECK_DEMANDS_MISSED_MIGRATION
'   End If
'
''   While Not Rst1.EOF
''      If Rst1.Fields.Item("D").Value > 0 Then
'''Sol  2. Remove all splits from RptSplits.
'Debug.Print time
'Debug.Print 1251
''         Conn1.Execute "DELETE * " & _
'Debug.print 1251
''                       "FROM   tlbReceiptSplit " & _
''                       "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";"
''
'''Sol  3. Get the list of allocation splits from RptTransaction
''         szSQL = "SELECT * " & _
''                 "FROM   RptTransactions " & _
''                 "WHERE  FromTran = " & Rst1.Fields.Item("TransactionID").Value & ";"
''         Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''
'''Sol  4. Create all splits in tlbReceiptSplit according to RptTransaction's splits
''         Rst3.Open "SELECT * FROM tlbReceiptSplit", Conn1, adOpenDynamic, adLockOptimistic
''         While Not Rst2.EOF
''            With Rst3
''               .AddNew
''               .Fields.Item("TransactionID").Value = UniqueID()
''               .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
''''#
'''               .Fields.Item("FundID").Value = flxSPayment.TextMatrix(iRow, 14)
'''               .Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 10)
'''               If flxSPayment.TextMatrix(iRow, 0) = "" Then
'''                  .Fields.Item("SplitID").Value = -1
'''               Else
'''                  .Fields.Item("SplitID").Value = flxSPayment.TextMatrix(iRow, 1)
'''               End If
'''               .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
'''               .Fields.Item("Description").Value = flxSPayment.TextMatrix(iRow, 7)
'''               .Fields.Item("AllocTranID").Value = flxSPayment.TextMatrix(iRow, 19)
''               .Update
''            End With
''
''
''            Rst2.MoveNext
''         Wend
''         Rst2.Close
''      End If
''      If Rst1.Fields.Item("D").Value < 0 Then
''      End If
''      Rst1.MoveNext
''   Wend
'
'
''  MigrateInvIntoReceipt method missed some demands to export to Receipt table.
''  This procedure will reset the flag to export to receipt.
''  When user will open the demand form, MigrateInvIntoReceipt method will export them again.
''  THIS MATHOD HAS TO BE RUN EVERY TIME WHEN SYSTEM OPENS.
''###############################################################################################################
'CHECK_DEMANDS_MISSED_MIGRATION:
'
'
''  Code has been removed to DemandMigration function
'
'   GoTo CHECK_BACS_EMAIL_tlbClientBanks
''  BACS From email address setting in the BACS form
''NOTE:
''     IF EMAIL IN THE tlbClientBanks = NULL AND Email1 IN ShoppingCentre <> NULL THEN
''     SYSTEM WILL COPY THE Email1 INTO EMAIL OF tlbClientBanks
''###############################################################################################################
'CHECK_BACS_EMAIL_tlbClientBanks:
'   szSQL = "SELECT Email1 " & _
'           "FROM   ShoppingCentre " & _
'           "WHERE  Email1 <> '';"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount = 0 Then GoTo NO_ACCTION
'
'   szEmail = Rst1.Fields.Item("Email1").Value
'   Rst1.Close
'   szSQL = "SELECT email " & _
'             "FROM   tlbClientBanks " & _
'             "WHERE  email = '' OR ISNULL(email);"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount = 0 Then GoTo NO_ACCTION
'   Rst1.Close
'
'Debug.Print time
'Debug.Print 1327
'   Conn1.Execute "UPDATE tlbClientBanks " & _
'                 "SET    email = '" & szEmail & "' " & _
'                 "WHERE  email = '' OR ISNULL(email);"
'   GoTo ADDNEW_COL_SentByEmail_DemandRecords
'
'NO_ACCTION:
'   Rst1.Close
'   GoTo ADDNEW_COL_SentByEmail_DemandRecords
'
''   Add new column SentByEmail on 16/09/2011 DemandRecords
''###############################################################################################################
'ADDNEW_COL_SentByEmail_DemandRecords:
'   On Error GoTo MISSING_ADDNEW_COL_SentByEmail_DemandRecords
'
'   Rst1.Open "SELECT SentByEmail FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_COL_CT_tlbBankPayment
'
'MISSING_ADDNEW_COL_SentByEmail_DemandRecords:
'Debug.Print time
'Debug.Print 1348
'   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN SentByEmail BYTE;"
'Debug.Print 1348
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column CT on 27/09/2011 tlbBankPayment
''###############################################################################################################
'ADDNEW_COL_CT_tlbBankPayment:
'   On Error GoTo MISSING_ADDNEW_COL_CT_tlbBankPayment
'
'   Rst1.Open "SELECT CT FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo FIX_OS_tlbReceiptSplit
'
'MISSING_ADDNEW_COL_CT_tlbBankPayment:
'Debug.Print time
'Debug.Print 1364
'   Conn1.Execute "ALTER TABLE tlbBankPayment DROP COLUMN spare7;"
'Debug.Print 1364
'Debug.Print time
'Debug.Print 1365
'   Conn1.Execute "ALTER TABLE tlbBankPayment DROP COLUMN spare8;"
'Debug.Print 1365
'Debug.Print time
'Debug.Print 1366
'   Conn1.Execute "ALTER TABLE tlbBankPayment ADD  COLUMN CT TEXT(1);"
'Debug.Print 1366
'
'   UpdateBR_BP_SameAcc
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Fix OS balance in the receipt_split on 06/10/2011
''   There is an inconsistency found in the total OS of splits and the OS of the receipt header.
''###############################################################################################################
'FIX_OS_tlbReceiptSplit:
'   On Error GoTo MISSING_FIX_OS_tlbReceiptSplit
'
'   Rst1.Open "SELECT R.TransactionID, R.OSAmount AS R_OS, S.S_OS " & _
'             "FROM tlbReceipt AS R, ( " & _
'                     "SELECT S.RptHeader, SUM(S.OSAmount) AS S_OS " & _
'                     "FROM tlbReceiptSplit AS S " & _
'                     "GROUP BY S.RptHeader " & _
'                     ") AS S " & _
'             "WHERE R.TransactionID =  S.RptHeader AND " & _
'                     "ROUND(R.OSAmount, 2) <> ROUND(S.S_OS, 2) " & _
'             "ORDER BY S.RptHeader;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount > 0 Then        'need to fix data. inconsistent data found
'      While Not Rst1.EOF
'         If Rst1.Fields.Item("R_OS").Value = 0 Then
'Debug.Print time
'Debug.Print 1392
'            Conn1.Execute "UPDATE tlbReceiptSplit " & _
'                          "SET    OSAmount = 0 " & _
'                          "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";"
'         End If
'         If Rst1.Fields.Item("R_OS").Value > 0 Then
'            cOS = Round(CCur(Rst1.Fields.Item("R_OS").Value), 2)
'            Rst2.Open "SELECT * FROM tlbReceiptSplit " & _
'                      "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";", _
'                      Conn1, adOpenDynamic, adLockOptimistic
'            While Not Rst2.EOF
'               If cOS = 0 Then
'                  Rst2.Fields.Item("OSAmount").Value = 0
'               End If
'               If cOS <= Round(CCur(Rst2.Fields.Item("Amount").Value), 2) And cOS > 0 Then
'                  Rst2.Fields.Item("OSAmount").Value = cOS
'                  cOS = 0
'               End If
'               If cOS > Round(CCur(Rst2.Fields.Item("Amount").Value), 2) Then
'                  Rst2.Fields.Item("OSAmount").Value = Rst2.Fields.Item("Amount").Value
'                  cOS = cOS - Rst2.Fields.Item("Amount").Value
'               End If
'
'               Rst2.MoveNext
'            Wend
'            Rst2.Close
'         End If
'
'         Rst1.MoveNext
'      Wend
'   End If
'
'   Rst1.Close
'
'   GoTo FIX_FUND_tlbBankPayment
'
'MISSING_FIX_OS_tlbReceiptSplit:
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check the system Bank Transactions which are without fund.
''   If there is any transaction found, then system will update the Fund = 1
''###############################################################################################################
'FIX_FUND_tlbBankPayment:
'
'   szSQL = "UPDATE tlbBankPayment " & _
'           "SET DEPT_ID = 1 " & _
'           "WHERE DEPT_ID = '' OR ISNULL(DEPT_ID);"
'Debug.Print time
'Debug.Print 1439
'   Conn1.Execute szSQL
'Debug.Print 1439
'
''   Add new column PrintBBF on 20/10/2011 tlbDRCurrentPrint
''###############################################################################################################
'ADDNEW_COL_PrintBBF_tlbDRCurrentPrint:
'   On Error GoTo MISSING_ADDNEW_COL_PrintBBF_tlbDRCurrentPrint
'
'   Rst1.Open "SELECT PrintBBF FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo FIX_EXPORT_tblpurinv
'
'MISSING_ADDNEW_COL_PrintBBF_tlbDRCurrentPrint:
'Debug.Print time
'Debug.Print 1453
'   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN PrintBBF BIT;"
'Debug.Print 1453
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check the system PI Transactions which have not exported to payment table.
''   If there is any transaction found, then system will export them
''###############################################################################################################
'FIX_EXPORT_tblpurinv:
''   szSQL = "SELECT MY_ID " & _
''           "FROM tblPurInv " & _
''           "WHERE TransactionType <> 25 AND MY_ID NOT IN (SELECT PI AS MY_ID FROM tlbPayment WHERE PI <> '') AND " & _
''               "TOTAL_AMOUNT <> 0;"
'
''issue  381 SQL that is taking long time to complete 20170512 fixed by anol
'
'  szSQL = "SELECT MY_ID FROM tblPurInv P LEFT JOIN tlbPayment M  ON P.MY_ID = M.PI WHERE P.TransactionType <> 25 AND P.TOTAL_AMOUNT <> 0 AND M.PI is NULL"
''Debug.Print szSQL
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      MsgBox "Your data need to be updated. Please contact with PCM", vbCritical + vbOKOnly, "PI not in Split table"
'      Rst1.Close
'      UpdateDatabase1 = -1
'
'' MigratePIIntoPayment method export only PI's header. MigratePIIntoPayment does not create the split in the
''  payment split table. Technically, there should not any PI in the purchase invoice table. because, when
''  users create a PI, system immidiately exports the transaction to the PP and PP_split table automatically.
''08/08/2012
'      Exit Function
'
''      szSQL = "UPDATE tblPurInv " & _
''              "SET    TrfPayment = FALSE " & _
''              "WHERE  MY_ID NOT IN (SELECT PI AS MY_ID FROM tlbPayment WHERE PI <> '') AND " & _
''                  "TOTAL_AMOUNT <> 0;"
'   'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
'     szSQL = "UPDATE tblPurInv P LEFT JOIN tlbPayment M ON P.MY_ID = M.PI   SET   P.TrfPayment = FALSE WHERE    TOTAL_AMOUNT <> 0 AND  PI is NULL;"
''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1490
'      Conn1.Execute szSQL
'Debug.Print 1490
'
'      MigratePIIntoPayment Conn1
'   End If
'   Rst1.Close
'
''   Add a column ULC on 27/10/2011 DemandTypes
''###############################################################################################################
'ADD_ULC_DemandTypes:
'   On Error GoTo CHANGE_ADD_ULC_DemandTypes
'
'   Rst1.Open "SELECT ULC FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo FIX_DEMANDTYPE_LLEASE
'
'CHANGE_ADD_ULC_DemandTypes:
'Debug.Print time
'Debug.Print 1508
'   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN ULC BIT;"
'Debug.Print 1508
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check the system for wrong demand type setup in the lease.
''   If there is any transaction found, then system will warn user and if user wants
''   they will able to print the list of lease need to be fix manually.
''   27/10/2011
''   The reports are generating only DemandType and Lease tables entry.
''   But still system can generating this warning by checking the demand and split table
''   Please check szSQL___
''   08/01/2014
''###############################################################################################################
'FIX_DEMANDTYPE_LLEASE:
'   On Error GoTo BUG_FIX_DEMANDTYPE_LLEASE
'
'   If bFixingDT Then GoTo MODIFY_TypeIE_NominalLedger
'
'   szSQL = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.BRDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LRentCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.BRDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.BRDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "Where PropertyID <> LeaseProperty"
''Debug.Print szSQL
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   szSQL_ = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.SCDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LServiceCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.SCDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.SCDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "WHERE PropertyID <> LeaseProperty"
''Debug.Print szSQL
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   szSQL__ = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.InsuranceDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LInsuranceCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.InsuranceDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.InsuranceDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "WHERE PropertyID <> LeaseProperty"
''Debug.Print szSQL
'   Rst3.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'' check other charges
'   szSQL___ = "SELECT ID, P1, P2 " & _
'           "FROM (" & _
'               "SELECT S.TypeOfDemand AS ID, U.PropertyID AS P1, DT.PropertyID AS P2 " & _
'               "FROM DemandRecords AS D, DemandSplitRecords AS S, Tenants AS T, LeaseDetails AS L, Units AS U, DemandTypes AS DT " & _
'               "WHERE L.Status AND D.DemandID = S.DemandID AND D.SageAccountNumber = T.SageAccountNumber AND " & _
'                     "T.SageAccountNumber = L.SageAccountNumber AND L.UnitNumber = U.UnitNumber AND " & _
'                     "S.TypeOfDemand = DT.ID " & _
'               "GROUP BY S.TypeOfDemand, U.PropertyID, DT.PropertyID " & _
'           ") AS Q " & _
'           "WHERE Q.P1 <> Q.P2"
''Debug.Print szSQL
'   Rst4.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Or Not Rst2.EOF Or Not Rst3.EOF Or Not Rst4.EOF Then
'      If MsgBox("There are some lease demand type need to fix." & _
'                "Do you want to do it now?", vbCritical + vbYesNo, "Demand Types") = vbNo Then
'
'         Rst1.Close
'         Rst2.Close
'         Rst3.Close
'         Rst4.Close
'         GoTo MODIFY_TypeIE_NominalLedger
'      End If
''      FIX_MODE__DT = True
'
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL & ");"
''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1589
'      Conn1.Execute szSQL
'Debug.Print 1589
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL_ & ");"
''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1594
'      Conn1.Execute szSQL
'Debug.Print 1594
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL__ & ");"
''Debug.Print szSQL
'Debug.Print time
'Debug.Print 1599
'      Conn1.Execute szSQL
'Debug.Print 1599
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_RC.rpt"
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_SC.rpt"
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_IC.rpt"
'
'      bFixingDT = True
'   Else
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ULC = FALSE;"
'Debug.Print time
'Debug.Print 1609
'      Conn1.Execute szSQL
'Debug.Print 1609
'   End If
'   Rst1.Close
'   Rst2.Close
'   Rst3.Close
'   Rst4.Close
'
'   GoTo MODIFY_TypeIE_NominalLedger
'
'BUG_FIX_DEMANDTYPE_LLEASE:
'   MsgBox "System could not fix your data. Please contact with PCM.", vbInformation + vbOKOnly, "Demand Types"
'   UpdateDatabase1 = 1
'   Exit Function
''
'''   Modify DATA TYPE NLTypeCode on 01/11/11 NLType
'''###############################################################################################################
''MODIFY_NLTypeCode_NLType:
''   On Error GoTo CHANGE_MODIFY_NLTypeCode_NLType
''
''   Rst1.Open "SELECT NLTypeCode FROM NLType;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields(0).Type = 202 Then
''      Rst1.Close
''   Else
''      Rst1.Close
'Debug.Print time
'Debug.Print 1634
''      Conn1.Execute "ALTER TABLE NominalLedger DROP CONSTRAINT NLTypeNominalLedger;"
'Debug.Print 1634
'Debug.Print time
'Debug.Print 1635
''      Conn1.Execute "ALTER TABLE NLType ALTER COLUMN NLTypeCode TEXT(50);"
'Debug.Print 1635
'Debug.Print time
'Debug.Print 1636
''      Conn1.Execute "ALTER TABLE NominalLedger ALTER COLUMN Type TEXT(50);"
'Debug.Print 1636
''   End If
''
''   GoTo MODIFY_TypeIE_NominalLedger
''
''CHANGE_MODIFY_NLTypeCode_NLType:
'''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
''   UpdateDatabase1 = 1
''   Exit Function
'
''   Modify DATA TYPE TypeIE on 01/11/11 NominalLedger
''###############################################################################################################
'MODIFY_TypeIE_NominalLedger:
'
'   On Error GoTo CHANGE_MODIFY_TypeIE_NominalLedger
'
'   Rst1.Open "SELECT TypeIE FROM NominalLedger WHERE TypeIE IN ('1', '2', '3');", Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 1655
'      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'IN' WHERE TypeIE = '1';"
'Debug.Print 1655
'Debug.Print time
'Debug.Print 1656
'      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'EX' WHERE TypeIE = '2';"
'Debug.Print 1656
'Debug.Print time
'Debug.Print 1657
'      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'BS' WHERE TypeIE = '3';"
'Debug.Print 1657
'   End If
'   Rst1.Close
'
'   GoTo MODIFY_Code_SecondaryCode
'
'CHANGE_MODIFY_TypeIE_NominalLedger:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Modify DATA Code on 01/11/11 SecondaryCode
''###############################################################################################################
'MODIFY_Code_SecondaryCode:
'   On Error GoTo CHANGE_MODIFY_Code_SecondaryCode
'
'   Rst1.Open "SELECT Code FROM SecondaryCode WHERE Code IN ('1', '2', '3');", Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 1676
'      Conn1.Execute "UPDATE SecondaryCode SET Code = 'IN' WHERE Code = 'INC' AND PrimaryCode = 'IE';"
'Debug.Print 1676
'Debug.Print time
'Debug.Print 1677
'      Conn1.Execute "UPDATE SecondaryCode SET Code = 'EX' WHERE Code = 'EXP' AND PrimaryCode = 'IE';"
'Debug.Print 1677
'   End If
'   Rst1.Close
'
'   GoTo ADD_EmailSC_Tenants
'
'CHANGE_MODIFY_Code_SecondaryCode:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column EmailSC on 16/11/2011 Tenants
''###############################################################################################################
'ADD_EmailSC_Tenants:
'   On Error GoTo CHANGE_ADD_EmailSC_Tenants
'
'   Rst1.Open "SELECT EmailSC FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_EmailSC_tlbDRCurrentPrint
'
'CHANGE_ADD_EmailSC_Tenants:
'Debug.Print time
'Debug.Print 1699
'   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailSC BIT;"
'Debug.Print 1699
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column EmailSC on 28/11/2011 tlbDRCurrentPrint
''###############################################################################################################
'ADD_EmailSC_tlbDRCurrentPrint:
'   On Error GoTo CHANGE_ADD_EmailSC_tlbDRCurrentPrint
'
'   Rst1.Open "SELECT EmailSC FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_SentStName_DemandRecords
'
'CHANGE_ADD_EmailSC_tlbDRCurrentPrint:
'Debug.Print time
'Debug.Print 1714
'   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN EmailSC BIT;"
'Debug.Print 1714
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column SentStName on 09/12/2011 DemandRecords
''###############################################################################################################
'ADD_SentStName_DemandRecords:
'   On Error GoTo CHANGE_ADD_SentStName_DemandRecords
'
'   Rst1.Open "SELECT SentStName FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_ReportedBy_PropertyMaintHistory
'
'CHANGE_ADD_SentStName_DemandRecords:
'Debug.Print time
'Debug.Print 1729
'   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN SentStName TEXT(100);"
'Debug.Print 1729
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new columns ReportedBy.. on 15/12/2011 PropertyMaintHistory
''###############################################################################################################
'ADD_ReportedBy_PropertyMaintHistory:
'   On Error GoTo CHANGE_ADD_ReportedBy_PropertyMaintHistory
'
'   Rst1.Open "SELECT ReportedBy FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo LeaseDetails_Lessee_Duplicate
'
'CHANGE_ADD_ReportedBy_PropertyMaintHistory:
'Debug.Print time
'Debug.Print 1744
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedBy TEXT(50);"
'Debug.Print 1744
'Debug.Print time
'Debug.Print 1745
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN AssignedIL TEXT(1);"
'Debug.Print 1745
'Debug.Print time
'Debug.Print 1746
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedIS TEXT(1);"
'Debug.Print 1746
'Debug.Print time
'Debug.Print 1747
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Urgent TEXT(1);"
'Debug.Print 1747
'Debug.Print time
'Debug.Print 1748
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Instruction MEMO;"
'Debug.Print 1748
'   UpdateDatabase1 = 1
'   Exit Function
''***************************************************************************************************************
''                  Check System data - System will check lease details table for multiple lessee's account     '
''                  If system found any lessee more than one time in the lease table as active leases           '
''                  then system will stop user to use the system and push them to report it to us               '
''                                                                                                              '
''                                               21/12/2011                                                     '
''###############################################################################################################
'LeaseDetails_Lessee_Duplicate:
'   Rst1.Open "SELECT * " & _
'             "From " & _
'             "( " & _
'              "SELECT COUNT(SageAccountNumber) AS A, SageAccountNumber " & _
'              "From LeaseDetails " & _
'              "Where Status " & _
'              "GROUP BY SageAccountNumber " & _
'              ") AS Q " & _
'             "Where Q.a > 1;", Conn1, adOpenStatic, adLockReadOnly
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo ADD_TemplateID_tlbLetterReports
'   Else
'      MsgBox "The system found an inconsistency in your database. Please contact PCM Consulting Support.", vbCritical + vbOKOnly, "Err. Multiple Lessee"
'      Rst1.Close
'      UpdateDatabase1 = -1
'      Exit Function
'   End If
'
''   Add new column TemplateID on 21/12/2011 tlbLetterReports
''###############################################################################################################
'ADD_TemplateID_tlbLetterReports:
'   On Error GoTo CHANGE_ADD_TemplateID_tlbLetterReports
'
'   Rst1.Open "SELECT TemplateID FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADDNEW_REC_TEMP_TYPE
'
'CHANGE_ADD_TemplateID_tlbLetterReports:
'Debug.Print time
'Debug.Print 1789
'   Conn1.Execute "ALTER TABLE tlbLetterReports ADD COLUMN TemplateID LONG;"
'Debug.Print 1789
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new record TEMP_TYPE on 22/12/11 PrimaryCode
''###############################################################################################################
'ADDNEW_REC_TEMP_TYPE:
'   On Error GoTo MissingData_ADDNEW_REC_TEMP_TYPE
'
'   With Rst1
'      .Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'TEMP_TYPE';", Conn1, adOpenStatic, adLockReadOnly
'
'      If .EOF Then
'         .Close
'         .Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item("Code").Value = "TEMP_TYPE"
'         .Fields.Item("Value").Value = "LETTER TEMPLATE TYPE"
'         .Fields.Item("Flexible").Value = False
'         .Update
'         .Close
'         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
'         .Fields.Item("Code").Value = "LT"
'         .Fields.Item("Value").Value = "Letter Template"
'         .Update
'         .AddNew
'         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
'         .Fields.Item("Code").Value = "RT"
'         .Fields.Item("Value").Value = "Reminder Template"
'         .Update
'         .AddNew
'         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
'         .Fields.Item("Code").Value = "OT"
'         .Fields.Item("Value").Value = "Other Template"
'         .Update
'      End If
'      .Close
'   End With
'
'   GoTo ADD_TempType_Template
'
'MissingData_ADDNEW_REC_TEMP_TYPE:
''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - TEMP_TYPE) - tlbReceipt"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column TempType on 22/12/2011 Template
''###############################################################################################################
'ADD_TempType_Template:
'
'   On Error GoTo CHANGE_ADD_TempType_Template
'
'   Rst1.Open "SELECT TempType FROM Template;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_ReportedFrom_PropertyMaintHistory
'
'CHANGE_ADD_TempType_Template:
'Debug.Print time
'Debug.Print 1849
'   Conn1.Execute "ALTER TABLE Template ADD COLUMN TempType TEXT(10);"
'Debug.Print 1849
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new columns ReportedFrom on 21/01/2012 PropertyMaintHistory
''###############################################################################################################
'ADD_ReportedFrom_PropertyMaintHistory:
'   On Error GoTo CHANGE_ADD_ReportedFrom_PropertyMaintHistory
'
'   Rst1.Open "SELECT ReportedFrom FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_Fund_TenantDeposit
'
'CHANGE_ADD_ReportedFrom_PropertyMaintHistory:
'Debug.Print time
'Debug.Print 1864
'   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedFrom TEXT(1);"
'Debug.Print 1864
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column FundID on 26/01/2012 TenantDeposit
''###############################################################################################################
'ADD_Fund_TenantDeposit:
'   On Error GoTo CHANGE_ADD_Fund_TenantDeposit
'
'   Rst1.Open "SELECT FundID FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo MODIFY_RefundRef_TenantDeposit
'
'CHANGE_ADD_Fund_TenantDeposit:
'Debug.Print time
'Debug.Print 1879
'   Conn1.Execute "ALTER TABLE TenantDeposit ADD COLUMN FundID Long;"
'Debug.Print 1879
'Debug.Print time
'Debug.Print 1880
'   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN TenantDeposit TEXT(50);"
'Debug.Print 1880
'Debug.Print time
'Debug.Print 1881
'   Conn1.Execute "ALTER TABLE TenantDeposit ALTER COLUMN DepositID TEXT(50);"
'Debug.Print 1881
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Modify DATA TYPE RefundRef on 26/01/2012 TenantDeposit
''###############################################################################################################
'MODIFY_RefundRef_TenantDeposit:
'   On Error GoTo CHANGE_MODIFY_RefundRef_TenantDeposit
'
'   Rst1.Open "SELECT RefundRef FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields(0).Type = 202 Then
'      Rst1.Close
'   Else
'      Rst1.Close
'Debug.Print time
'Debug.Print 1896
'      Conn1.Execute "ALTER TABLE TenantDeposit ALTER COLUMN RefundRef TEXT(50);"
'Debug.Print 1896
'      GoTo CHANGE_MODIFY_RefundRef_TenantDeposit
'   End If
'
'   GoTo ALTER_TENANT_ADDRESS_SecondaryCode
'
'CHANGE_MODIFY_RefundRef_TenantDeposit:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Altered Data TenantAddress on 23/12/2011 SecondaryCode
''###############################################################################################################
'ALTER_TENANT_ADDRESS_SecondaryCode:
'
'   Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'TENANT ADDRESS';", Conn1, adOpenStatic, adLockReadOnly
'   If Not Rst1.EOF Then
'      Rst1.Close
'      Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'TENANT ADDRESS';", Conn1, adOpenDynamic, adLockOptimistic
'      Rst1.Fields.Item("Value").Value = "Lessee Address"
'      Rst1.Update
'      Rst1.Close
'      Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'ALTERNATIVE ADDRESS';", Conn1, adOpenDynamic, adLockOptimistic
'      Rst1.Fields.Item("Value").Value = "Alternative Address"
'      Rst1.Update
'   End If
'   Rst1.Close
'
''   Add new column TYPE on 08/05/2013 Supplier
''###############################################################################################################
'ADDNEW_COL_TYPE_Supplier:
'   On Error GoTo MISSING_ADDNEW_COL_TYPE_Supplier
'
'   Rst1.Open "SELECT TYPE FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo DUPLICATE_ID_LCLSA
'
'MISSING_ADDNEW_COL_TYPE_Supplier:
'Debug.Print time
'Debug.Print 1936
'   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN TYPE TEXT(20);"
'Debug.Print 1936
'Debug.Print time
'Debug.Print 1937
'   Conn1.Execute "UPDATE Supplier SET TYPE = 'SUPPLIER';"
'Debug.Print 1937
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check Duplicate ID on 12/01/2012 Lessee, Client, Landlord, Supplier & MA
''###############################################################################################################
'DUPLICATE_ID_LCLSA:
'   szSQL = "SELECT ID "
'   szSQL = szSQL & "FROM (SELECT ID, COUNT(ID) AS C "
'   szSQL = szSQL & "FROM ("
'   szSQL = szSQL & "SELECT SupplierID AS ID "
'   szSQL = szSQL & "FROM Supplier WHERE TYPE = 'SUPPLIER' UNION ALL "
'   szSQL = szSQL & "SELECT ClientID AS ID "
'   szSQL = szSQL & "FROM Client UNION ALL "
'   szSQL = szSQL & "SELECT SageAccountNumber AS ID "
'   szSQL = szSQL & "FROM Tenants UNION ALL "
'   szSQL = szSQL & "SELECT AgentID AS ID "
'   szSQL = szSQL & "FROM Agent UNION ALL "
'   szSQL = szSQL & "SELECT LandlordID AS ID "
'   szSQL = szSQL & "From Landlord "
'   szSQL = szSQL & ") "
'   szSQL = szSQL & " GROUP BY ID "
'   szSQL = szSQL & ") "
'   szSQL = szSQL & "WHERE C > 1;"
'
''Debug.Print szSQL
'
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then                                  'Duplicate ID found
'Debug.Print time
'Debug.Print 1967
'      Conn1.Execute "DELETE Tenants.* " & _
'                    "FROM Tenants, Landlord " & _
'                    "WHERE Tenants.SageAccountNumber = Landlord.LandlordID;"
'      Rst1.Close
'      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'      If Not Rst1.EOF Then                                  'Duplicate ID found
'         szSQL = SQL2String(Rst1, 0)
'         Rst1.Close
'
'         MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
'      Else
'         Rst1.Close
'      End If
'   Else
'      Rst1.Close
'   End If
''
'''  Export Clients into Supplier table on 08/05/2013         ~~~NEVER REMOVED~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''###############################################################################################################
''   Rst1.Open "SELECT * " & _
''             "FROM Client " & _
''             "WHERE ClientID NOT IN (" & _
''               "SELECT SupplierID FROM Supplier WHERE TYPE = 'CLIENT');", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
''      While Not Rst1.EOF
''         Rst2.AddNew
''         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("ClientID").Value
''         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("ClientName").Value
''         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("ClientAddressLine1").Value
''         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("ClientAddressLine2").Value
''         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("ClientAddressLine3").Value
''         Rst2.Fields.Item("SupplierPostCode").Value = IIf(IsNull(Rst1.Fields.Item("ClientPostCode").Value), "", Rst1.Fields.Item("ClientPostCode").Value)
''         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
''         Rst2.Fields.Item("TYPE").Value = "CLIENT"
''         Rst2.Update
''         Rst1.MoveNext
''      Wend
''      Rst2.Close
''   End If
''   Rst1.Close
''
'''  Export Agents into Supplier table on 08/05/2013         ~~~NEVER REMOVED~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''###############################################################################################################
''   Rst1.Open "SELECT * " & _
''             "FROM Agent " & _
''             "WHERE AgentID NOT IN (" & _
''               "SELECT SupplierID FROM Supplier WHERE TYPE = 'AGENT');", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
''      While Not Rst1.EOF
''         Rst2.AddNew
''         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("AgentID").Value
''         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("AgentName").Value
''         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("AgentAddressLine1").Value
''         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("AgentAddressLine2").Value
''         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("AgentAddressLine3").Value
''         Rst2.Fields.Item("SupplierPostCode").Value = Rst1.Fields.Item("AgentPostCode").Value
''         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
''         Rst2.Fields.Item("TYPE").Value = "AGENT"
''         Rst2.Update
''         Rst1.MoveNext
''      Wend
''      Rst2.Close
''   End If
''   Rst1.Close
'
''   Add new column TEmail on 06/02/2012 tlbLetterReports
''###############################################################################################################
'ADD_TEmail_tlbLetterReports:
'
'   On Error GoTo CHANGE_ADD_TEmail_tlbLetterReports
'
'   Rst1.Open "SELECT TEmail FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo RESIZE_SageSuppAC_Supplier
'
'CHANGE_ADD_TEmail_tlbLetterReports:
'Debug.Print time
'Debug.Print 2046
'   Conn1.Execute "ALTER TABLE tlbLetterReports ADD COLUMN TEmail TEXT(40);"
'Debug.Print 2046
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Extend the field size SageSuppAC on 29/02/2012 Supplier
''###############################################################################################################
'RESIZE_SageSuppAC_Supplier:
'
'   Rst1.Open "SELECT SageSuppAC FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields.Item("SageSuppAC").DefinedSize = 10 Then
'      Rst1.Close
'      Set Rst1 = Nothing
'
'Debug.Print time
'Debug.Print 2060
'      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SageSuppAC TEXT(50)"
'Debug.Print 2060
'   Else
'      Rst1.Close
'      Set Rst1 = Nothing
'   End If
'
''   Extend the field size of Field8 on 29/02/2012 ShoppingCentre
''###############################################################################################################
'   Rst1.Open "SELECT Field8 FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields.Item("Field8").DefinedSize = 50 Then
'      Rst1.Close
'      Set Rst1 = Nothing
'
'Debug.Print time
'Debug.Print 2074
'      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN Field8 TEXT(255)"
'Debug.Print 2074
'   Else
'      Rst1.Close
'      Set Rst1 = Nothing
'   End If
'
''   Add new record CT on 09/02/2012 Client
''###############################################################################################################
'ADD_CT_Client:
'   On Error GoTo MissingData_ADD_CT_Client
'
'   Rst1.Open "SELECT CT FROM Client;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_FileLoc__tlbClientBanks
''GoTo UPDATE_AMOUNT_2_DECIMAL
'
'MissingData_ADD_CT_Client:
'Debug.Print time
'Debug.Print 2092
'   Conn1.Execute "ALTER TABLE Client ADD COLUMN CT TEXT(25);"
'Debug.Print 2092
'Debug.Print time
'Debug.Print 2093
'   Conn1.Execute "UPDATE Client SET CT = 'Property Management';"
'Debug.Print 2093
'   UpdatePI_ClientID
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column FileLoc_ on 06/03/2012 tlbClientBanks
''###############################################################################################################
'ADD_FileLoc__tlbClientBanks:
'   On Error GoTo CHANGE_ADD_FileLoc__tlbClientBanks
'
'   Rst1.Open "SELECT FileLoc_ FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo tlbPaymentSplit
'
'CHANGE_ADD_FileLoc__tlbClientBanks:
'Debug.Print time
'Debug.Print 2109
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN FileLoc_ TEXT(255);"
'Debug.Print 2109
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 14/02/2012 tlbPaymentSplit
''###############################################################################################################
'tlbPaymentSplit:
'   On Error GoTo MissingTable_tlbPaymentSplit
'
'   Rst1.Open "SELECT * FROM tlbPaymentSplit;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount = 0 Then
''     We have to check here: is there any PP partially allocated?
''     If any PP found then user has to fix it before upgrade the system to PP split.
''     System will generate a report for user to fix PP.
'      Rst2.Open "SELECT * FROM tlbPayment " & _
'                "WHERE Type > 7 AND OSAmount > 0 AND " & _
'                      "OSAmount < Amount;", Conn1, adOpenStatic, adLockReadOnly
'      If Not Rst2.EOF Then
'         Rst2.Close
'         Rst1.Close
'         ShowReport App.Path & szReportPath & "\PP_PartialAlloc.rpt"
'         MsgBox "Please print this report and unallocate these transactions." & Chr(13) & _
'                "After upgrade the system, re-allocate these transactions.", _
'                vbInformation + vbOKOnly, "Partially Allocated Purchase Payment"
'
'         GoTo FixDataByUser
'      Else
'         Rst2.Close
'      End If
'
'      If Not UpdateTlbPaymentSplit Then
'         UpdateDatabase1 = -1
'         Exit Function
'      End If
'   End If
'
'   Rst1.Close
'
'   GoTo ADDNEW_REC_LL
'
'MissingTable_tlbPaymentSplit:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbPaymentSplit"
''Debug.Print ERR.description
'
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Add new record Landlord on 09/03/2012 SecondaryCode
''###############################################################################################################
'ADDNEW_REC_LL:
'   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'SCODE' AND Code = 'LL';", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'      With Rst1
'         .AddNew
'         !PrimaryCode = "SCODE"
'         !Code = "LL"
'         !Value = "Landlord"
'         .Update
'      End With
'   End If
'   Rst1.Close
'
''   Add new column RAS on 25/04/2012 Property
''###############################################################################################################
'ADDNEW_COL_RAS_Property:
'   On Error GoTo MISSING_ADDNEW_COL_RAS_Property
'
'   Rst1.Open "SELECT RAS FROM Property;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo FIX_LeaseRef_DemandRecords
'
'MISSING_ADDNEW_COL_RAS_Property:
'Debug.Print time
'Debug.Print 2187
'   Conn1.Execute "ALTER TABLE Property ADD COLUMN RAS TEXT(1);"
'Debug.Print 2187
'   UpdateDatabase1 = 1
'   Exit Function
'
''  Check DemandRecords table 'LeaseRef' column.
''  If 'LeaseRef' is empty then system will update according LeaseDetails table.
''  If the lease is expired then system will update 'LeaseRef' with latest expired lease
''###############################################################################################################
'FIX_LeaseRef_DemandRecords:
'   On Error GoTo MISSING_FIX_LeaseRef_DemandRecords
'
'   Rst1.Open "SELECT D.* " & _
'             "FROM DemandRecords AS D INNER JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'             "WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef));", Conn1, adOpenDynamic, adLockOptimistic
'
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo UPDATE_UNITNUMBER_RECEIPT
'   Else
'Debug.Print time
'Debug.Print 2206
'      Conn1.Execute "UPDATE DemandRecords AS D, " & _
'                        "[SELECT * FROM LeaseDetails WHERE Status = TRUE]. AS L, " & _
'                        "[SELECT D.LeaseRef, D.DemandID " & _
'                        " FROM DemandRecords AS D INNER JOIN " & _
'                        "     DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
'                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
'                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
'                    "WHERE D.DemandID = SQ.DemandID AND L.Status AND " & _
'                        "L.SageAccountNumber = D.SageAccountNumber;"
'
'Debug.Print time
'Debug.Print 2217
'      Conn1.Execute "UPDATE DemandRecords AS D, LeaseDetails AS L, " & _
'                        "[SELECT D.LeaseRef, D.DemandID " & _
'                        " FROM DemandRecords AS D INNER JOIN " & _
'                              "DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
'                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
'                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
'                    "WHERE D.DemandID = SQ.DemandID AND " & _
'                        "L.SageAccountNumber = D.SageAccountNumber;"
'   End If
'
'   Rst1.Close
'   GoTo UPDATE_UNITNUMBER_RECEIPT
'
'MISSING_FIX_LeaseRef_DemandRecords:
'   Debug.Print Err.description
'   MsgBox "This company database needs to be updated. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Lease Reference in Demand"
'
''###############################################################################################################
''  Update the UnitNumber in the transactions for expired lessee
''  Previously we used NOC as unit number if user creates SI after expired the lease
''  Now, system will update NOC with a lastest unit number, which will be found in the lease details
''     table where latest expired lease
'UPDATE_UNITNUMBER_RECEIPT:
'   On Error GoTo ERROR_UPDATE_UNITNUMBER_RECEIPT
'
'   Rst1.Open "SELECT * " & _
'             "FROM tlbReceipt AS R " & _
'             "WHERE R.UnitID = 'NOC' OR ISNULL(R.UnitID);", Conn1, adOpenDynamic, adLockOptimistic
'
'   While Not Rst1.EOF
'      Rst2.Open "SELECT L.UnitNumber, L.Status " & _
'                "FROM  LeaseDetails AS L " & _
'                "WHERE L.SageAccountNumber = '" & Rst1.Fields.Item("SageAccountNumber").Value & "' " & _
'                "ORDER BY L.Status, L.EndDate DESC;", Conn1, adOpenStatic, adLockReadOnly
''Debug.Print "SELECT L.UnitNumber, L.Status " & _
'                "FROM  LeaseDetails AS L " & _
'                "WHERE L.SageAccountNumber = '" & Rst1.Fields.Item("SageAccountNumber").Value & "' " & _
'                "ORDER BY L.Status, L.EndDate DESC;"
'      If Not Rst2.EOF Then
'         Rst1.Fields.Item("UnitID").Value = Rst2.Fields.Item(0).Value
'         Rst1.Update
'      End If
'
'      Rst1.MoveNext
'      Rst2.Close
'   Wend
'   Rst1.Close
'
'   GoTo UPDATE_AMOUNT_2_DECIMAL
'
'ERROR_UPDATE_UNITNUMBER_RECEIPT:
'   MsgBox "Unit information could not be updated automatically", vbCritical + vbOKOnly, "Unit Id for expired lessee in the receipt."
'   UpdateDatabase1 = 1
'   Exit Function
'
''###############################################################################################################
''  THE FOLLOWING UPDATE SQL WILL FIX AMOUNT FIGURE INTO 2 DECIMAL PLACE.
''  THIS UPDATE PROCESS WILL RUN IN MY PC. I WILL UPDATE ALL DATABASE IN THE PCM
''  WHEN I WILL CONFIRM ALL TRANSACTION SAVING PROCESS TO 2 DECIMAL PROCESS,
''  THIS UPDATE ROCESS WILL NOT BE REQUIRED ANY MORE.
'UPDATE_AMOUNT_2_DECIMAL:
'   On Error GoTo ERROR_UPDATE_AMOUNT_2_DECIMAL
'
'   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
'GoTo RESIZE_Email1_Tenants
'      Update2DecimalPlace "DemandSplitRecords", "DSR", "Amount"
'      Update2DecimalPlace "DemandSplitRecords", "DSR", "TotalAmount"
'      Update2DecimalPlace "DemandSplitRecords", "DSR", "VATAmount"
'
'      Update2DecimalPlace "tlbReceipt", "TransactionID", "Amount"
'      Update2DecimalPlace "tlbReceipt", "TransactionID", "OSAmount"
'      Update2DecimalPlace "tlbReceipt", "TransactionID", "ReceiptAmount"
'
'      Update2DecimalPlace "tlbReceiptSplit", "TransactionID", "Amount"
'      Update2DecimalPlace "tlbReceiptSplit", "TransactionID", "OSAmount"
'
'      Update2DecimalPlace "tlbPayment", "TransactionID", "Amount"
'      Update2DecimalPlace "tlbPayment", "TransactionID", "OSAmount"
'
'      Update2DecimalPlace "tlbPaymentSplit", "TransactionID", "Amount"
'      Update2DecimalPlace "tlbPaymentSplit", "TransactionID", "OSAmount"
'
'      Update2DecimalPlace "PayTransactions", "TransactionID", "PaymentAmount"
'      Update2DecimalPlace "RptTransactions", "TransactionID", "ReceiptAmount"
'
'      Update2DecimalPlace "tblPurInv", "MY_ID", "TOTAL_AMOUNT"
'
'      Update2DecimalPlace "tblPurInvSRec", "MY_ID", "NET_AMOUNT"
'      Update2DecimalPlace "tblPurInvSRec", "MY_ID", "TOTAL_AMOUNT"
'
'      Update2DecimalPlace "LServiceCharges", "ServiceCharge", "SCTotal"
'      Update2DecimalPlace "LServiceCharges", "ServiceCharge", "SCAmount"
'      Update2DecimalPlace "LRentCharges", "RentCharges", "BRTotal"
'      Update2DecimalPlace "LRentCharges", "RentCharges", "BRAmount"
'      Update2DecimalPlace "LInsuranceCharges", "InsCharges", "InsuranceEachPeriod"
'      Update2DecimalPlace "LInsuranceCharges", "InsCharges", "TotalYearlyInsurance"
''
''AFTER RUNNING THIS SCRIPT, CHECK BY: SELECT R.TransactionID, R.Amount, R.DemandRef, S.DT FROM tlbReceipt AS R, (SELECT D.DemandID,  SUM(S.TotalAmount) AS DT FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID GROUP BY D.DemandID) AS S WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND ROUND(R.Amount, 2) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 2);
''IF THIS SQL GETS ANY RECORDS, THEN FIX THE DEMANDSPLIT TABLE
''
'   End If
'
'   GoTo RESIZE_Email1_Tenants
'
'ERROR_UPDATE_AMOUNT_2_DECIMAL:
'   MsgBox "System could not update all transactions to two decimal places", vbCritical + vbOKOnly, "Update Transactions"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Extend the field size Email on 12/06/2012 Tenants
''###############################################################################################################
'RESIZE_Email1_Tenants:
'
'   On Error GoTo MissingTable_RESIZE_Email1_Tenants
'
'   Rst1.Open "SELECT Email1 FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields.Item("Email1").DefinedSize = 40 Or Rst1.Fields.Item("Email1").DefinedSize = 99 Then
'      Rst1.Close
'      Set Rst1 = Nothing
'
'Debug.Print time
'Debug.Print 2339
'      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN Email1 TEXT(100)"
'Debug.Print 2339
'Debug.Print time
'Debug.Print 2340
'      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN Email2 TEXT(100)"
'Debug.Print 2340
'Debug.Print time
'Debug.Print 2341
'      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientOfficeEmail TEXT(100)"
'Debug.Print 2341
'Debug.Print time
'Debug.Print 2342
'      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientPersonalEmail TEXT(100)"
'Debug.Print 2342
'Debug.Print time
'Debug.Print 2343
'      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentOfficeEmail TEXT(100)"
'Debug.Print 2343
'Debug.Print time
'Debug.Print 2344
'      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentPersonalEmail TEXT(100)"
'Debug.Print 2344
'Debug.Print time
'Debug.Print 2345
'      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeEmail TEXT(100)"
'Debug.Print 2345
'Debug.Print time
'Debug.Print 2346
'      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierPersonalEmail TEXT(100)"
'Debug.Print 2346
'Debug.Print time
'Debug.Print 2347
'      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SageSuppAC TEXT(100)"
'Debug.Print 2347
'   End If
'   Rst1.Close
'   Set Rst1 = Nothing
'
''   Exit Function
'   GoTo BUGFIX_DemandSplitRecords_SplitID
'
'MissingTable_RESIZE_Email1_Tenants:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - Email1 of DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''BUGFIX: DEMAND SPLIT ID
''        When users delete the first line of demand record split which SplitID=1
''        System does not reschedule the split id
''        This creates the problem in the receipt unallocation
''        To solve, ..
'BUGFIX_DemandSplitRecords_SplitID:
'   Rst1.Open "SELECT Q1.* " & _
'             "FROM ( " & _
'               "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
'               "FROM DemandSplitRecords AS DS " & _
'               "GROUP BY DS.DemandID " & _
'             ") AS Q1 INNER JOIN " & _
'             "( " & _
'               "SELECT DS.DemandID, COUNT(DSR) AS C " & _
'               "FROM DemandSplitRecords AS DS " & _
'               "GROUP BY DS.DemandID " & _
'             ") AS Q2 ON Q2.DemandID = Q1.DemandID " & _
'             "WHERE Q1.M <> Q2.C", Conn1, adOpenStatic, adLockReadOnly
'   If Not Rst1.EOF Then                                                 'There are demand split without SplitID
'      Rst1.Close
'Debug.Print time
'Debug.Print 2380
'      Conn1.Execute "UPDATE DemandSplitRecords " & _
'                    "Set SplitID = SplitID - 1 " & _
'                    "WHERE DSR IN ( " & _
'                    "SELECT DS.DSR " & _
'                    "FROM ((" & _
'                       "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
'                       "FROM DemandSplitRecords AS DS " & _
'                       "GROUP BY DS.DemandID " & _
'                    ") AS Q1 INNER JOIN " & _
'                    "( " & _
'                       "SELECT DS.DemandID, COUNT(DSR) AS C " & _
'                       "FROM DemandSplitRecords AS DS " & _
'                       "GROUP BY DS.DemandID " & _
'                    ") AS Q2 ON Q2.DemandID = Q1.DemandID) INNER JOIN " & _
'                       "DemandSplitRecords AS DS ON Q1.DemandID = DS.DemandID " & _
'                    "WHERE Q1.M <> Q2.C);"
'
'Debug.Print time
'Debug.Print 2397
'         Conn1.Execute "UPDATE tlbReceiptSplit AS R, DemandSplitRecords AS S " & _
'                       "SET R.SplitID = S.SplitID " & _
'                       "WHERE R.AllocTranID = S.DSR AND R.Amount = S.TotalAmount;"
'
'         UpdateDatabase1 = 1
'
'         Exit Function
'   End If
'   Rst1.Close
'
''   Add new column SelFund on 01/08/2012 Fund
''###############################################################################################################
'ADD_SelFund_Fund:
'   On Error GoTo CHANGE_ADD_SelFund_Fund
'
'   Rst1.Open "SELECT SelFund FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_szFundID_Fund
'
'CHANGE_ADD_SelFund_Fund:
'Debug.Print time
'Debug.Print 2418
'   Conn1.Execute "ALTER TABLE Fund ADD COLUMN SelFund TEXT(10);"
'Debug.Print 2418
'Debug.Print time
'Debug.Print 2419
'   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN SelBanks TEXT(10);"
'Debug.Print 2419
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column szFundID on 10/08/2012 Fund
''###############################################################################################################
'ADD_szFundID_Fund:
'
'   On Error GoTo CHANGE_ADD_szFundID_Fund
'
'   Rst1.Open "SELECT szFundID FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo CREAT_SPLIT_tblPurInv
'
'CHANGE_ADD_szFundID_Fund:
'Debug.Print time
'Debug.Print 2435
'   Conn1.Execute "ALTER TABLE Fund ADD COLUMN szFundID TEXT(10);"
'Debug.Print 2435
'Debug.Print time
'Debug.Print 2436
'   Conn1.Execute "UPDATE Fund SET szFundID = CSTR(FundID);"
'Debug.Print 2436
'Debug.Print time
'Debug.Print 2437
'   Conn1.Execute "ALTER TABLE Fund ADD COLUMN FundList TEXT(255);"
'Debug.Print 2437
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Check the system PI Transactions which have not been exported to payment table.
''   If there is any transaction found, then system will export them
''###############################################################################################################
'CREAT_SPLIT_tblPurInv:
'' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
''   szSQL = "SELECT * " & _
''           "FROM   tblPurInv " & _
''           "WHERE MY_ID NOT IN (SELECT ParentID FROM tblPurInvSRec GROUP BY ParentID);"
'    szSQL = "SELECT A.* FROM   tblPurInv A Left JOIN tblPurInvSRec B ON A.MY_ID=B.ParentID WHERE A.MY_ID<>B.ParentID;"
''Debug.Print szSql
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      szSQL = "SELECT * FROM tblPurInvSRec"
'      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
'
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("MY_ID").Value = UniqueID()
'            .Fields.Item("ParentID").Value = Rst1.Fields.Item("MY_ID").Value
'            .Fields.Item("TRAN_ID").Value = 1
'            .Fields.Item("TRANS").Value = Rst1.Fields.Item("PropertyID").Value
'            .Fields.Item("NOMINAL_CODE").Value = "0000"
'            .Fields.Item("DEPT_ID").Value = 0
'            .Fields.Item("description").Value = "DELETED PURCHASE TRANSACTIONS"
'            .Fields.Item("NET_AMOUNT").Value = 0
'            .Fields.Item("TAX_CODE").Value = "T9"
'            .Fields.Item("VAT").Value = 0
'            .Fields.Item("TOTAL_AMOUNT").Value = 0
'            .Fields.Item("RecoverablePt").Value = 0
'
'            .Update
'         End With
'
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close
'' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
''   szSQL = "SELECT * " & _
''           "From tlbPayment " & _
''           "WHERE TransactionID NOT IN (" & _
''               "SELECT PayHeader FROM tlbPaymentSplit GROUP BY PayHeader);"
' szSQL = "SELECT A.* From tlbPayment A LEFT JOIN tlbPaymentSplit B ON A.TransactionID= B.PayHeader WHERE  B.PayHeader is NULL;"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''Debug.Print szSql
'
'   If Not Rst1.EOF Then
'      szSQL = "SELECT * FROM tlbPaymentSplit;"
'      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
'
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("TransactionID").Value = UniqueID()
'            .Fields.Item("PayHeader").Value = Rst1.Fields.Item("TransactionID").Value
'            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
'            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
'            .Fields.Item("OSAmount").Value = Rst1.Fields.Item("OSAmount").Value
''Debug.Print Rst1.Fields.Item("Amount").Value
'            .Fields.Item("SplitID").Value = 1
'            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
'            .Fields.Item("Description").Value = "DELETED PURCHASE TRANSACTIONS"
'            .Update
'         End With
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close
''********************************************************************************************
''   There are some PI found in the header table, which amounts don't match with split total
''********************************************************************************************
'   szSQL = "SELECT P.*, S.ST " & _
'           "FROM tblPurInv AS P, ( " & _
'               "SELECT ParentID, SUM(TOTAL_AMOUNT) AS ST " & _
'               "FROM tblPurInvSRec " & _
'               "GROUP BY ParentID) AS S " & _
'           "WHERE P.MY_ID = S.ParentID And P.TOTAL_AMOUNT <> S.ST"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      While Not Rst1.EOF
'         If Val(Rst1.Fields.Item("TOTAL_AMOUNT").Value) = 0 Then
'Debug.Print time
'Debug.Print 2527
'            Conn1.Execute "UPDATE tblPurInvSRec AS S " & _
'                          "SET    NET_AMOUNT = 0, VAT = 0, TOTAL_AMOUNT = 0 " & _
'                          "WHERE  S.ParentID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         Else
'Debug.Print time
'Debug.Print 2531
'            Conn1.Execute "UPDATE tblPurInv AS P " & _
'                          "SET    P.TOTAL_AMOUNT = " & Rst1.Fields.Item("ST").Value & " " & _
'                          "WHERE  P.MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         End If
'
'         Rst1.MoveNext
'      Wend
'   End If
'   Rst1.Close
'
''   New table on 03/09/2012 tlbBankReconcilation
''###############################################################################################################
'NEW_TABLE_TlbBankReconcilation:
'
'   On Error GoTo MissingTable_tlbBankReconcilation
'
'   Rst1.Open "SELECT * FROM tlbBankReconcilation;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   Rst1.Open "SELECT * FROM tlbBankReconcilation WHERE MY_ID = 'SAMRAT';", Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 2553
'      Conn1.Execute "DELETE * FROM tlbBankReconcilation;"
'Debug.Print 2553
'      CreateBankReconSplits
'   End If
'   Rst1.Close
'
'   GoTo NEW_TABLE_tlbBankReconClosingBal
'
'MissingTable_tlbBankReconcilation:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbBankReconcilation"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   New table on 03/09/2012 tlbBankReconClosingBal
''###############################################################################################################
'NEW_TABLE_tlbBankReconClosingBal:
'
'   On Error GoTo MissingTable_tlbBankReconClosingBal
'
'   Rst1.Open "SELECT * FROM tlbBankReconClosingBal;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'   GoTo ADD_szTransactionID_tlbReceipt
'
'MissingTable_tlbBankReconClosingBal:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbBankReconClosingBal"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   Add new column szTransactionID on 11/09/2012 tlbReceipt
''###############################################################################################################
'ADD_szTransactionID_tlbReceipt:
'   On Error GoTo CHANGE_ADD_szTransactionID_tlbReceipt
'
'   Rst1.Open "SELECT szTransactionID FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_szTransactionID_tlbPayment
'
'CHANGE_ADD_szTransactionID_tlbReceipt:
'Debug.Print time
'Debug.Print 2592
'   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN szTransactionID TEXT(20);"
'Debug.Print 2592
'Debug.Print time
'Debug.Print 2593
'   Conn1.Execute "UPDATE tlbReceipt SET szTransactionID = CSTR(TransactionID)"
'Debug.Print 2593
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column szTransactionID on 11/09/2012 tlbPayment
''###############################################################################################################
'ADD_szTransactionID_tlbPayment:
'   On Error GoTo CHANGE_ADD_szTransactionID_tlbPayment
'
'   Rst1.Open "SELECT szTransactionID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   'GoTo ADD_RRR_Fund
'   GoTo FIXING_SplitID_tlbPaymentSplit
'
'CHANGE_ADD_szTransactionID_tlbPayment:
'Debug.Print time
'Debug.Print 2610
'   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN szTransactionID TEXT(20);"
'Debug.Print 2610
'Debug.Print time
'Debug.Print 2611
'   Conn1.Execute "UPDATE tlbPayment SET szTransactionID = CSTR(TransactionID)"
'Debug.Print 2611
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a column RRR on 30/10/2012 Fund
''###############################################################################################################
''ADD_RRR_Fund:
''   On Error GoTo CHANGE_ADD_RRR_Fund
''
''   Rst1.Open "SELECT RRR FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo FIXING_SplitID_tlbPaymentSplit
''
''CHANGE_ADD_RRR_Fund:
'Debug.Print time
'Debug.Print 2626
''   Conn1.Execute "ALTER TABLE Fund ADD COLUMN RRR TEXT(10);"
'Debug.Print 2626
''   UpdateDatabase1 = 1
''   Exit Function
'
''********************************************************************************************
''   There are some PI found which FundID don't match with split FundID
''********************************************************************************************
'FIXING_SplitID_tlbPaymentSplit:
'   szSQL = "SELECT A.PayHeader, A.X " & _
'           "FROM ( " & _
'                 "SELECT PayHeader, SplitID, COUNT(SplitID) AS X " & _
'                 "From tlbPaymentSplit " & _
'                 "GROUP BY PayHeader, SplitID " & _
'           ") AS A " & _
'           "WHERE A.X > 1;"
''
''            Debug.Print szSQL
''             Debug.Print time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''   Debug.Print time
'   If Not Rst1.EOF Then
'      'keep a log in the error log table
'Debug.Print time
'Debug.Print 2648
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbPaymentSplit SplitID duplicated " & Rst1.Fields("PayHeader").Value & "' )"
'Debug.Print 2648
'      While Not Rst1.EOF
'         Rst2.Open "SELECT * FROM tlbPaymentSplit AS S " & _
'                   "WHERE S.PayHeader = " & Rst1.Fields.Item(0).Value & ";", Conn1, adOpenDynamic, adLockOptimistic
'
'         For i = 1 To RecordCount(Rst2) 'Rst1.Fields.Item("X").Value 'fixed by anol 20181119
'            Rst2.Fields.Item("SplitID").Value = i
'            Rst2.Update
'            Rst2.MoveNext
'         Next i
'
'         Rst1.MoveNext
'         Rst2.Close
'      Wend
'   End If
'   Rst1.Close
''   Debug.Print time
'
''********************************************************************************************
''   FIXING DATA: There are some PI found which FundID don't match with split FundID
''********************************************************************************************
'''FIXING_FUNDID_tlbPayment_tlbPaymentSplit:
'''   szSQL = "SELECT P.* " & _
'''           "FROM tlbPaymentSplit AS S, tlbPayment AS P " & _
'''           "WHERE P.TransactionID = S.PayHeader AND " & _
'''                 "P.FundID <> S.FundID AND S.Splitid = 1;"
'''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'''
'''   If Not Rst1.EOF Then
'''   'we are not doing this anymore by anol 20181119
'Debug.Print time
'Debug.Print 2678
''''      Conn1.Execute "UPDATE tlbPayment AS P, tlbPaymentSplit AS S " & _
'Debug.print 2678
''''                    "SET    P.FundID = S.FundID " & _
''''                    "WHERE  P.TransactionID = S.PayHeader AND " & _
''''                           "S.Splitid = 1;"
'''   End If
'''   Rst1.Close
'
''********************************************************************************************
''   FIXING DATA: tlbReceiptSplit table has some transaction with FundID = 0
''********************************************************************************************
'FIXING_FUNDID_tlbReceiptSplit:
'   szSQL = "SELECT S.* " & _
'           "FROM tlbReceiptSplit AS S " & _
'           "WHERE S.FundID = 0;"
'   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
'
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 2695
'      Conn1.Execute "UPDATE tlbReceiptSplit AS S1, tlbReceiptSplit AS S2, " & _
'                        "tlbReceipt AS R1,tlbReceipt AS R2, RptTransactions AS T " & _
'                    "Set S1.FundID = S2.FundID, S1.SplitID = S2.SplitID " & _
'                    "WHERE S1.FundID = 0 AND R1.TransactionID = T.FromTran AND " & _
'                        "T.ToTran = R2.TransactionID AND R1.TransactionID = S1.RptHeader AND " & _
'                        "R2.TransactionID = S2.RptHeader  AND VAL(S1.AllocTranID) = S2.RptHeader;"
'   End If
'   Rst1.Close
'
''********************************************************************************************
''   FIXING DATA: tlbPayment, tlbReceipt & tlbBankReconcilation tables have some transactions
''                which have amount transaction method is 1 (CHQ, DD, etc)
''********************************************************************************************
'FIXING_AMT_MTH_tlbPayment:
'   szSQL = "UPDATE tlbPayment " & _
'           "SET PayAmtType = 'CHQ' " & _
'           "WHERE PayAmtType = '1';"
'Debug.Print time
'Debug.Print 2712
'   Conn1.Execute szSQL
'Debug.Print 2712
'
'FIXING_AMT_MTH_tlbReceipt:
'   szSQL = "UPDATE tlbReceipt " & _
'           "SET RptAmtType = 'CHQ' " & _
'           "WHERE RptAmtType = '1';"
'Debug.Print time
'Debug.Print 2718
'   Conn1.Execute szSQL
'Debug.Print 2718
'
'FIXING_AMT_MTH_tlbBankReconcilation:
'   szSQL = "UPDATE tlbBankReconcilation " & _
'           "SET TranMth = 'CHQ' " & _
'           "WHERE TranMth = '1';"
'Debug.Print time
'Debug.Print 2724
'   Conn1.Execute szSQL
'Debug.Print 2724
'
''********************************************************************************************
''   FIXING DATA: There are some reconciled transactions found dated after their
''                reconciled statement date. which tran dt is later then the bank reconcilation statement date
''********************************************************************************************
''Below procedure is rem out by anol 27 07 2016
'''FIXING_TRANS_DT_tlbPayment:
'''   Dim sChoice As Single
'''
'''   szSQL = "SELECT * " & _
'''           "FROM  tlbBankReconcilation " & _
'''           "WHERE TDate > ReconDate "
'''
'''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'''   If Not Rst1.EOF Then
'''      sChoice = MsgBox("There are some reconciled transactions found dated after their reconciled statement date." + Chr(13) + _
'''                    "Please contact PCM Support to correct these transactions before proceeding further." + Chr(13) + _
'''                    "Click OK to print a list of these transactions.", vbCritical + vbOKCancel, _
'''                    "Purchase Payment")
'''      If sChoice = vbOK Then
'''         Rst1.Close
'''         ShowReport App.Path & szReportPath & "\TranDt_BankRecDate.rpt"
'''         UpdateDatabase1 = 0
'''         Exit Function
'''      End If
'''   End If
'''   Rst1.Close
'   'end of rem
''Update tlbBankReconcilation
''SET TDate = ReconDate, DDate = ReconDate
''Where TDate > ReconDate And (TransactionType = 11 OR TransactionType = 12 OR TransactionType = 8 OR TransactionType = 9 OR TransactionType = 3 OR TransactionType = 4)
''--------------------------------------------------------------------
''Update tlbBankPayment
''Set TRAN_DATE = CDate(Left(ReconNow, 10))
''Where TRAN_DATE > CDate(Left(ReconNow, 10)) And (TransactionType = 11 OR TransactionType = 12) and ReconNow <> "" and not isnull(ReconNow)
''--------------------------------------------------------------------
''Update tlbPayment
''Set PDATE = CDate(Left(ReconNow, 10))
''WHERE PDATE > CDATE(LEFT(ReconNow, 10))  and (Type = 8 OR Type = 9) and ReconNow <> "" and not isnull(ReconNow)
''--------------------------------------------------------------------
''Update tlbReceipt
''SET RDATE = CDATE(LEFT(ReconNow, 10)), DDATE = CDATE(LEFT(ReconNow, 10))
''WHERE RDATE > CDATE(LEFT(ReconNow, 10))  and (Type = 3 or Type = 4) and ReconNow <> "" and not isnull(ReconNow)
'
'
'
'
''********************************************************************************************
''   FIXING DATA: There are some SRR found in the receipt and split table, their split OS <> header OS
''********************************************************************************************
'FIXING_SRR_OS_Split:
'   szSQL = "UPDATE  tlbReceiptSplit AS S, tlbReceipt AS R " & _
'           "SET     R.OSAmount = S.OSAmount " & _
'           "WHERE   R.TransactionID = S.RptHeader AND " & _
'                   "R.Type = 23 AND " & _
'                   "R.OSAmount <> S.OSAmount;"
'Debug.Print time
'Debug.Print 2781
'   Conn1.Execute szSQL
'Debug.Print 2781
'
'   szSQL = "UPDATE  tlbReceiptSplit AS S " & _
'           "SET     S.SplitID = 1 " & _
'           "WHERE   S.SplitID = -1;"
'Debug.Print time
'Debug.Print 2786
'   Conn1.Execute szSQL
'Debug.Print 2786
'
''********************************************************************************************
''   FIXING DATA: There are some SA found in the receipt table without split in the split table
''********************************************************************************************
'' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
''   szSQL = "SELECT * " & _
''           "FROM  tlbReceipt " & _
''           "WHERE Type = 4 AND Amount <> 0 AND " & _
''               "TransactionID NOT IN (" & _
''                  "SELECT RptHeader " & _
''                  "FROM   tlbReceiptSplit " & _
''                  "GROUP BY RptHeader);"
'   szSQL = "SELECT A.* FROM  tlbReceipt A Left join tlbReceiptSplit B ON A.TransactionID =B.RptHeader WHERE B.RptHeader IS NULL AND Type = 4 AND A.Amount <> 0 ;"
''Debug.Print szSQL
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      szSQL = "SELECT * FROM tlbReceiptSplit;"
'      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
'
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("TransactionID").Value = UniqueID()
'            .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
'            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
'            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
'            .Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
'            .Fields.Item("SplitID").Value = 1
'            .Fields.Item("DueDate").Value = Format(Rst1.Fields.Item("DDate").Value, "dd mmmm yyyy")
'            .Fields.Item("Description").Value = "Receipt on Account"
'            .Update
'         End With
'
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'
'   Rst1.Close
'
''********************************************************************************************
''   FIXING DATA:
''********************************************************************************************
'FIXING_HEADER_n_SPLIT_SI_n_PI:
'   Call Pi_Check_pre(Conn1) 'written by anol issue 791 Batch Payments (Support - WPM) 2019-07-25
'
'   Call SiPi_Check(Conn1, "PI", "25875")
'   Call SiPi_Check(Conn1, "SI", "15875")
'
''How to Fix:
''     Look at the total of tlbPayment header and split amount total
''     With the help of PayTransaction table, determine the split's amount
'
''********************************************************************************************
''   There are some PP found in the header table, which O/S amounts don't match with split O/S total
''********************************************************************************************
''   szSQL = "SELECT P.TransactionID, S.ST " & _
''           "FROM tlbPayment AS P, ( " & _
''               "SELECT PayHeader, SUM(OSAmount) AS ST " & _
''               "FROM tlbPaymentSplit " & _
''               "GROUP BY PayHeader) AS S " & _
''           "WHERE P.TransactionID = S.PayHeader And P.OSAmount <> S.ST"
''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''
''   If Not Rst1.EOF Then
'''      szSQL = SQL2String(Rst1, 0)
'''Debug.Print szSQL
'''      MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
''      While Not Rst1.EOF
''         szSQL = "UPDATE tlbPayment " & _
''                 "SET OSAmount = " & Rst1.Fields.Item(1).Value & " " & _
''                 "WHERE TransactionID = " & Rst1.Fields.Item(0).Value & ";"
'''Debug.Print szSQL
'Debug.Print time
'Debug.Print 2861
''         Conn1.Execute szSQL
'Debug.Print 2861
''         Rst1.MoveNext
''      Wend
'
''      While Not Rst1.EOF
''         If Val(Rst1.Fields.Item("TOTAL_AMOUNT").Value) = 0 Then
'Debug.Print time
'Debug.Print 2867
''            Conn1.Execute "UPDATE tblPurInvSRec AS S " & _
'Debug.print 2867
''                          "SET    NET_AMOUNT = 0, VAT = 0, TOTAL_AMOUNT = 0 " & _
''                          "WHERE  S.ParentID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
''         Else
'Debug.Print time
'Debug.Print 2871
''            Conn1.Execute "UPDATE tblPurInv AS P " & _
'Debug.print 2871
''                          "SET    P.TOTAL_AMOUNT = " & Rst1.Fields.Item("ST").Value & " " & _
''                          "WHERE  P.MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
''         End If
''
''         Rst1.MoveNext
''      Wend
''   End If
''   Rst1.Close
'
''   Add a columns RRTotal on 08/08/2011 LRentCharges
''###############################################################################################################
'ADD_RRTotal_LRentCharges:
'   On Error GoTo CHANGE_ADD_RRTotal_LRentCharges
'
'   Rst1.Open "SELECT RRTotal FROM LRentCharges;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo LeaseHistory
'
'CHANGE_ADD_RRTotal_LRentCharges:
'Debug.Print time
'Debug.Print 2893
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRTotal DOUBLE;"
'Debug.Print 2893
'Debug.Print time
'Debug.Print 2894
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRAmount DOUBLE;"
'Debug.Print 2894
'Debug.Print time
'Debug.Print 2895
'   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRPrint TEXT(1);"
'Debug.Print 2895
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 10/12/2012 LeaseHistory
''###############################################################################################################
'LeaseHistory:
'   On Error GoTo MissingTable_LeaseHistory
'
'   Rst1.Open "SELECT * FROM LeaseHistory;", Conn1, adOpenStatic, adLockReadOnly
'   lSlNumber = Rst1.RecordCount
'   Rst1.Close
'
'   If lSlNumber = 0 Then
'      Rst1.Open "SELECT * FROM LeaseHistory;", Conn1, adOpenDynamic, adLockPessimistic
'      Rst2.Open "SELECT * FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly
'
'      While Not Rst2.EOF
'         Rst1.AddNew
'         Rst1.Fields.Item("HistoryID").Value = UniqueID()
'         For i = 1 To Rst1.Fields.count - 1
'            For lSlNumber = 0 To Rst2.Fields.count - 1
'               If Rst1.Fields.Item(i).Name = Rst2.Fields.Item(lSlNumber).Name Then
'                  Rst1.Fields.Item(i).Value = Rst2.Fields.Item(lSlNumber).Value
'                  Exit For
'               End If
'            Next lSlNumber
'         Next i
'         Rst1.Update
'         Rst2.MoveNext
'      Wend
'
'      Rst1.Close
'      Rst2.Close
'   End If
'
'   GoTo MODIFY_Code_FREQ_SecondaryCode
'
'MissingTable_LeaseHistory:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - LeaseHistory"
'   UpdateDatabase1 = -1
'   Exit Function
'
'
''   Amend DATA Code FREQ on 30/01/13 SecondaryCode
''###############################################################################################################
'MODIFY_Code_FREQ_SecondaryCode:
'   On Error GoTo CHANGE_Code_FREQ_SecondaryCode
'
'   Rst1.Open "SELECT Code FROM SecondaryCode WHERE Code = 'DAILY' AND PrimaryCode = 'FREQ';", Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'Debug.Print time
'Debug.Print 2947
'      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Code = 'QTR', SC.Value = 'QUARTERLY', SC.Description = '4' WHERE SC.Code = 'DAILY' AND SC.PrimaryCode = 'FREQ';"
'Debug.Print 2947
'Debug.Print time
'Debug.Print 2948
'      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Code = 'HY', SC.Value = 'HALF YEARLY', SC.Description = '2' WHERE SC.Code = 'MTHADV' AND SC.PrimaryCode = 'FREQ';"
'Debug.Print 2948
'Debug.Print time
'Debug.Print 2949
'      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Description = '12' WHERE SC.Code = 'MONTHLY' AND SC.PrimaryCode = 'FREQ';"
'Debug.Print 2949
'Debug.Print time
'Debug.Print 2950
'      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Description = '52' WHERE SC.Code = 'WEEKLY' AND SC.PrimaryCode = 'FREQ';"
'Debug.Print 2950
'   End If
'   Rst1.Close
'
'   GoTo NEW_TABLE_FinancialYear
'
'CHANGE_Code_FREQ_SecondaryCode:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 30/01/13 FinancialYear
''###############################################################################################################
'NEW_TABLE_FinancialYear:
'   On Error GoTo MissingTable_FinancialYear
'
'   szEmail = "FinancialYear"
'   Rst1.Open "SELECT * FROM " & szEmail & ";", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'   szEmail = "Periods"
'   Rst1.Open "SELECT * FROM " & szEmail & ";", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_LChildsRef_DemandSplPreview
'
'MissingTable_FinancialYear:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - " & szEmail
'   UpdateDatabase1 = -1
'   Exit Function
'
'FixDataByUser:
''System will jump here without modifying further. System wants user to do some job.
''Check the caller point for further information.
'   UpdateDatabase1 = 0
'
''   Add a columns LChildsRef on 08/03/2013 DemandSplPreview
''###############################################################################################################
'ADD_LChildsRef_DemandSplPreview:
'   On Error GoTo CHANGE_ADD_LChildsRef_DemandSplPreview
'
'   Rst1.Open "SELECT LChildsRef FROM DemandSplPreview;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_FinancialYear_GlobalSC
'
'CHANGE_ADD_LChildsRef_DemandSplPreview:
'Debug.Print time
'Debug.Print 2997
'   Conn1.Execute "ALTER TABLE DemandSplPreview ADD COLUMN LChildsRef TEXT(25);"
'Debug.Print 2997
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add a columns FinancialYear on 19/04/2013 GlobalSC
''###############################################################################################################
'ADD_FinancialYear_GlobalSC:
'   On Error GoTo CHANGE_ADD_FinancialYear_GlobalSC
'
'   Rst1.Open "SELECT FinancialYear FROM GlobalSC;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_COL_CBY_Property
'
'CHANGE_ADD_FinancialYear_GlobalSC:
'Debug.Print time
'Debug.Print 3013
'   Conn1.Execute "ALTER TABLE GlobalSC ADD COLUMN FinancialYear TEXT(20);"
'Debug.Print 3013
'Debug.Print time
'Debug.Print 3014
'   Conn1.Execute "ALTER TABLE GlobalInsurance ADD COLUMN FinancialYear TEXT(20);"
'Debug.Print 3014
'Debug.Print time
'Debug.Print 3015
'   Conn1.Execute "ALTER TABLE GlobalRC ADD COLUMN FinancialYear TEXT(20);"
'Debug.Print 3015
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column CBY on 22/04/2013 Property
''###############################################################################################################
'ADDNEW_COL_CBY_Property:
'   On Error GoTo MISSING_ADDNEW_COL_CBY_Property
'
'   Rst1.Open "SELECT CBY FROM Property;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADDNEW_REC_DTKW
'
'MISSING_ADDNEW_COL_CBY_Property:
'Debug.Print time
'Debug.Print 3031
'   Conn1.Execute "ALTER TABLE Property ADD COLUMN CBY TEXT(20);"
'Debug.Print 3031
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new record EMAIL_DEMAND_TEMPLATE on 30/04/2013 PrimaryCode
''###############################################################################################################
'ADDNEW_REC_DTKW:
'   With Rst1
'      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'DTKW';", Conn1, adOpenStatic, adLockReadOnly
'
'      If .EOF Then
'         .Close
'         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         !Code = "DTKW"
'         !Value = "DEMAND TEMPLATE KEYWORD"
'         !Flexible = False
'         .Update
'         .Close
'         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item(0).Value = "DTKW"
'         .Fields.Item(1).Value = "CN"
'         .Fields.Item(2).Value = "CLIENT NAME"
'         .Fields.Item(3).Value = "<CLIENT NAME>"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "DTKW"
'         .Fields.Item(1).Value = "LN"
'         .Fields.Item(2).Value = "LESSEE NAME"
'         .Fields.Item(3).Value = "<LESSEE NAME>"
'         .Update
'      End If
'      .Close
'   End With
'   GoTo NEW_TABLE_NJ_Header
'
''   ----SKIPPED-----
''   Control account has been moved into the NominalLedger table.
''   New table on 07/05/2013 SpareTable1         ~~~~ Control Account ~~~~
''###############################################################################################################
'NEW_TABLE_SpareTable1:
'   On Error GoTo MissingTable_SpareTable1
'
'   Rst1.Open "SELECT * FROM SpareTable1;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   Rst1.Open "SELECT * FROM SpareTable1;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields.count = 1 Or Rst1.Fields.count = 7 Then
'      If Rst1.Fields.count = 1 Then
'Debug.Print time
'Debug.Print 3082
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN CAName    TEXT(50);"
'Debug.Print 3082
'Debug.Print time
'Debug.Print 3083
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN NCode     TEXT(10);"
'Debug.Print 3083
'Debug.Print time
'Debug.Print 3084
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN NName     TEXT(100);"
'Debug.Print 3084
'Debug.Print time
'Debug.Print 3085
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN ClientID  TEXT(10);"
'Debug.Print 3085
'Debug.Print time
'Debug.Print 3086
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Fixed     BIT;"
'Debug.Print 3086
'Debug.Print time
'Debug.Print 3087
'         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN DisOrder  Single;"
'Debug.Print 3087
'      End If
'      Rst1.Close
'Debug.Print time
'Debug.Print 3090
'      Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Posting   BIT;"
'Debug.Print 3090
'Debug.Print time
'Debug.Print 3091
'      Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Type      TEXT(1);"
'Debug.Print 3091
'
'      Rst2.Open "SELECT * FROM Client;", Conn1, adOpenStatic, adLockReadOnly
'      While Not Rst2.EOF
'Debug.Print time
'Debug.Print 3095
'         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
'                       "Values ('" & UniqueID() & "', 'Sales Ledger Control', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 1, FALSE, 'S');"
'Debug.Print time
'         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
'                       "Values ('" & UniqueID() & "', 'Purchase Ledger Control', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 2, FALSE, 'P');"
'Debug.Print time
'Debug.Print 3099
'         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
'                       "Values ('" & UniqueID() & "', 'Input VAT', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 3, FALSE, 'I');"
'Debug.Print time
'Debug.Print 3101
'         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
'                       "Values ('" & UniqueID() & "', 'Output VAT', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 4, FALSE, 'O');"
'Debug.Print time
'Debug.Print 3103
'         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
'                       "Values ('" & UniqueID() & "', 'Retained Earnings', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 5, FALSE, 'R');"
'         Rst2.MoveNext
'      Wend
'      Rst2.Close
'   Else
'      Rst1.Close
'   End If
'
'   GoTo NEW_TABLE_NJ_Header
'
'MissingTable_SpareTable1:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - SpareTable1"
'   UpdateDatabase1 = -1
'   Exit Function
'
''   New table on 14/05/2013 NJ_Header            ~~~~  Nominal Journal Header  ~~~~
''###############################################################################################################
'NEW_TABLE_NJ_Header:
'   On Error GoTo MissingTable_NJ_Header
'
'   Rst1.Open "SELECT * FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo NEW_TABLE_NJ_Split
'
'MissingTable_NJ_Header:
'Debug.Print time
'Debug.Print 3130
'   Conn1.Execute _
'      "CREATE TABLE NJ_Header " & _
'         "(" & _
'            "RecordID      LONG NOT NULL PRIMARY KEY, " & _
'            "ClientID      TEXT(10) NOT NULL, " & _
'            "PropertyID    TEXT(4), " & _
'            "NJDate        TEXT(20), " & _
'            "NJTitle       TEXT(100), " & _
'            "History       BIT, " & _
'            "Posted        BIT, " & _
'            "PrintThis     BIT" & _
'         ");"
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 14/05/2013 NJ_Split         ~~~~  Nominal Journal Splits  ~~~~
''###############################################################################################################
'NEW_TABLE_NJ_Split:
'   On Error GoTo MissingTable_NJ_Split
'
'   Rst1.Open "SELECT * FROM NJ_Split;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo NEW_TABLE_ReportCategory
'
'MissingTable_NJ_Split:
'Debug.Print time
'Debug.Print 3157
'   Conn1.Execute _
'      "CREATE TABLE NJ_Split " & _
'         "(" & _
'            "RecordID      TEXT(50) NOT NULL PRIMARY KEY, " & _
'            "ParentID      LONG NOT NULL, " & _
'            "NC            TEXT(10), " & _
'            "SpLineDes     TEXT(200), " & _
'            "FundID        LONG, " & _
'            "TYPE_ID       BYTE, " & _
'            "NetAmt        CURRENCY, " & _
'            "VAT_CODE      TEXT(5), " & _
'            "VATAmt        CURRENCY, " & _
'            "TotalAmt      CURRENCY" & _
'         ");"
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   New table on 17/05/2013 ReportCategory         ~~~~  Report Category  ~~~~
''###############################################################################################################
'NEW_TABLE_ReportCategory:
'   On Error GoTo MissingTable_ReportCategory
'
'   Rst1.Open "SELECT * FROM ReportCategory;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADD_ClientID_tlbPayment
'
'MissingTable_ReportCategory:
'Debug.Print time
'Debug.Print 3186
'   Conn1.Execute _
'      "CREATE TABLE ReportCategory " & _
'         "(" & _
'            "RecordID      TEXT(50) NOT NULL PRIMARY KEY, " & _
'            "ClientID      TEXT(10) NOT NULL, " & _
'            "CategoryCode  TEXT(8), " & _
'            "CategoryName  TEXT(100), " & _
'            "CatDesc       TEXT(255)" & _
'         ");"
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new column ClientID on 24/05/2013 tlbPayment
''###############################################################################################################
'ADD_ClientID_tlbPayment:
'   On Error GoTo ERROR_ADD_ClientID_tlbPayment
'
'   Rst1.Open "SELECT ClientID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo VIEWS_NJ
'ERROR_ADD_ClientID_tlbPayment:
'Debug.Print time
'Debug.Print 3209
'   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN ClientID TEXT(10);"
'Debug.Print 3209
'
''  Update ClientID if there is existing PP
'
'Debug.Print time
'Debug.Print 3213
'   Conn1.Execute "UPDATE tlbPayment AS P, tlbClientBanks AS B " & _
'                 "SET P.ClientID = B.CLIENT_ID " & _
'                 "WHERE NOT ISNULL(P.NominalCode) AND P.NominalCode <> '' AND " & _
'                     "P.NominalCode =  B.NominalCode;"
'
'   UpdateDatabase1 = 1
'   Exit Function
'
''  CREATE a View NJ_HeaderTotal ON 28/05/2013
''###############################################################################################################
'VIEWS_NJ:
'   On Error GoTo ERROR_VIEWS_NJ
'
'   Rst1.Open "SELECT * FROM NJ_HeaderTotal;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo ADDNEW_REC_CAT
'
'ERROR_VIEWS_NJ:
'Debug.Print time
'Debug.Print 3232
'   Conn1.Execute "CREATE VIEW NJ_HeaderTotal " & _
'                  "AS SELECT D.RecordID, D.ClientID, D.PropertyID, D.NJDate, D.NJTitle, D.History, SUM(S.TotalAmt) AS TotalAmt " & _
'                  "FROM NJ_Header AS D INNER JOIN NJ_Split AS S ON D.RecordID=S.ParentID " & _
'                  "GROUP BY D.RecordID, D.ClientID, D.PropertyID, D.NJDate, D.NJTitle, D.History;"
'Debug.Print time
'Debug.Print 3236
''   Conn1.Execute "CREATE VIEW NJ_ID " & _
'Debug.print 3236
''                  "AS SELECT MAX(RecordID)+1 AS NJ " & _
''                  "FROM NJ_Header;"
'   UpdateDatabase1 = 1
'   Exit Function
'
''   Add new record CAT on 27/06/13 PrimarySecondaryCode     --> Control Account Type
''###############################################################################################################
'ADDNEW_REC_CAT:
'   On Error GoTo MissingRec_REC_CAT
'
'   With Rst1
'
'      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'CAT';", Conn1, adOpenStatic, adLockReadOnly
'
'      If .EOF Then
'         .Close
'         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         !Code = "CAT"
'         !Value = "Control Account Type"
'         .Update
'         .Close
'         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
'         .AddNew
'         .Fields.Item(0).Value = "CAT"
'         .Fields.Item(1).Value = "S"
'         .Fields.Item(2).Value = "Sales"
'         .Fields.Item(3).Value = "Sales Control Account"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "CAT"
'         .Fields.Item(1).Value = "P"
'         .Fields.Item(2).Value = "Purchase"
'         .Fields.Item(3).Value = "Purchase Control Account"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "CAT"
'         .Fields.Item(1).Value = "I"
'         .Fields.Item(2).Value = "Input VAT"
'         .Fields.Item(3).Value = "Input VAT"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "CAT"
'         .Fields.Item(1).Value = "O"
'         .Fields.Item(2).Value = "Output VAT"
'         .Fields.Item(3).Value = "Output VAT"
'         .Update
'         .AddNew
'         .Fields.Item(0).Value = "CAT"
'         .Fields.Item(1).Value = "R"
'         .Fields.Item(2).Value = "Retained Earnings"
'         .Fields.Item(3).Value = "Retained Earnings"
'         .Update
'         .Close
'      End If
'      .Close
'   End With
'   GoTo Modify_TABLE_PropertyAnalysis
'   'Exit Function
'
'MissingRec_REC_CAT:
''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
'   UpdateDatabase1 = 1
'   Exit Function
'Modify_TABLE_PropertyAnalysis:
'   On Error GoTo Mod__PropertyAnalysis
'   Rst1.Open "SELECT Reference FROM PropertyAnalysis;", Conn1, adOpenStatic, adLockReadOnly
'      If Rst1.State = 1 Then
'         Rst1.Close
'      End If
'      Exit Function
'Mod__PropertyAnalysis:
''Debug.Print time
''Debug.Print 3309
'      Conn1.Execute "ALTER TABLE PropertyAnalysis add COLUMN Reference text(255);"
''Debug.Print 3309
''Debug.Print time
''Debug.Print 3310
'      Conn1.Execute "ALTER TABLE PropertyAnalysis add COLUMN AnalysisValue1 Number;"
''Debug.Print 3310
'      Exit Function
'
'   UpdateDatabase1 = 1
'   Exit Function
'End Function
Private Function UpdateDatabase4(Conn1 As ADODB.Connection)
     'Written by anol 2022-09-01
   Dim Rst2       As New ADODB.Recordset
   Dim Rst3       As New ADODB.Recordset
   Dim Rst4       As New ADODB.Recordset
   Dim lSlNumber  As Long
   Dim i          As Integer
   Dim iRec       As Integer
   Dim szEmail    As String
   Dim szSQL      As String
   Dim szSQL_     As String
   Dim szSQL__    As String
   Dim szSQL___    As String
   Dim cOS        As Currency
    'Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record
 '********************************************************************************************
   szSQL = "SELECT A.* From tlbPayment A LEFT JOIN tlbPaymentSplit B ON A.TransactionID= B.PayHeader WHERE  B.PayHeader is NULL;"
   'Debug.Print "Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record:" & time
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   'Debug.Print "Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record:" & time
   If Not Rst1.EOF Then
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'TlbPayment has entry but tlbPaymentSplit has no record" & Rst1.Fields.Item("TransactionID").Value & "' )"
      szSQL = "SELECT * FROM tlbPaymentSplit;"
      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
      While Not Rst1.EOF
         With Rst2
            .AddNew
            .Fields.Item("TransactionID").Value = UniqueID()
            .Fields.Item("PayHeader").Value = Rst1.Fields.Item("TransactionID").Value
            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
            .Fields.Item("OSAmount").Value = Rst1.Fields.Item("OSAmount").Value
            .Fields.Item("SplitID").Value = 1
            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
            .Fields.Item("Description").Value = "DELETED PURCHASE TRANSACTIONS"
            .Update
         End With
         Rst1.MoveNext
      Wend
      Rst2.Close
   End If
   Rst1.Close

FIXING_SplitID_tlbPaymentSplit:
   szSQL = "SELECT A.PayHeader, A.X " & _
           "FROM ( " & _
                 "SELECT PayHeader, SplitID, COUNT(SplitID) AS X " & _
                 "From tlbPaymentSplit " & _
                 "GROUP BY PayHeader, SplitID " & _
           ") AS A " & _
           "WHERE A.X > 1;"
'
'            'Debug.Print szSQL
'             'Debug.Print time
    'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
'   'Debug.Print time
   If Not Rst1.EOF Then
      'keep a log in the error log table
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbPaymentSplit SplitID duplicated " & Rst1.Fields("PayHeader").Value & "' )"
      While Not Rst1.EOF
         Rst2.Open "SELECT * FROM tlbPaymentSplit AS S " & _
                   "WHERE S.PayHeader = " & Rst1.Fields.Item(0).Value & ";", Conn1, adOpenDynamic, adLockOptimistic

         For i = 1 To RecordCount(Rst2) 'Rst1.Fields.Item("X").Value 'fixed by anol 20181119
            Rst2.Fields.Item("SplitID").Value = i
            Rst2.Update
            Rst2.MoveNext
         Next i

         Rst1.MoveNext
         Rst2.Close
      Wend
   End If
   Rst1.Close
FIXING_SplitID_tlbReceiptSplit:
   szSQL = "SELECT A.RptHeader, A.X " & _
           "FROM ( " & _
                 "SELECT RptHeader, SplitID, COUNT(SplitID) AS X " & _
                 "From tlbReceiptSplit " & _
                 "GROUP BY RptHeader, SplitID " & _
           ") AS A " & _
           "WHERE A.X > 1;"
'
'            'Debug.Print szSQL
'             'Debug.Print time
    'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
'   'Debug.Print time
   If Not Rst1.EOF Then
      'keep a log in the error log table
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbReceiptSplit SplitID duplicated " & Rst1.Fields("RptHeader").Value & "' )"
      While Not Rst1.EOF
         Rst2.Open "SELECT * FROM tlbReceiptSplit AS S " & _
                   "WHERE S.RptHeader = " & Rst1.Fields.Item(0).Value & ";", Conn1, adOpenDynamic, adLockOptimistic

         For i = 1 To RecordCount(Rst2) 'Rst1.Fields.Item("X").Value 'fixed by anol 20181119
            Rst2.Fields.Item("SplitID").Value = i
            Rst2.Update
            Rst2.MoveNext
         Next i

         Rst1.MoveNext
         Rst2.Close
      Wend
   End If
   Rst1.Close
End Function
'Private Function UpdateDatabase4_old(Conn1 As ADODB.Connection) 'I am writing a new one beacuse I need to stop checking less important checks becuase of time it takes to load the software
'    'This function has been created for the replacement of UpdateDatabase1.
'    'Benifit 10 sec loading time and it records the Error in error log table
'    'Prevoius function took 27 seconds
'    'New function take 17 Seconds
'    'Written by anol 2019-10-21
'   Dim Rst2       As New ADODB.Recordset
'   Dim Rst3       As New ADODB.Recordset
'   Dim Rst4       As New ADODB.Recordset
'   Dim lSlNumber  As Long
'   Dim i          As Integer
'   Dim iRec       As Integer
'   Dim szEmail    As String
'   Dim szSQL      As String
'   Dim szSQL_     As String
'   Dim szSQL__    As String
'   Dim szSQL___    As String
'   Dim cOS        As Currency
'
'   Rst1.Open "Select Code from Sparetable5", Conn1, adOpenStatic, adLockReadOnly
'   If Rst1("Code").DefinedSize <> 255 Then
'        Rst1.Close
'        Conn1.Execute "ALTER Table SpareTable5 ALTER COLUMN Code Text(255);"
'   Else
'        Rst1.Close
'   End If
'
''***************************************************************************************************************
''                  Check System data - RoA/SRR without Fund 20/03/10 tlbReceipt                                '
''###############################################################################################################
'Fund_4_RoA:
''Debug.Print "RoA/SRR without Fund 20/03/10 tlbReceipt." & time
'   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
'                  "RDate, Details, Amount, BankCode, ExtRef " & _
'             "FROM tlbReceipt " & _
'             "WHERE (Type = 4 OR Type = 23) AND " & _
'                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
'             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly
''Debug.Print "RoA/SRR without Fund 20/03/10 tlbReceipt ." & time
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo CheckData_Multiple_SI_tlbReceipt
'   End If
'
'   If MsgBox("There are receipts on account and receipt refunds with no fund assigned." & Chr(13) & _
'          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
'          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
'      Load frmFund4RoA_RF
'      frmFund4RoA_RF.Show 1
'   End If
'
'   Rst1.Close
'
'
'
''  Check SI in receipt table on 22/10/2010 tlbReceipt
''  There is a data corraption has arirse. Duplicate SI has been exported to receipt table.
''  System identify them and clear from the db.
''###############################################################################################################
'CheckData_Multiple_SI_tlbReceipt:
'
'   szSQL = "SELECT X.SageAccountNumber, X.DemandRef " & _
'             "FROM " & _
'               "(" & _
'                 "SELECT COUNT(DemandRef) AS C, SageAccountNumber, DemandRef " & _
'                 "FROM tlbReceipt " & _
'                 "WHERE type = 1 " & _
'                 "GROUP by demandref, SageAccountNumber " & _
'               ") AS X " & _
'             "Where C > 1;"
'
''Debug.Print "CheckData_Multiple_SI_tlbReceipt." & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''Debug.Print "CheckData_Multiple_SI_tlbReceipt." & time
'   If Not Rst1.EOF Then
'      'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'CheckData_Multiple_SI_tlbReceipt" & Rst1.Fields.Item("SageAccountNumber").Value & "' )"
'      Rst1.Close
'      If MsgBox("System will update your data." + Chr(10) + _
'                "Please make sure everyone is out of the system before you click YES.", vbYesNo, "Receipt Record Update") = vbNo Then
'         Exit Function
'      Else
'         szSQL = "SELECT X.DemandRef, X.C " & _
'                   "FROM " & _
'                     "(" & _
'                       "SELECT COUNT(DemandRef) AS C, DemandRef " & _
'                       "FROM tlbReceipt " & _
'                       "WHERE type = 1 " & _
'                       "GROUP by demandref, SageAccountNumber " & _
'                     ") AS X " & _
'                   "Where C > 1;"
'         Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'         While Not Rst1.EOF
'            szSQL = "SELECT * " & _
'                      "FROM tlbReceipt " & _
'                      "WHERE DemandRef = " & Rst1.Fields.Item("DemandRef").Value & " " & _
'                      "ORDER BY TransactionID;"
'
'            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'            For i = 1 To Rst1.Fields.Item("C").Value - 1
'               Conn1.Execute "DELETE tlbReceiptSplit.* " & _
'                             "FROM tlbReceipt, tlbReceiptSplit " & _
'                             "WHERE tlbReceiptSplit.RptHeader = tlbReceipt.TransactionID AND " & _
'                                 "tlbReceipt.TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"
'
'               Conn1.Execute "DELETE * " & _
'                             "FROM tlbReceipt " & _
'                             "WHERE TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"
'               Rst2.MoveNext
'            Next i
'            Rst2.Close
'            Rst1.MoveNext
'         Wend
'      End If
'   End If
'
'   Rst1.Close
''************************************************************
''Demand table and receipt table total amount comparision
'
'CHECK_DATA_CORRUPTION_DEMAND:
'   szSQL = "SELECT R.TransactionID, R.Amount, R.DemandRef, S.DT " & _
'             "FROM tlbReceipt AS R, " & _
'                 "(SELECT D.DemandID,  SUM(S.TotalAmount) AS DT " & _
'                  "FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                  "GROUP BY D.DemandID) AS S " & _
'             "WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND " & _
'                  "ROUND(R.Amount, 2) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 2);"
''Debug.Print "CHECK_DATA_CORRUPTION_DEMAND." & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "CHECK_DATA_CORRUPTION_DEMAND." & time
'
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo CHECK_DATA_CORRUPTION_RECEIPT 'CHECK_BACS_EMAIL_tlbClientBanks
'   End If
''   'Debug.Print szSQL
'   MsgBox "DATA ERROR: PCM_1254" & Chr(13) & "PLEASE CONTACT WITH PCM SUPPORT.", vbCritical + vbOKOnly, "PRESTIGE SYSTEM"
''   Conn1.Execute "Delete from tlbReceipt where transactionid=" & Rst1("TransactionID").Value & ""
''   Conn1.Execute "Delete from tlbReceiptSplit where RptHeader=" & Rst1("TransactionID").Value & ""
''   Conn1.Execute "Update DemandSplitRecords set TrfReceipt = FALSE  where DemandID=" & Rst1("DemandRef").Value & ""
''   MigrateInvIntoReceipt Conn1
'   Rst1.Close
'   UpdateDatabase4 = -1
'   Exit Function
'
'
'
'
'
'CHECK_DATA_CORRUPTION_RECEIPT:
''  Check is there any transaction in the receipt table with 0 value in the header but non-0 value in the split.
''  if found then make the split value to 0. otherwise divident by zero will arise.
'   szSQL = "SELECT R.TransactionID " & _
'           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
'                                  "From tlbReceiptSplit " & _
'                                  "GROUP BY RptHeader " & _
'                                 ") AS Q " & _
'           "WHERE R.TransactionID=Q.RptHeader AND " & _
'                 "Q.A > 0 AND R.Amount=0;"
'''Debug.Print szSQL
''Debug.Print "CHECK_DATA_CORRUPTION_RECEIPT." & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "CHECK_DATA_CORRUPTION_RECEIPT." & time
'   If Not Rst1.EOF Then
'         'keep a log in the error log table
'         Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'Transaction in the receipt table with 0 value in the header but non-0 value in the split." & Rst1.Fields.Item("TransactionID").Value & "' )"
'   End If
'   While Not Rst1.EOF
'      Conn1.Execute "UPDATE tlbReceiptSplit " & _
'                    "SET    Amount = 0, OSAmount = 0 " & _
'                    "WHERE  RptHeader = " & Rst1.Fields.Item(0).Value & ";"
'      Rst1.MoveNext
'   Wend
'
'   Rst1.Close
''**********************************************************************************
''Check tlbReceiptSplit and tlbReceipt amount some and OSamount are consistence
''Sol 1: Collect the list of TransactionID which are corrupted.
'
'   szSQL = "SELECT R.TransactionID, Q.A-R.Amount AS D, Q.A / R.Amount AS R " & _
'           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
'                                  "From tlbReceiptSplit " & _
'                                  "GROUP BY RptHeader " & _
'                                 ") AS Q " & _
'           "WHERE R.TransactionID = Q.RptHeader AND " & _
'                 "ROUND(R.Amount, 2) <> ROUND(Q.A, 2) AND " & _
'                 "R.Amount+Q.A>0 AND R.OSAmount + Q.A <> R.Amount;"
'''Debug.Print szSQL
''Debug.Print "Collect the list of TransactionID which are corrupted." & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''Debug.Print "Collect the list of TransactionID which are corrupted" & time
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo FIX_OS_tlbReceiptSplit
'   End If
'
'   iRec = Rst1.RecordCount
'   i = 0
'   While Not Rst1.EOF
'      If Val(Rst1.Fields.Item("D").Value) > 0 And _
'            CInt(Rst1.Fields.Item("R").Value) = Val(Rst1.Fields.Item("R").Value) Then
'            Conn1.Execute "DELETE * FROM tlbReceiptSplit " & _
'                       "WHERE RptHeader = " & Rst1.Fields.Item("TransactionID").Value & " AND " & _
'                           "SplitID > 1;"
'         i = i + 1
'      End If
'      Rst1.MoveNext
'   Wend
'   Rst1.Close
'
'   If iRec > 0 Then
'      Conn1.Execute "DELETE S.*  FROM tlbReceiptSplit AS S LEFT JOIN tlbReceipt AS R ON  S.RptHeader = R.TransactionID  WHERE R.TransactionID is NULL"
'   End If
'   If iRec <> i Then
'      MsgBox "Warning: This Company data need to be updated. Please contact with PCM.", vbExclamation + vbOKOnly, "Data Update"
'      GoTo FIX_OS_tlbReceiptSplit
'   End If
'
''   Fix OS balance in the receipt_split on 06/10/2011
''   There is an inconsistency found in the total OS of splits and the OS of the receipt header.
''###############################################################################################################
'FIX_OS_tlbReceiptSplit:
''Debug.Print "There is an inconsistency found in the total OS of splits and the OS of the receipt header." & time
'   Rst1.Open "SELECT R.TransactionID, R.OSAmount AS R_OS, S.S_OS " & _
'             "FROM tlbReceipt AS R, ( " & _
'                     "SELECT S.RptHeader, SUM(S.OSAmount) AS S_OS " & _
'                     "FROM tlbReceiptSplit AS S " & _
'                     "GROUP BY S.RptHeader " & _
'                     ") AS S " & _
'             "WHERE R.TransactionID =  S.RptHeader AND " & _
'                     "ROUND(R.OSAmount, 2) <> ROUND(S.S_OS, 2) " & _
'             "ORDER BY S.RptHeader;", Conn1, adOpenStatic, adLockReadOnly
''Debug.Print "There is an inconsistency found in the total OS of splits and the OS of the receipt header." & time
'
'   If Rst1.RecordCount > 0 Then        'need to fix data. inconsistent data found
'    'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'There is an inconsistency found in the total OS of splits and the OS of the receipt header." & Rst1.Fields.Item("TransactionID").Value & "' )"
'      While Not Rst1.EOF
'         If Rst1.Fields.Item("R_OS").Value = 0 Then
'            Conn1.Execute "UPDATE tlbReceiptSplit " & _
'                          "SET    OSAmount = 0 " & _
'                          "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";"
'         End If
'         If Rst1.Fields.Item("R_OS").Value > 0 Then
'            cOS = Round(CCur(Rst1.Fields.Item("R_OS").Value), 2)
'            Rst2.Open "SELECT * FROM tlbReceiptSplit " & _
'                      "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";", _
'                      Conn1, adOpenDynamic, adLockOptimistic
'            While Not Rst2.EOF
'               If cOS = 0 Then
'                  Rst2.Fields.Item("OSAmount").Value = 0
'               End If
'               If cOS <= Round(CCur(Rst2.Fields.Item("Amount").Value), 2) And cOS > 0 Then
'                  Rst2.Fields.Item("OSAmount").Value = cOS
'                  cOS = 0
'               End If
'               If cOS > Round(CCur(Rst2.Fields.Item("Amount").Value), 2) Then
'                  Rst2.Fields.Item("OSAmount").Value = Rst2.Fields.Item("Amount").Value
'                  cOS = cOS - Rst2.Fields.Item("Amount").Value
'               End If
'
'               Rst2.MoveNext
'            Wend
'            Rst2.Close
'         End If
'
'         Rst1.MoveNext
'      Wend
'   End If
'
'   Rst1.Close
'
'   GoTo FIX_DEMANDTYPE_LLEASE
''
''
'''   Check the system PI Transactions which have not exported to payment table.
'''   If there is any transaction found, then system will export them
'''###############################################################################################################
''FIX_EXPORT_tblpurinv:
''  szSQL = "SELECT MY_ID FROM tblPurInv P LEFT JOIN tlbPayment M  ON P.MY_ID = M.PI WHERE P.TransactionType <> 25 AND P.TOTAL_AMOUNT <> 0 AND M.PI is NULL"
''''Debug.Print szSQL
''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''
''   If Not Rst1.EOF Then
''        'Insert into Error table no msg
'''      MsgBox "Your data need to be updated. Please contact with PCM", vbCritical + vbOKOnly, "PI not in Split table"
'''      Rst1.Close
'''      UpdateDatabase4 = -1
''
''' MigratePIIntoPayment method export only PI's header. MigratePIIntoPayment does not create the split in the
''' payment split table. Technically, there should not any PI in the purchase invoice table. because, when
''' users create a PI, system immidiately exports the transaction to the PP and PP_split table automatically.
'''08/08/2012
''      Exit Function
'''      szSQL = "UPDATE tblPurInv P LEFT JOIN tlbPayment M ON P.MY_ID = M.PI   SET   P.TrfPayment = FALSE WHERE    TOTAL_AMOUNT <> 0 AND  PI is NULL;"
'''      Conn1.Execute szSQL
'''      MigratePIIntoPayment Conn1
''   End If
''   Rst1.Close
''
''
''
'''   Check the system for wrong demand type setup in the lease.
''   If there is any transaction found, then system will warn user and if user wants
''   they will able to print the list of lease need to be fix manually.
''   27/10/2011
''   The reports are generating only DemandType and Lease tables entry.
''   But still system can generating this warning by checking the demand and split table
''   Please check szSQL___
''   08/01/2014
''###############################################################################################################
'FIX_DEMANDTYPE_LLEASE:
'   szSQL = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.BRDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LRentCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.BRDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.BRDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "Where PropertyID <> LeaseProperty"
'''Debug.Print szSQL
''Debug.Print "Check the system for wrong demand type setup in the lease." & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "Check the system for wrong demand type setup in the lease." & time
'   szSQL_ = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.SCDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LServiceCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.SCDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.SCDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "WHERE PropertyID <> LeaseProperty"
'''Debug.Print szSQL
'   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   szSQL__ = "SELECT ID " & _
'           "FROM (" & _
'               "SELECT R.InsuranceDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
'               "FROM DemandTypes AS T, LInsuranceCharges AS R, LeaseDetails AS L, Units AS U " & _
'               "Where T.ID = R.InsuranceDemandType And R.LeaseID = L.LeaseID And " & _
'                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
'               "GROUP BY R.InsuranceDemandType, T.PropertyID, U.PropertyID " & _
'           ") AS Q " & _
'           "WHERE PropertyID <> LeaseProperty"
'''Debug.Print szSQL
'   Rst3.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'' check other charges
'   szSQL___ = "SELECT ID, P1, P2 " & _
'           "FROM (" & _
'               "SELECT S.TypeOfDemand AS ID, U.PropertyID AS P1, DT.PropertyID AS P2 " & _
'               "FROM DemandRecords AS D, DemandSplitRecords AS S, Tenants AS T, LeaseDetails AS L, Units AS U, DemandTypes AS DT " & _
'               "WHERE L.Status AND D.DemandID = S.DemandID AND D.SageAccountNumber = T.SageAccountNumber AND " & _
'                     "T.SageAccountNumber = L.SageAccountNumber AND L.UnitNumber = U.UnitNumber AND " & _
'                     "S.TypeOfDemand = DT.ID " & _
'               "GROUP BY S.TypeOfDemand, U.PropertyID, DT.PropertyID " & _
'           ") AS Q " & _
'           "WHERE Q.P1 <> Q.P2"
'''Debug.Print szSQL
'   Rst4.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Or Not Rst2.EOF Or Not Rst3.EOF Or Not Rst4.EOF Then
'       'insert into error log table
'       'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'wrong demand type setup in the lease PropertyID not Equal LeaseProperty." & Rst1.Fields.Item("TransactionID").Value & "' )"
'      If MsgBox("There are some lease demand type need to fix." & _
'                "Do you want to do it now?", vbCritical + vbYesNo, "Demand Types") = vbNo Then
'
'         Rst1.Close
'         Rst2.Close
'         Rst3.Close
'         Rst4.Close
'         GoTo LeaseDetails_Lessee_Duplicate
'      End If
''      FIX_MODE__DT = True
'
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL & ");"
'''Debug.Print szSQL
'      Conn1.Execute szSQL
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL_ & ");"
'''Debug.Print szSQL
'      Conn1.Execute szSQL
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ID IN (" & szSQL__ & ");"
'''Debug.Print szSQL
'      Conn1.Execute szSQL
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_RC.rpt"
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_SC.rpt"
'      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_IC.rpt"
'
'      bFixingDT = True
'   Else
'      szSQL = "UPDATE DemandTypes " & _
'              "SET ULC = TRUE " & _
'              "WHERE ULC = FALSE;"
'      Conn1.Execute szSQL
'   End If
'   Rst1.Close
'   Rst2.Close
'   Rst3.Close
'   Rst4.Close
'
'
'
''***************************************************************************************************************
''                  Check System data - System will check lease details table for multiple lessee's account     '
''                  If system found any lessee more than one time in the lease table as active leases           '
''                  then system will stop user to use the system and push them to report it to us               '
''                                                                                                              '
''                                               21/12/2011                                                     '
''###############################################################################################################
'LeaseDetails_Lessee_Duplicate:
' 'Debug.Print "LeaseDetails_Lessee_Duplicate" & time
'   Rst1.Open "SELECT * " & _
'             "From " & _
'             "( " & _
'              "SELECT COUNT(SageAccountNumber) AS A, SageAccountNumber " & _
'              "From LeaseDetails " & _
'              "Where Status " & _
'              "GROUP BY SageAccountNumber " & _
'              ") AS Q " & _
'             "Where Q.a > 1;", Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "LeaseDetails_Lessee_Duplicate" & time
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo DUPLICATE_ID_LCLSA
'   Else
'    'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login',#" & Date & "# ,'Multiple Active lessee account found in LeaseDetails table " & Rst1.Fields.Item("SageAccountNumber").Value & "' )"
'      MsgBox "The system found an inconsistency in your database. Please contact PCM Consulting Support.", vbCritical + vbOKOnly, "Err. Multiple Lessee"
'      Rst1.Close
'      UpdateDatabase4 = -1
'      Exit Function
'   End If
'
'
'
'
''   Check Duplicate ID on 12/01/2012 Lessee, Client, Landlord, Supplier & MA
''###############################################################################################################
'DUPLICATE_ID_LCLSA:
'   szSQL = "SELECT ID "
'   szSQL = szSQL & "FROM (SELECT ID, COUNT(ID) AS C "
'   szSQL = szSQL & "FROM ("
'   szSQL = szSQL & "SELECT SupplierID AS ID "
'   szSQL = szSQL & "FROM Supplier WHERE TYPE = 'SUPPLIER' UNION ALL "
'   szSQL = szSQL & "SELECT ClientID AS ID "
'   szSQL = szSQL & "FROM Client UNION ALL "
'   szSQL = szSQL & "SELECT SageAccountNumber AS ID "
'   szSQL = szSQL & "FROM Tenants UNION ALL "
'   szSQL = szSQL & "SELECT AgentID AS ID "
'   szSQL = szSQL & "FROM Agent UNION ALL "
'   szSQL = szSQL & "SELECT LandlordID AS ID "
'   szSQL = szSQL & "From Landlord "
'   szSQL = szSQL & ") "
'   szSQL = szSQL & " GROUP BY ID "
'   szSQL = szSQL & ") "
'   szSQL = szSQL & "WHERE C > 1;"
'
'''Debug.Print szSQL
'   'Debug.Print "DUPLICATE_ID_LCLSA" & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'    'Debug.Print "DUPLICATE_ID_LCLSA" & time
'
'   If Not Rst1.EOF Then                                  'Duplicate ID found
''      Conn1.Execute "DELETE Tenants.* " & _
''                    "FROM Tenants, Landlord " & _
''                    "WHERE Tenants.SageAccountNumber = Landlord.LandlordID;"
'      Rst1.Close
'      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'      If Not Rst1.EOF Then                                  'Duplicate ID found
'         szSQL = SQL2String(Rst1, 0)
'         Rst1.Close
'        'keep a log in the error log table
'        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'he following ID(s) are duplicating(Lessee, Client, Landlord, Supplier):" & szSQL & "' )"
'         MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
'      Else
'         Rst1.Close
'      End If
'   Else
'      Rst1.Close
'   End If
'
'
'
''tlbPaymentSplit:
''   On Error GoTo MissingTable_tlbPaymentSplit
''
''   Rst1.Open "SELECT * FROM tlbPaymentSplit;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.RecordCount = 0 Then
'''     We have to check here: is there any PP partially allocated?
'''     If any PP found then user has to fix it before upgrade the system to PP split.
'''     System will generate a report for user to fix PP.
''      Rst2.Open "SELECT * FROM tlbPayment " & _
''                "WHERE Type > 7 AND OSAmount > 0 AND " & _
''                      "OSAmount < Amount;", Conn1, adOpenStatic, adLockReadOnly
''      If Not Rst2.EOF Then
''         Rst2.Close
''         Rst1.Close
''         ShowReport App.Path & szReportPath & "\PP_PartialAlloc.rpt"
''         MsgBox "Please print this report and unallocate these transactions." & Chr(13) & _
''                "After upgrade the system, re-allocate these transactions.", _
''                vbInformation + vbOKOnly, "Partially Allocated Purchase Payment"
''
''         GoTo FixDataByUser
''      Else
''         Rst2.Close
''      End If
''
''      If Not UpdateTlbPaymentSplit Then
''         UpdateDatabase4 = -1
''         Exit Function
''      End If
''   End If
''
''   Rst1.Close
'
'
''  Check DemandRecords table 'LeaseRef' column.
''  If 'LeaseRef' is empty then system will update according LeaseDetails table.
''  If the lease is expired then system will update 'LeaseRef' with latest expired lease
''###############################################################################################################
'FIX_LeaseRef_DemandRecords:
'    'Debug.Print "'LeaseRef' is empty then system will update according LeaseDetails table:" & time
'   Rst1.Open "SELECT D.* " & _
'             "FROM DemandRecords AS D INNER JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'             "WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef));", Conn1, adOpenDynamic, adLockOptimistic
'    'Debug.Print "'LeaseRef' is empty then system will update according LeaseDetails table" & time
'   If Rst1.EOF Then
'      Rst1.Close
'   Else
'    'Insert into Error log table
'    'keep a log in the error log table
'      Rst1.Close
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'LeaseRef is empty in DemandRecords' )"
'      Conn1.Execute "UPDATE DemandRecords AS D, " & _
'                        "[SELECT * FROM LeaseDetails WHERE Status = TRUE]. AS L, " & _
'                        "[SELECT D.LeaseRef, D.DemandID " & _
'                        " FROM DemandRecords AS D INNER JOIN " & _
'                        "     DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
'                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
'                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
'                    "WHERE D.DemandID = SQ.DemandID AND L.Status AND " & _
'                        "L.SageAccountNumber = D.SageAccountNumber;"
'
'      Conn1.Execute "UPDATE DemandRecords AS D, LeaseDetails AS L, " & _
'                        "[SELECT D.LeaseRef, D.DemandID " & _
'                        " FROM DemandRecords AS D INNER JOIN " & _
'                              "DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
'                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
'                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
'                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
'                    "WHERE D.DemandID = SQ.DemandID AND " & _
'                        "L.SageAccountNumber = D.SageAccountNumber;"
'   End If
'
'  ' Rst1.Close
''********************************************************************************************
' 'Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record
' '********************************************************************************************
'   szSQL = "SELECT A.* From tlbPayment A LEFT JOIN tlbPaymentSplit B ON A.TransactionID= B.PayHeader WHERE  B.PayHeader is NULL;"
'   'Debug.Print "Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record:" & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "Add tlbPaymentSplit record where tlbPayment has entry but tlbPaymentSplit has no record:" & time
'   If Not Rst1.EOF Then
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'TlbPayment has entry but tlbPaymentSplit has no record" & Rst1.Fields.Item("TransactionID").Value & "' )"
'      szSQL = "SELECT * FROM tlbPaymentSplit;"
'      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("TransactionID").Value = UniqueID()
'            .Fields.Item("PayHeader").Value = Rst1.Fields.Item("TransactionID").Value
'            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
'            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
'            .Fields.Item("OSAmount").Value = Rst1.Fields.Item("OSAmount").Value
'            .Fields.Item("SplitID").Value = 1
'            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
'            .Fields.Item("Description").Value = "DELETED PURCHASE TRANSACTIONS"
'            .Update
'         End With
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close
'
'
'
''********************************************************************************************
''   There are some PI found in the header table, which amounts don't match with split total
''********************************************************************************************
'   szSQL = "SELECT P.*, S.ST " & _
'           "FROM tblPurInv AS P, ( " & _
'               "SELECT ParentID, SUM(TOTAL_AMOUNT) AS ST " & _
'               "FROM tblPurInvSRec " & _
'               "GROUP BY ParentID) AS S " & _
'           "WHERE P.MY_ID = S.ParentID And round(P.TOTAL_AMOUNT,2) <> round(S.ST,2)"
'  'Debug.Print "There are some PI found in the header table, which amounts don't match with split total:" & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "There are some PI found in the header table, which amounts don't match with split total:" & time
'
'   If Not Rst1.EOF Then
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'There are some PI found in the header table, which amounts doesnt match with split total " & Rst1.Fields.Item("MY_ID").Value & "' )"
'      While Not Rst1.EOF
'         If Val(Rst1.Fields.Item("TOTAL_AMOUNT").Value) = 0 Then
'            Conn1.Execute "UPDATE tblPurInvSRec AS S " & _
'                          "SET    NET_AMOUNT = 0, VAT = 0, TOTAL_AMOUNT = 0 " & _
'                          "WHERE  S.ParentID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         Else
'            Conn1.Execute "UPDATE tblPurInv AS P " & _
'                          "SET    P.TOTAL_AMOUNT = " & Rst1.Fields.Item("ST").Value & " " & _
'                          "WHERE  P.MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         End If
'
'         Rst1.MoveNext
'      Wend
'   End If
'   Rst1.Close
'
'
'
'
'
'FIXING_SplitID_tlbPaymentSplit:
'   szSQL = "SELECT A.PayHeader, A.X " & _
'           "FROM ( " & _
'                 "SELECT PayHeader, SplitID, COUNT(SplitID) AS X " & _
'                 "From tlbPaymentSplit " & _
'                 "GROUP BY PayHeader, SplitID " & _
'           ") AS A " & _
'           "WHERE A.X > 1;"
''
''            'Debug.Print szSQL
''             'Debug.Print time
'    'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   'Debug.Print "FIXING_SplitID_tlbPaymentSplit:" & time
''   'Debug.Print time
'   If Not Rst1.EOF Then
'      'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbPaymentSplit SplitID duplicated " & Rst1.Fields("PayHeader").Value & "' )"
'      While Not Rst1.EOF
'         Rst2.Open "SELECT * FROM tlbPaymentSplit AS S " & _
'                   "WHERE S.PayHeader = " & Rst1.Fields.Item(0).Value & ";", Conn1, adOpenDynamic, adLockOptimistic
'
'         For i = 1 To RecordCount(Rst2) 'Rst1.Fields.Item("X").Value 'fixed by anol 20181119
'            Rst2.Fields.Item("SplitID").Value = i
'            Rst2.Update
'            Rst2.MoveNext
'         Next i
'
'         Rst1.MoveNext
'         Rst2.Close
'      Wend
'   End If
'   Rst1.Close
'
'
''********************************************************************************************
''   FIXING DATA: tlbReceiptSplit table has some transaction with FundID = 0
''********************************************************************************************
'FIXING_FUNDID_tlbReceiptSplit:
'   szSQL = "SELECT S.* " & _
'           "FROM tlbReceiptSplit AS S " & _
'           "WHERE S.FundID = 0;"
'   'Debug.Print "tlbReceiptSplit table has some transaction with FundID = 0:" & time
'   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
'   'Debug.Print "tlbReceiptSplit table has some transaction with FundID = 0:" & time
'
'   If Not Rst1.EOF Then
'      'keep a log in the error log table
'      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbReceiptSplit table has some transaction with FundID = 0 " & Rst1.Fields.Item("TransactionID").Value & "' )"
'      Conn1.Execute "UPDATE tlbReceiptSplit AS S1, tlbReceiptSplit AS S2, " & _
'                        "tlbReceipt AS R1,tlbReceipt AS R2, RptTransactions AS T " & _
'                    "Set S1.FundID = S2.FundID, S1.SplitID = S2.SplitID " & _
'                    "WHERE S1.FundID = 0 AND R1.TransactionID = T.FromTran AND " & _
'                        "T.ToTran = R2.TransactionID AND R1.TransactionID = S1.RptHeader AND " & _
'                        "R2.TransactionID = S2.RptHeader  AND VAL(S1.AllocTranID) = S2.RptHeader;"
'   End If
'   Rst1.Close
'
'
'
'
''********************************************************************************************
''   FIXING DATA: There are some SRR found in the receipt and split table, their split OS <> header OS
''********************************************************************************************
'FIXING_SRR_OS_Split:
'   szSQL = "SELECT S.* From tlbReceiptSplit AS S, tlbReceipt AS R WHERE   R.TransactionID = S.RptHeader AND R.Type = 23 AND R.OSAmount <> S.OSAmount;"
'   'Debug.Print "There are some SRR found in the receipt and split table, their split OS <> header OS:" & time
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'    'Debug.Print "There are some SRR found in the receipt and split table, their split OS <> header OS:" & time
'   If Not Rst1.EOF Then
'        'insert into error log table
'        'keep a log in the error log table
'        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'SRR found in tlbReceiptSplit table where split OS Not Equal header OS" & Rst1.Fields.Item("TransactionID").Value & "' )"
'        szSQL = "UPDATE  tlbReceiptSplit AS S, tlbReceipt AS R " & _
'                "SET     R.OSAmount = S.OSAmount " & _
'                "WHERE   R.TransactionID = S.RptHeader AND " & _
'                        "R.Type = 23 AND " & _
'                        "R.OSAmount <> S.OSAmount;"
'        Conn1.Execute szSQL
'   End If
'   Rst1.Close
'
'   szSQL = "SELECT S.* FROM  tlbReceiptSplit AS S  WHERE S.SplitID = -1;"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'        'insert into error log table
'        'keep a log in the error log table
'        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'In tlbReceiptSplit table S.SplitID = -1 Found" & Rst1.Fields.Item("TransactionID").Value & "' )"
'        szSQL = "UPDATE  tlbReceiptSplit AS S " & _
'                "SET     S.SplitID = 1 " & _
'                "WHERE   S.SplitID = -1;"
'        Conn1.Execute szSQL
'   End If
'   Rst1.Close
''Debug.Print "Checking There are some SA found in the receipt table without split in the split table:" & time
''********************************************************************************************
''   FIXING DATA: There are some SA found in the receipt table without split in the split table
''********************************************************************************************
''
'   szSQL = "SELECT A.* FROM  tlbReceipt A Left join tlbReceiptSplit B ON A.TransactionID =B.RptHeader WHERE B.RptHeader IS NULL AND Type = 4 AND A.Amount <> 0 ;"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'    'keep a log in the error log table
'    Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'SA found in the receipt table without split in tlbReceipt table " & Rst1.Fields.Item("TransactionID").Value & "' )"
'      szSQL = "SELECT * FROM tlbReceiptSplit;"
'      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic
'
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("TransactionID").Value = UniqueID()
'            .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
'            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
'            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
'            .Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
'            .Fields.Item("SplitID").Value = 1
'            .Fields.Item("DueDate").Value = Format(Rst1.Fields.Item("DDate").Value, "dd mmmm yyyy")
'            .Fields.Item("Description").Value = "Receipt on Account"
'            .Update
'         End With
'
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'
'   Rst1.Close
'   'Debug.Print "Checking There are some SA found in the receipt table without split in the split table:" & time
'   'Debug.Print "Checking SIPI:" & time
'
'   Call Pi_Check_pre(Conn1) 'written by anol issue 791 Batch Payments (Support - WPM) 2019-07-25
'
'   Call SiPi_Check(Conn1, "PI", "25875")
'   Call SiPi_Check(Conn1, "SI", "15875")
'   'Debug.Print "Checking SIPI:" & time
'End Function
Private Function UpdateDatabase1(Conn1 As ADODB.Connection)
   Dim Rst2       As New ADODB.Recordset
   Dim Rst3       As New ADODB.Recordset
   Dim Rst4       As New ADODB.Recordset
   Dim lSlNumber  As Long
   Dim i          As Integer
   Dim iRec       As Integer
   Dim szEmail    As String
   Dim szSQL      As String
   Dim szSQL_     As String
   Dim szSQL__    As String
   Dim szSQL___    As String
   Dim cOS        As Currency

'   loopcount = loopcount + 1



'   DoEvents
   UpdateDatabase1 = 0

'   'Resolved by BOSL
''0000468: Posting dates not implemented correctly
''added by anol anol 23 Nov 2014
'Modify_TABLE_tlbPayment:
'   On Error GoTo Mod_TABLE_tlbPayment
'
'   Rst1.Open "SELECT PostingDate FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields(0).Type = 135 Then
'      Rst1.Close
'      GoTo NEW_TABLE_MemoDetails
'   Else
'Mod_TABLE_tlbPayment:
'      Rst1.Close
'      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN PostingDate DateTime;"
'      GoTo NEW_TABLE_MemoDetails
'   End If
'   UpdateDatabase1 = 1
'   Exit Function

 '   Add new column  on 18/08/14 tlbClientBanks
'###############################################################################################################
'Resolved by BOSL
'Issue 458: Batch Payments Crashes with Upgraded Data
'Resolved by anol 18 Aug 2014
'Issue 466: Batch Payments Crashes with Upgraded Data
'Resolved by anol 27 Aug 2014
ADD_COLUMN_PostingDate:
 On Error GoTo ERR_tblBatchPayment
   Rst1.Open "SELECT PostingDate FROM tblBatchPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_PCB_tlbClientBanks
ERR_tblBatchPayment:
   Conn1.Execute "ALTER TABLE tblBatchPayment ADD COLUMN PostingDate Date;"
   UpdateDatabase1 = 1
   Exit Function



'   Add new column PCB on 29/04/10 tlbClientBanks
'###############################################################################################################
ADD_PCB_tlbClientBanks:
   On Error GoTo CHANGE_ADD_PCB_tlbClientBanks

   Rst1.Open "SELECT PCB FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_email_tlbClientBanks
CHANGE_ADD_PCB_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN PCB Currency;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column email on 29/04/XX tlbClientBanks
'###############################################################################################################
ADD_email_tlbClientBanks:
   On Error GoTo CHANGE_ADD_email_tlbClientBanks

   Rst1.Open "SELECT email FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_Memo_Property
CHANGE_ADD_email_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN email TEXT(100);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column Memo on 16/03/10 Property
'###############################################################################################################
ADDNEW_Memo_Property:
   On Error GoTo Error_ADDNEW_Memo_Property
   'Resolved by BOSL
   'issue 466
    'Debug.Print
   Rst1.Open "SELECT MemoText FROM Property;", Conn1, adOpenStatic, adLockReadOnly
Debug.Print Err.description
   Rst1.Close

   GoTo ADD_Fund_tlbReceipt

Error_ADDNEW_Memo_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN MemoText MEMO"
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col - RptAmtType) - DemandRecords"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column FundID on 18/03/10 tlbReceipt
'###############################################################################################################
ADD_Fund_tlbReceipt:
   On Error GoTo CHANGE_ADD_Fund_tlbReceipt

   Rst1.Open "SELECT FundID FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo Fund_4_RoA
CHANGE_ADD_Fund_tlbReceipt:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN FundID Long;"
   UpdateDatabase1 = 1
   Exit Function

'***************************************************************************************************************
'                  Check System data - RoA/SRR without Fund 20/03/10 tlbReceipt                                '
'###############################################################################################################
Fund_4_RoA:
   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
                  "RDate, Details, Amount, BankCode, ExtRef " & _
             "FROM tlbReceipt " & _
             "WHERE (Type = 4 OR Type = 23) AND " & _
                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      GoTo FundID_4_PoA
   End If

   If MsgBox("There are receipts on account and receipt refunds with no fund assigned." & Chr(13) & _
          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
      LoadForm frmFund4RoA_RF
'      frmFund4RoA_RF.Show 1
   End If

   Rst1.Close

'   Add new column FundID on 31/03/10 tlbPayment
'###############################################################################################################
FundID_4_PoA:
   On Error GoTo ERROR_Fund_4_PoA

   Rst1.Open "SELECT FundID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo DrCr_4_NominalLedger
ERROR_Fund_4_PoA:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN FundID Long;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column DrCr on 21/04/10 NominalLedger
'###############################################################################################################
DrCr_4_NominalLedger:
   On Error GoTo ERROR_DrCr_4_NominalLedger

   Rst1.Open "SELECT DrCr FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo Fund_4_PoA
ERROR_DrCr_4_NominalLedger:
   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN DrCr TEXT(2);"
   UpdateDatabase1 = 1
   Exit Function

'***************************************************************************************************************
'                  Check System data - PoA/PPR without Fund 20/03/10 tlbPayment                                '
'###############################################################################################################
Fund_4_PoA:

   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
                  "PDate, Details, Amount, BankCode, ExtRef " & _
             "FROM tlbPayment " & _
             "WHERE (Type = 9 OR Type = 24) AND " & _
                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      GoTo Fund_4_SR
   End If

   If MsgBox("There are payment on account and payment refunds with no fund assigned." & Chr(13) & _
          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
      LoadForm frmFund4PoA
'      frmFund4PoA.Show 1

      Rst1.Close

      UpdateDatabase1 = 1
      Exit Function
   End If

'
'   Rst1.Open "SELECT TransactionID, SageAccountNumber, UnitID, " & _
'                  "PDate, Details, Amount, BankCode, ExtRef " & _
'             "FROM tlbPayment " & _
'             "WHERE (Type = 9 OR Type = 24) AND " & _
'                  "(ISNULL(FundID) OR VAL(FundID) = 0) " & _
'             "ORDER BY TransactionID;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'      GoTo Fund_4_SR
'   End If
'
'   If MsgBox("There are payment on account and payment refunds with no fund assigned." & Chr(13) & _
'          "Do you wish to assign a fund now, then press 'YES' otherwise " & Chr(13) & _
'          "press 'NO' to assign later", vbInformation + vbYesNo, "Fund") = vbYes Then
'      Load frmFund4PoA
'      frmFund4PoA.Show 1
'   End If
'   Rst1.Close

Fund_4_SR:

   Rst1.Open "SELECT TransactionID FROM tlbReceipt WHERE (ISNULL(FundID) OR FundID = 0) AND Type = 3;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      GoTo Receipt_Splits
   End If
   Rst1.Close

'***************************************************************************************************************
'                  Check System data - System will check receipts' splits in the new table tlbReceiptSplit     '
'                  Existing data will not have any split in the split table.                                   '
'                  If there any receipt found without a split line then system will try to create the split    '
'                  record.
'                                               13/05/2010
'###############################################################################################################
Receipt_Splits:
   On Error GoTo RptSpitMissing

'  THE FOLLOWING PROCEDURE IS DISABLED FOR CLARKE PROPERTY. I HAVE UPDATED THEIR DATA.
'  THIS PROCEDURE CAUSES THEIR SYSTEM VERY SLOW TO LOAD
'   GenerateReceiptSplit Conn1
   GoTo ADD_CtrlCode_tlbTransactionTypes

RptSpitMissing:
   MsgBox "System could not update your database. Please contact with PCM.", vbCritical + vbOKOnly, "Err. Generating Receipt Split"
   UpdateDatabase1 = -1
   Exit Function

'   Add new column CtrlCode on 04/06/2010 tlbTransactionTypes
'###############################################################################################################
ADD_CtrlCode_tlbTransactionTypes:
   On Error GoTo CHANGE_ADD_CtrlCode_tlbTransactionTypes

   Rst1.Open "SELECT CtrlCode FROM tlbTransactionTypes;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo ADD_Group_tlbTransactionTypes

CHANGE_ADD_CtrlCode_tlbTransactionTypes:
   Conn1.Execute "ALTER TABLE tlbTransactionTypes ADD COLUMN CtrlCode TEXT(15);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column Group on 04/02/2014 tlbTransactionTypes
'###############################################################################################################
ADD_Group_tlbTransactionTypes:
   On Error GoTo CHANGE_ADD_Group_tlbTransactionTypes

   Rst1.Open "SELECT Group FROM tlbTransactionTypes;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo ADD_NEW_TABLE_NLPosting

CHANGE_ADD_Group_tlbTransactionTypes:
   Conn1.Execute "ALTER TABLE tlbTransactionTypes ADD COLUMN Group TEXT(50);"
   UpdateDatabase1 = 1
   Exit Function

'   New table on 07/06/2010 NLPosting
'###############################################################################################################
ADD_NEW_TABLE_NLPosting:
   On Error GoTo CHANGE_ADD_NEW_TABLE_NLPosting

   Rst1.Open "SELECT * FROM NLPosting;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo BUGFIX_tlbPayment_OSAmount

CHANGE_ADD_NEW_TABLE_NLPosting:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - NLPosting"
   UpdateDatabase1 = -1
   Exit Function
'
''   Update table 'tlbPayment' data on 16/06/2010 UnitID
''###############################################################################################################
'UPDATE_tlbPayment_UnitID:
'   On Error GoTo ERROR_UPDATE_tlbPayment_UnitID
'
'   Rst1.Open "SELECT * FROM tlbPayment WHERE ISNULL(UnitID) OR UnitID='';", Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      Rst1.Close
'      Conn1.Execute "UPDATE tlbPayment, Property SET UnitID = PropertyID " & _
'                    "WHERE ISNULL(UnitID) OR UnitID='';"
'   Else
'      Rst1.Close
'      GoTo BUGFIX_tlbPayment_OSAmount
'   End If
'
'   GoTo BUGFIX_tlbPayment_OSAmount
'
'ERROR_UPDATE_tlbPayment_UnitID:
'   Debug.Print ERR.Number & " " & ERR.description
'   Exit Function

'   Bugfix table 'tlbPayment' data on 21/06/2010 OSAmount
'###############################################################################################################
BUGFIX_tlbPayment_OSAmount:
   Rst1.Open "Select * from  tlbPayment WHERE Amount < OSAmount;", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then
        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbpayment Amount < OSAmount " & Rst1.Fields("TransactionID").Value & "' )"
        Conn1.Execute "UPDATE tlbPayment " & _
                 "SET OSAmount = Amount " & _
                 "WHERE Amount < OSAmount;"
   End If
   Rst1.Close
'  THE FOLLOWING SEGMENT OF CODE IS DISABLED FOR CLARKE PROPERTY. I HAVE UPDATED THEIR DATA.
'  THESE CODE CAUSES THEIR SYSTEM VERY SLOW TO LOAD
''
''   Update table 'tlbReceiptSplit' data on 01/07/2010
''###############################################################################################################
'   szsql = "SELECT * FROM tlbReceipt " & _
'             "WHERE TransactionID NOT IN (" & _
'               "SELECT RptHeader FROM tlbReceiptSplit);"
''Debug.Print szsql
'   Rst1.Open szsql, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      Rst2.Open "SELECT * FROM tlbReceiptSplit;", Conn1, adOpenDynamic, adLockOptimistic
'
'      While Not Rst1.EOF
'         With Rst2
'            .AddNew
'            .Fields.Item("TransactionID").Value = UniqueID()
'            .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
'            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
'            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
'            .Fields.Item("SplitID").Value = 1
'            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
'            .Fields.Item("Description").Value = Rst1.Fields.Item("Details").Value
'            .Update
'         End With
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close

'   Fixing data Amount in RptSplit on 08/07/2010
'###############################################################################################################
    Rst1.Open "Select * from tlbReceipt AS R, tlbReceiptSplit AS S WHERE S.Amount > R.Amount AND R.Type > 1 AND R.TransactionID = S.RptHeader;", Conn1, adOpenStatic, adLockReadOnly
    If Not Rst1.EOF Then
        Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbReceipt Amount < OSAmount " & Rst1.Fields("TransactionID").Value & "' )"
        szSQL = "UPDATE tlbReceipt AS R, tlbReceiptSplit AS S " & _
             "SET S.Amount = R.Amount " & _
             "WHERE S.Amount > R.Amount AND " & _
                  "R.Type > 1 AND " & _
                  "R.TransactionID = S.RptHeader;"
        Conn1.Execute szSQL
    End If
    Rst1.Close
'   Fixing data SlNumber in 'DemandRecords' & 'tlbReceipt' on 05/07/2010
'###############################################################################################################
   szSQL = "SELECT DMDSLNO, COUNT(DMDSLNO) AS X, TransactionType " & _
             "From DEMANDRECORDS " & _
             "GROUP BY DMDSLNO, TransactionType " & _
             "Having count(DMDSLNO) > 1 " & _
             "ORDER BY DMDSLNO;"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
'#
Debug.Print "In SI serial number duplicating!!!!! SAMRAT check it."
'Debug.Print szsql
      lSlNumber = SlNumber(IIf(Rst1.Fields.Item("TransactionType").Value = 1, "SI", "SC"), "DemandRecords", Conn1)
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'SI serial number duplicating,dmslno:" & Rst1.Fields.Item(0).Value & "' )"
      While Not Rst1.EOF
         szSQL = "SELECT DMDSLNO " & _
                   "FROM DEMANDRECORDS " & _
                   "WHERE DMDSLNO = " & Rst1.Fields.Item(0).Value & " AND " & _
                     "TransactionType = " & Rst1.Fields.Item("TransactionType").Value & " " & _
                   "ORDER BY DemandID;"
         Rst2.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
         Rst2.MoveNext
         While Not Rst2.EOF
            Rst2.Fields.Item(0).Value = lSlNumber
            Rst2.Update
            Rst2.MoveNext
            lSlNumber = lSlNumber + 1
         Wend
         Rst2.Close
         Rst1.MoveNext
      Wend
      Rst1.Close

      Conn1.Execute _
         "UPDATE tlbReceipt AS R, DemandRecords AS D " & _
         "SET    R.SlNumber = D.DmdSlNo " & _
         "WHERE  R.DemandRef = D.DemandID;"
   Else
      Rst1.Close
   End If

'   Fixing data SlNumber in 'tlbReceipt' on 13/07/2010
'###############################################################################################################
   szSQL = "UPDATE tlbReceipt AS R, DemandRecords AS D " & _
             "SET R.SlNumber = D.DmdSlNo " & _
             "WHERE R.DemandRef = D.DemandID AND " & _
               "R.SlNumber <> D.DmdSlNo AND R.Type = 1;"
   Conn1.Execute szSQL

'   New table on 07/06/2010 tblBtRptTran
'###############################################################################################################
ADD_NEW_TABLE_tblBtRptTran:
   On Error GoTo CHANGE_ADD_NEW_TABLE_tblBtRptTran

   Rst1.Open "SELECT * FROM tblBtRptTran;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo ADD_RptDt_tblBtRptTran

CHANGE_ADD_NEW_TABLE_tblBtRptTran:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBtRptTran"
   UpdateDatabase1 = -1
   Exit Function

'   Add new column RptDt on 19/07/2010 tblBtRptTran
'###############################################################################################################
ADD_RptDt_tblBtRptTran:
   On Error GoTo CHANGE_ADD_RptDt_tblBtRptTran

   Rst1.Open "SELECT RptDt FROM tblBtRptTran;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo FundID_4_PayTransactions
CHANGE_ADD_RptDt_tblBtRptTran:
   Conn1.Execute "ALTER TABLE tblBtRptTran ADD COLUMN RptDt TEXT(15);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column FundID on 29/07/10 PayTransactions
'###############################################################################################################
FundID_4_PayTransactions:
   On Error GoTo ERROR_Fund_4_PayTransactions

   Rst1.Open "SELECT FundID FROM PayTransactions;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_FundID_RptTransactions
ERROR_Fund_4_PayTransactions:
   Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN FundID Long;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new a columns FundID on 03/08/2010 RptTransactions
'###############################################################################################################
ADD_FundID_RptTransactions:
   On Error GoTo CHANGE_ADD_FundID_RptTransactions

   Rst1.Open "SELECT FundID FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_LeaseValue_LeaseDetails
CHANGE_ADD_FundID_RptTransactions:
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN FundID Long;"
   UpdateDatabase1 = 1
   Exit Function

'   Add a columns LeaseValue on 11/08/10 LeaseDetails
'###############################################################################################################
ADD_LeaseValue_LeaseDetails:
   On Error GoTo CHANGE_ADD_LeaseValue_LeaseDetails

   Rst1.Open "SELECT LeaseValue FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_GPrataDmd_LeaseDetails

CHANGE_ADD_LeaseValue_LeaseDetails:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN LeaseValue CURRENCY;"

   UpdateDatabase1 = 1
   Exit Function

'   Add a columns GPrataDmd on 11/08/10 LeaseDetails
'###############################################################################################################
ADD_GPrataDmd_LeaseDetails:
   On Error GoTo CHANGE_ADD_GPrataDmd_LeaseDetails

   Rst1.Open "SELECT GPrataDmd FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_SCYE_LServiceCharges

CHANGE_ADD_GPrataDmd_LeaseDetails:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN GPrataDmd BIT;"
   Conn1.Execute "UPDATE LeaseDetails SET GPrataDmd = TRUE;"

   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN FDD TEXT(25);"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN FDD TEXT(25);"
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN FDD TEXT(25);"
   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN StopAutoDmd BIT;"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN StopAutoDmd BIT;"
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN StopAutoDmd BIT;"
   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN GPD BIT;"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN GPD BIT;"
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN GPD BIT;"

   UpdateFDD Conn1
   UpdateStopAutoDmd Conn1

   UpdateDatabase1 = 1
   Exit Function

'   Add a columns SCYE on 08/08/2011 LServiceCharges
'###############################################################################################################
ADD_SCYE_LServiceCharges:
   On Error GoTo CHANGE_ADD_SCYE_LServiceCharges

   Rst1.Open "SELECT SCYE FROM LServiceCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_DueDate_tblPurInv

CHANGE_ADD_SCYE_LServiceCharges:
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN SCYE BIT;"
   UpdateDatabase1 = 1
   Exit Function

'   Add a columns DueDate on 11/08/10 tblPurInv
'###############################################################################################################
ADD_DueDate_tblPurInv:
   On Error GoTo CHANGE_ADD_DueDate_tblPurInv

   Rst1.Open "SELECT DueDate FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_PropertyID_ClientGlobalData

CHANGE_ADD_DueDate_tblPurInv:
   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN DueDate TEXT(20);"
   UpdateDatabase1 = 1
   Exit Function
'
''   Add a columns JobID on 18/08/10 DemandSplitRecords_tlbReceiptSplit
''###############################################################################################################
'ADD_JobID_DemandSplitRecords_tlbReceiptSplit:
'   On Error GoTo CHANGE_ADD_JobID_DemandSplitRecords_tlbReceiptSplit
'
'   Rst1.Open "SELECT JobID FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_PropertyID_ClientGlobalData
'
'CHANGE_ADD_JobID_DemandSplitRecords_tlbReceiptSplit:
'   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN JobID TEXT(10);"
'   Conn1.Execute "ALTER TABLE tlbReceiptSplit ADD COLUMN JobID TEXT(10);"
'   UpdateDatabase1 = 1
'   Exit Function

'   Add a columns PropertyID on 24/08/10 ClientGlobalData
'###############################################################################################################
ADD_PropertyID_ClientGlobalData:
   On Error GoTo CHANGE_ADD_PropertyID_ClientGlobalData

   Rst1.Open "SELECT PropertyID FROM ClientGlobalData;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_RecoverableExp_ChargeTypes

CHANGE_ADD_PropertyID_ClientGlobalData:
   Conn1.Execute "ALTER TABLE ClientGlobalData ADD COLUMN PropertyID TEXT(4);"
   UpdateDatabase1 = 1
   Exit Function

'   Add a columns RecoverableExp on 24/08/10 ChargeTypes
'###############################################################################################################
ADD_RecoverableExp_ChargeTypes:
   On Error GoTo CHANGE_ADD_RecoverableExp_ChargeTypes

   Rst1.Open "SELECT RecoverableExp FROM ChargeTypes;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_COL_RCCComments_Tenants

CHANGE_ADD_RecoverableExp_ChargeTypes:
   Conn1.Execute "ALTER TABLE ChargeTypes ADD COLUMN RecoverableExp TEXT(1);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column RCCComments on 23/09/2010 Tenants
'###############################################################################################################
ADDNEW_COL_RCCComments_Tenants:
   On Error GoTo MISSING_ADDNEW_COL_RCCComments_Tenants

   Rst1.Open "SELECT RCCComments FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_RCC_Property

MISSING_ADDNEW_COL_RCCComments_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN RCCComments TEXT(250);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column RCC on 24/09/2010 Property
'###############################################################################################################
ADDNEW_COL_RCC_Property:
   On Error GoTo MISSING_ADDNEW_COL_RCC_Property

   Rst1.Open "SELECT RCC FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_Comments_Supplier

MISSING_ADDNEW_COL_RCC_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN RCC TEXT(1);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column Comments on 06/12/2010 Supplier
'###############################################################################################################
ADDNEW_COL_Comments_Supplier:
   On Error GoTo MISSING_ADDNEW_COL_Comments_Supplier

   Rst1.Open "SELECT Comments FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_PCB_ABR_tlbClientBanks

MISSING_ADDNEW_COL_Comments_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN Comments TEXT(200);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column PCB_ABR on 24/12/10 tlbClientBanks
'###############################################################################################################
ADD_PCB_ABR_tlbClientBanks:
   On Error GoTo CHANGE_ADD_PCB_ABR_tlbClientBanks

   Rst1.Open "SELECT PCB_ABR FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo CheckData_Multiple_SI_tlbReceipt

CHANGE_ADD_PCB_ABR_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN PCB_ABR Currency;"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN SOB Currency;"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN LSD TEXT(50);"
   UpdateDatabase1 = 1
   Exit Function

'  Check SI in receipt table on 22/10/2010 tlbReceipt
'  There is a data corraption has arirse. Duplicate SI has been exported to receipt table.
'  System identify them and clear from the db.
'###############################################################################################################
CheckData_Multiple_SI_tlbReceipt:

   On Error GoTo ERROR_CheckData_Multiple_SI_tlbReceipt

   szSQL = "SELECT X.SageAccountNumber, X.DemandRef " & _
             "FROM " & _
               "(" & _
                 "SELECT COUNT(DemandRef) AS C, SageAccountNumber, DemandRef " & _
                 "FROM tlbReceipt " & _
                 "WHERE type = 1 " & _
                 "GROUP by demandref, SageAccountNumber " & _
               ") AS X " & _
             "Where C > 1;"

   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Rst1.Close
      If MsgBox("System will update your data." + Chr(10) + _
                "Please make sure everyone is out of the system before you click YES.", vbYesNo, "Receipt Record Update") = vbNo Then
         Exit Function
      Else
         szSQL = "SELECT X.DemandRef, X.C " & _
                   "FROM " & _
                     "(" & _
                       "SELECT COUNT(DemandRef) AS C, DemandRef " & _
                       "FROM tlbReceipt " & _
                       "WHERE type = 1 " & _
                       "GROUP by demandref, SageAccountNumber " & _
                     ") AS X " & _
                   "Where C > 1;"
         Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

         While Not Rst1.EOF
            szSQL = "SELECT * " & _
                      "FROM tlbReceipt " & _
                      "WHERE DemandRef = " & Rst1.Fields.Item("DemandRef").Value & " " & _
                      "ORDER BY TransactionID;"

            Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
            For i = 1 To Rst1.Fields.Item("C").Value - 1
               Conn1.Execute "DELETE tlbReceiptSplit.* " & _
                             "FROM tlbReceipt, tlbReceiptSplit " & _
                             "WHERE tlbReceiptSplit.RptHeader = tlbReceipt.TransactionID AND " & _
                                 "tlbReceipt.TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"

               Conn1.Execute "DELETE * " & _
                             "FROM tlbReceipt " & _
                             "WHERE TransactionID = " & Rst2.Fields.Item("TransactionID").Value & ";"
               Rst2.MoveNext
            Next i
            Rst2.Close
            Rst1.MoveNext
         Wend
      End If
   End If

   Rst1.Close

   GoTo ADD_EmailDmd_Tenants

ERROR_CheckData_Multiple_SI_tlbReceipt:
   MsgBox "System could not update your database. Please contact with PCM Consulting Ltd.", vbCritical + vbInformation + "Receipt Correction Error"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column EmailDmd on 14/02/2011 Tenants
'###############################################################################################################
ADD_EmailDmd_Tenants:
   On Error GoTo CHANGE_ADD_EmailDmd_Tenants

   Rst1.Open "SELECT EmailDmd FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_EmailSt_Tenants

CHANGE_ADD_EmailDmd_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailDmd BIT;"
   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare12;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column EmailSt on 31/03/2011 Tenants
'###############################################################################################################
ADD_EmailSt_Tenants:
   On Error GoTo CHANGE_ADD_EmailSt_Tenants

   Rst1.Open "SELECT EmailSt FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_CombEmail_Tenants

CHANGE_ADD_EmailSt_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailSt BIT;"
   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare11;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column CombEmail on 31/03/2011 Tenants
'###############################################################################################################
ADD_CombEmail_Tenants:
   On Error GoTo CHANGE_ADD_CombEmail_Tenants

   Rst1.Open "SELECT CombEmail FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_CombEmail_tlbDRCurrentPrint

CHANGE_ADD_CombEmail_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN CombEmail BIT;"
   Conn1.Execute "ALTER TABLE Tenants DROP COLUMN spare10;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column CombEmail on 01/04/2011 tlbDRCurrentPrint
'###############################################################################################################
ADD_CombEmail_tlbDRCurrentPrint:
   On Error GoTo CHANGE_ADD_CombEmail_tlbDRCurrentPrint

   Rst1.Open "SELECT CombEmail FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_UName_ShoppingCentre

CHANGE_ADD_CombEmail_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN CombEmail BIT;"
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint DROP COLUMN spare11;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column UName on 05/04/2011 ShoppingCentre
'###############################################################################################################
ADD_UName_ShoppingCentre:
   On Error GoTo CHANGE_ADD_UName_ShoppingCentre

   Rst1.Open "SELECT UName FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo RESIZE_SMTP_ShoppingCentre

CHANGE_ADD_UName_ShoppingCentre:
   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN UName TEXT(100);"
   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN Pws TEXT(50);"
   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN Port TEXT(5);"
   Conn1.Execute "ALTER TABLE ShoppingCentre DROP COLUMN Field4;"
   Conn1.Execute "ALTER TABLE ShoppingCentre DROP COLUMN Field5;"
   UpdateDatabase1 = 1
   Exit Function

'   Extend the field size SMTP on 05/04/2011 ShoppingCentre
'###############################################################################################################
RESIZE_SMTP_ShoppingCentre:

   On Error GoTo MissingTable_RESIZE_SMTP_ShoppingCentre

   Rst1.Open "SELECT SMTP FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("SMTP").DefinedSize = 15 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN SMTP TEXT(100)"
   End If
   Rst1.Close
   Set Rst1 = Nothing

   GoTo ADDNEW_COL_RecoverablePt_tblPurInvSRec

MissingTable_RESIZE_SMTP_ShoppingCentre:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - SMTP of DSR"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column RecoverablePt on 24/05/2011 tblPurInvSRec
'###############################################################################################################
ADDNEW_COL_RecoverablePt_tblPurInvSRec:
   On Error GoTo MISSING_ADDNEW_COL_RecoverablePt_tblPurInvSRec

   Rst1.Open "SELECT RecoverablePt FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_NEW_TABLE_Fund

MISSING_ADDNEW_COL_RecoverablePt_tblPurInvSRec:
   Conn1.Execute "ALTER TABLE tblPurInvSRec ADD COLUMN RecoverablePt Single;"
   Conn1.Execute "UPDATE tblPurInvSRec SET RecoverablePt = 100;"
   UpdateDatabase1 = 1
   Exit Function

'   New table on 07/06/XXXX Fund
'###############################################################################################################
ADD_NEW_TABLE_Fund:
   On Error GoTo CHANGE_ADD_NEW_TABLE_Fund

   Rst1.Open "SELECT * FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo ADDNEW_COL_CategoryCode_Fund

CHANGE_ADD_NEW_TABLE_Fund:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - Fund"
   UpdateDatabase1 = -1
   Exit Function

'   Add new column CategoryCode on 01/07/2011 Fund
'###############################################################################################################
ADDNEW_COL_CategoryCode_Fund:
   On Error GoTo MISSING_ADDNEW_COL_CategoryCode_Fund

   Rst1.Open "SELECT CategoryCode FROM Fund;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_Recoverable_tlbPayment

MISSING_ADDNEW_COL_CategoryCode_Fund:
   Conn1.Execute "ALTER TABLE Fund ADD COLUMN CategoryCode BYTE;"
   Conn1.Execute "UPDATE Fund SET CategoryCode = 1  WHERE INSTR(FundName, 'Rent')>0;"
   Conn1.Execute "UPDATE Fund SET CategoryCode = 2  WHERE INSTR(FundName, 'SERVICE')>0;"
   Conn1.Execute "UPDATE Fund SET CategoryCode = 3  WHERE INSTR(FundName, 'INSURANCE')>0;"
   Conn1.Execute "UPDATE Fund SET CategoryCode = 4  WHERE ISNULL(CategoryCode);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column Recoverable on 04/07/11 tlbPayment
'###############################################################################################################
ADDNEW_COL_Recoverable_tlbPayment:
   On Error GoTo MISSING_ADDNEW_COL_Recoverable_tlbPayment

   Rst1.Open "SELECT Recoverable FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo TypeIE_4_NominalLedger
MISSING_ADDNEW_COL_Recoverable_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN Recoverable SINGLE;"
   Conn1.Execute "UPDATE tlbPayment SET Recoverable = 0;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column TypeIE on 23/08/11 NominalLedger
'###############################################################################################################
TypeIE_4_NominalLedger:
   On Error GoTo ERROR_TypeIE_4_NominalLedger

   Rst1.Open "SELECT TypeIE FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_REC_IE
ERROR_TypeIE_4_NominalLedger:
   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN TypeIE TEXT(2);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new record on 23/08/11 PrimarySecondaryCode
'###############################################################################################################
ADDNEW_REC_IE:

   On Error GoTo MissingRec_REC_IE

   With Rst1

      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'IE';", Conn1, adOpenStatic, adLockReadOnly

      If .EOF Then
         .Close
         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         !Code = "IE"
         !Value = "INC_EXP"
         .Update
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item(0).Value = "IE"
         .Fields.Item(1).Value = "INC"
         .Fields.Item(2).Value = "Income"
         .Fields.Item(3).Value = "Nominal Ledger Type - Income"
         .Update
         .AddNew
         .Fields.Item(0).Value = "IE"
         .Fields.Item(1).Value = "EXP"
         .Fields.Item(2).Value = "Expenditure"
         .Fields.Item(3).Value = "Nominal Ledger Type - Expenditure"
         .Update
         .Close
      End If
      .Close

      .Open "SELECT Code FROM SecondaryCode WHERE PrimaryCode = 'IE' AND Value = 'Balance Sheet';", Conn1, adOpenStatic, adLockReadOnly

      If .EOF Then
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item(0).Value = "IE"
         .Fields.Item(1).Value = "BS"
         .Fields.Item(2).Value = "Balance Sheet"
         .Fields.Item(3).Value = "Nominal Ledger Type - Balance Sheet"
         .Update
      End If
      .Close
   End With

   GoTo MODIFY_DEPT_ID_tblPurInvSRec

MissingRec_REC_IE:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase1 = 1
   Exit Function

'   Modify DATA TYPE DEPT_ID on 24/10/07 tblPurInvSRec
'###############################################################################################################
MODIFY_DEPT_ID_tblPurInvSRec:

   On Error GoTo CHANGE_MODIFY_DEPT_ID_tblPurInvSRec

   Rst1.Open "SELECT DEPT_ID FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = 3 Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE tblPurInvSRec ALTER COLUMN DEPT_ID LONG;"
      GoTo CHANGE_MODIFY_DEPT_ID_tblPurInvSRec
   End If
   'If Month(Date) = 6 Then 'added by anol 2019-06-12 issue 785
    'GoTo CHECK_DATA_CORRUPTION_RECEIPT
   'Else
    GoTo CHECK_DATA_CORRUPTION_DEMAND
   'End If
CHANGE_MODIFY_DEPT_ID_tblPurInvSRec:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase1 = 1
   Exit Function

'   Check Data corruption in Demand & Receipt table on 10/02/2011
'NOTE:
'     SYSTEM INVESTIGATES THE DATA CORRUPTION.
'     CAUSES - SYSTEM COULD NOT UPDATE SI IN THE RECEIPT TABLE IF THE SI WAS MODIFIED.
'     SOLUTION - WHEN USER WILL GET THIS MSG -> MEANS THERE ARE SOME CORRUPTION.
'                1. DELETE THE SI FROM THE RECEIPT TABLE
'                2. UPDATE THE TrfReceipt = FALSE IN THE DEMAND SPLIT TABLE
'###############################################################################################################
CHECK_DATA_CORRUPTION_DEMAND:
   szSQL = "SELECT R.TransactionID, R.Amount, R.DemandRef, S.DT " & _
             "FROM tlbReceipt AS R, " & _
                 "(SELECT D.DemandID,  SUM(S.TotalAmount) AS DT " & _
                  "FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
                  "GROUP BY D.DemandID) AS S " & _
             "WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND " & _
                  "ROUND(R.Amount, 2) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 2);"

   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close

      GoTo CHECK_DATA_CORRUPTION_RECEIPT 'CHECK_BACS_EMAIL_tlbClientBanks
   End If
'   Debug.Print szSQL
   MsgBox "DATA ERROR: PCM_1254" & Chr(13) & "PLEASE CONTACT WITH PCM SUPPORT.", vbCritical + vbOKOnly, "PRESTIGE SYSTEM"
'   Conn1.Execute "Delete from tlbReceipt where transactionid=" & Rst1("TransactionID").Value & ""
'   Conn1.Execute "Delete from tlbReceiptSplit where RptHeader=" & Rst1("TransactionID").Value & ""
'   Conn1.Execute "Update DemandSplitRecords set TrfReceipt = FALSE  where DemandID=" & Rst1("DemandRef").Value & ""
'   MigrateInvIntoReceipt Conn1
   Rst1.Close
   UpdateDatabase1 = -1
   Exit Function

'  Check Data corruption in the Receipt Split table.
'     Data corrupt when user saves a receipt against a multiple SI splits.
'     System saved duplicated SR splits when there is multiple line of SI.
'Solution:
'  1. Collect the list of TransactionID which are corrupted.
'  2. Remove all splits from RptSplits.
'  3. Get the list of allocation splits from RptTransaction
'  4. Create all splits in tlbReceiptSplit according to RptTransaction's splits
'
'###############################################################################################################
CHECK_DATA_CORRUPTION_RECEIPT:
'  Check is there any transaction in the receipt table with 0 value in the header but non-0 value in the split.
'     if found then make the split value to 0. otherwise divident by zero will arise.
   szSQL = "SELECT R.TransactionID " & _
           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
                                  "From tlbReceiptSplit " & _
                                  "GROUP BY RptHeader " & _
                                 ") AS Q " & _
           "WHERE R.TransactionID=Q.RptHeader AND " & _
                 "Q.A > 0 AND R.Amount=0;"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   While Not Rst1.EOF
      Conn1.Execute "UPDATE tlbReceiptSplit " & _
                    "SET    Amount = 0, OSAmount = 0 " & _
                    "WHERE  RptHeader = " & Rst1.Fields.Item(0).Value & ";"
      Rst1.MoveNext
   Wend

   Rst1.Close

''  Update SI split in tlbReceiptSplit according to DemandSplit where they dont match and they has not paid full/part.
'   szSQL = "UPDATE tlbReceiptSplit AS RS, tlbReceipt AS R, " & _
'                  "DemandRecords as D, DemandSplitRecords as DS " & _
'           "SET    RS.Amount = DS.TotalAmount, RS.OSAmount = DS.TotalAmount " & _
'           "WHERE  D.DemandID = DS.DemandID AND " & _
'                  "R.TransactionID = RS.RptHeader AND " & _
'                  "R.DemandRef = D.DemandID AND " & _
'                  "DS.TotalAmount = RS.OSAmount AND " & _
'                  "RS.Amount <> DS.TotalAmount AND " & _
'                  "DS.SplitID = RS.SplitID;"
''Debug.Print szSQL
'   Conn1.Execute szSQL
'
''  Update SI split in tlbReceiptSplit according to DemandSplit where they dont match and they has not paid full/part.
'   szSQL = "UPDATE tlbReceiptSplit AS RS, tlbReceipt AS R, " & _
'                  "DemandRecords as D, DemandSplitRecords as DS " & _
'           "SET    RS.Amount = DS.TotalAmount, RS.OSAmount = DS.TotalAmount " & _
'           "WHERE  D.DemandID = DS.DemandID AND " & _
'                  "R.TransactionID = RS.RptHeader AND " & _
'                  "R.DemandRef = D.DemandID AND " & _
'                  "RS.Amount = RS.OSAmount AND " & _
'                  "RS.Amount <> DS.TotalAmount AND " & _
'                  "DS.SplitID = RS.SplitID;"
'Debug.Print szSQL
'   Conn1.Execute szSQL
'
'Sol 1: Collect the list of TransactionID which are corrupted.
   szSQL = "SELECT R.TransactionID, Q.A-R.Amount AS D, Q.A / R.Amount AS R " & _
           "FROM tlbReceipt AS R, (SELECT RptHeader, SUM(Amount) AS A " & _
                                  "From tlbReceiptSplit " & _
                                  "GROUP BY RptHeader " & _
                                 ") AS Q " & _
           "WHERE R.TransactionID = Q.RptHeader AND " & _
                 "ROUND(R.Amount, 2) <> ROUND(Q.A, 2) AND " & _
                 "R.Amount+Q.A>0 AND R.OSAmount + Q.A <> R.Amount;"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      GoTo CHECK_DEMANDS_MISSED_MIGRATION
   End If

   iRec = Rst1.RecordCount
   i = 0
   While Not Rst1.EOF
      If Val(Rst1.Fields.Item("D").Value) > 0 And _
            CInt(Rst1.Fields.Item("R").Value) = Val(Rst1.Fields.Item("R").Value) Then
            Conn1.Execute "DELETE * FROM tlbReceiptSplit " & _
                       "WHERE RptHeader = " & Rst1.Fields.Item("TransactionID").Value & " AND " & _
                           "SplitID > 1;"
         i = i + 1
      End If
      Rst1.MoveNext
   Wend
   Rst1.Close

   If iRec > 0 Then
'      Conn1.Execute "DELETE * " & _
'                    "FROM tlbReceiptSplit AS S " & _
'                    "WHERE S.RptHeader NOT IN " & _
'                         "(SELECT TransactionID FROM  tlbReceipt AS R)"
'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
      Conn1.Execute "DELETE S.*  FROM tlbReceiptSplit AS S LEFT JOIN tlbReceipt AS R ON  S.RptHeader = R.TransactionID  WHERE R.TransactionID is NULL"
   End If
   If iRec <> i Then
      MsgBox "Warning: This Company data need to be updated. Please contact with PCM.", vbExclamation + vbOKOnly, "Data Update"
      GoTo CHECK_DEMANDS_MISSED_MIGRATION
   End If

'   While Not Rst1.EOF
'      If Rst1.Fields.Item("D").Value > 0 Then
''Sol  2. Remove all splits from RptSplits.
'         Conn1.Execute "DELETE * " & _
'                       "FROM   tlbReceiptSplit " & _
'                       "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";"
'
''Sol  3. Get the list of allocation splits from RptTransaction
'         szSQL = "SELECT * " & _
'                 "FROM   RptTransactions " & _
'                 "WHERE  FromTran = " & Rst1.Fields.Item("TransactionID").Value & ";"
'         Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
''Sol  4. Create all splits in tlbReceiptSplit according to RptTransaction's splits
'         Rst3.Open "SELECT * FROM tlbReceiptSplit", Conn1, adOpenDynamic, adLockOptimistic
'         While Not Rst2.EOF
'            With Rst3
'               .AddNew
'               .Fields.Item("TransactionID").Value = UniqueID()
'               .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
'''#
''               .Fields.Item("FundID").Value = flxSPayment.TextMatrix(iRow, 14)
''               .Fields.Item("Amount").Value = flxSPayment.TextMatrix(iRow, 10)
''               If flxSPayment.TextMatrix(iRow, 0) = "" Then
''                  .Fields.Item("SplitID").Value = -1
''               Else
''                  .Fields.Item("SplitID").Value = flxSPayment.TextMatrix(iRow, 1)
''               End If
''               .Fields.Item("DueDate").Value = Format(Now, "dd mmmm yyyy")
''               .Fields.Item("Description").Value = flxSPayment.TextMatrix(iRow, 7)
''               .Fields.Item("AllocTranID").Value = flxSPayment.TextMatrix(iRow, 19)
'               .Update
'            End With
'
'
'            Rst2.MoveNext
'         Wend
'         Rst2.Close
'      End If
'      If Rst1.Fields.Item("D").Value < 0 Then
'      End If
'      Rst1.MoveNext
'   Wend


'  MigrateInvIntoReceipt method missed some demands to export to Receipt table.
'  This procedure will reset the flag to export to receipt.
'  When user will open the demand form, MigrateInvIntoReceipt method will export them again.
'  THIS MATHOD HAS TO BE RUN EVERY TIME WHEN SYSTEM OPENS.
'###############################################################################################################
CHECK_DEMANDS_MISSED_MIGRATION:


'  Code has been removed to DemandMigration function

   GoTo CHECK_BACS_EMAIL_tlbClientBanks
'  BACS From email address setting in the BACS form
'NOTE:
'     IF EMAIL IN THE tlbClientBanks = NULL AND Email1 IN ShoppingCentre <> NULL THEN
'     SYSTEM WILL COPY THE Email1 INTO EMAIL OF tlbClientBanks
'###############################################################################################################
CHECK_BACS_EMAIL_tlbClientBanks:
   szSQL = "SELECT Email1 " & _
           "FROM   ShoppingCentre " & _
           "WHERE  Email1 <> '';"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount = 0 Then GoTo NO_ACCTION

   szEmail = Rst1.Fields.Item("Email1").Value
   Rst1.Close
   szSQL = "SELECT email " & _
             "FROM   tlbClientBanks " & _
             "WHERE  email = '' OR ISNULL(email);"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount = 0 Then GoTo NO_ACCTION
   Rst1.Close

   Conn1.Execute "UPDATE tlbClientBanks " & _
                 "SET    email = '" & szEmail & "' " & _
                 "WHERE  email = '' OR ISNULL(email);"
   GoTo ADDNEW_COL_SentByEmail_DemandRecords

NO_ACCTION:
   Rst1.Close
   GoTo ADDNEW_COL_SentByEmail_DemandRecords

'   Add new column SentByEmail on 16/09/2011 DemandRecords
'###############################################################################################################
ADDNEW_COL_SentByEmail_DemandRecords:
   On Error GoTo MISSING_ADDNEW_COL_SentByEmail_DemandRecords

   Rst1.Open "SELECT SentByEmail FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_CT_tlbBankPayment

MISSING_ADDNEW_COL_SentByEmail_DemandRecords:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN SentByEmail BYTE;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column CT on 27/09/2011 tlbBankPayment
'###############################################################################################################
ADDNEW_COL_CT_tlbBankPayment:
   On Error GoTo MISSING_ADDNEW_COL_CT_tlbBankPayment

   Rst1.Open "SELECT CT FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo FIX_OS_tlbReceiptSplit

MISSING_ADDNEW_COL_CT_tlbBankPayment:
   Conn1.Execute "ALTER TABLE tlbBankPayment DROP COLUMN spare7;"
   Conn1.Execute "ALTER TABLE tlbBankPayment DROP COLUMN spare8;"
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD  COLUMN CT TEXT(1);"

   UpdateBR_BP_SameAcc

   UpdateDatabase1 = 1
   Exit Function

'   Fix OS balance in the receipt_split on 06/10/2011
'   There is an inconsistency found in the total OS of splits and the OS of the receipt header.
'###############################################################################################################
FIX_OS_tlbReceiptSplit:
   On Error GoTo MISSING_FIX_OS_tlbReceiptSplit

   Rst1.Open "SELECT R.TransactionID, R.OSAmount AS R_OS, S.S_OS " & _
             "FROM tlbReceipt AS R, ( " & _
                     "SELECT S.RptHeader, SUM(S.OSAmount) AS S_OS " & _
                     "FROM tlbReceiptSplit AS S " & _
                     "GROUP BY S.RptHeader " & _
                     ") AS S " & _
             "WHERE R.TransactionID =  S.RptHeader AND " & _
                     "ROUND(R.OSAmount, 2) <> ROUND(S.S_OS, 2) " & _
             "ORDER BY S.RptHeader;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount > 0 Then        'need to fix data. inconsistent data found
      While Not Rst1.EOF
         If Rst1.Fields.Item("R_OS").Value = 0 Then
            Conn1.Execute "UPDATE tlbReceiptSplit " & _
                          "SET    OSAmount = 0 " & _
                          "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";"
         End If
         If Rst1.Fields.Item("R_OS").Value > 0 Then
            cOS = Round(CCur(Rst1.Fields.Item("R_OS").Value), 2)
            Rst2.Open "SELECT * FROM tlbReceiptSplit " & _
                      "WHERE  RptHeader = " & Rst1.Fields.Item("TransactionID").Value & ";", _
                      Conn1, adOpenDynamic, adLockOptimistic
            While Not Rst2.EOF
               If cOS = 0 Then
                  Rst2.Fields.Item("OSAmount").Value = 0
               End If
               If cOS <= Round(CCur(Rst2.Fields.Item("Amount").Value), 2) And cOS > 0 Then
                  Rst2.Fields.Item("OSAmount").Value = cOS
                  cOS = 0
               End If
               If cOS > Round(CCur(Rst2.Fields.Item("Amount").Value), 2) Then
                  Rst2.Fields.Item("OSAmount").Value = Rst2.Fields.Item("Amount").Value
                  cOS = cOS - Rst2.Fields.Item("Amount").Value
               End If

               Rst2.MoveNext
            Wend
            Rst2.Close
         End If

         Rst1.MoveNext
      Wend
   End If

   Rst1.Close

   GoTo FIX_FUND_tlbBankPayment

MISSING_FIX_OS_tlbReceiptSplit:
   UpdateDatabase1 = 1
   Exit Function

'   Check the system Bank Transactions which are without fund.
'   If there is any transaction found, then system will update the Fund = 1
'###############################################################################################################
FIX_FUND_tlbBankPayment:

   szSQL = "UPDATE tlbBankPayment " & _
           "SET DEPT_ID = 1 " & _
           "WHERE DEPT_ID = '' OR ISNULL(DEPT_ID);"
   Conn1.Execute szSQL

'   Add new column PrintBBF on 20/10/2011 tlbDRCurrentPrint
'###############################################################################################################
ADDNEW_COL_PrintBBF_tlbDRCurrentPrint:
   On Error GoTo MISSING_ADDNEW_COL_PrintBBF_tlbDRCurrentPrint

   Rst1.Open "SELECT PrintBBF FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo FIX_EXPORT_tblpurinv

MISSING_ADDNEW_COL_PrintBBF_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN PrintBBF BIT;"
   UpdateDatabase1 = 1
   Exit Function

'   Check the system PI Transactions which have not exported to payment table.
'   If there is any transaction found, then system will export them
'###############################################################################################################
FIX_EXPORT_tblpurinv:
'   szSQL = "SELECT MY_ID " & _
'           "FROM tblPurInv " & _
'           "WHERE TransactionType <> 25 AND MY_ID NOT IN (SELECT PI AS MY_ID FROM tlbPayment WHERE PI <> '') AND " & _
'               "TOTAL_AMOUNT <> 0;"

'issue  381 SQL that is taking long time to complete 20170512 fixed by anol

  szSQL = "SELECT MY_ID FROM tblPurInv P LEFT JOIN tlbPayment M  ON P.MY_ID = M.PI WHERE P.TransactionType <> 25 AND P.TOTAL_AMOUNT <> 0 AND M.PI is NULL"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      MsgBox "Your data need to be updated. Please contact with PCM", vbCritical + vbOKOnly, "PI not in Split table"
      Rst1.Close
      UpdateDatabase1 = -1

' MigratePIIntoPayment method export only PI's header. MigratePIIntoPayment does not create the split in the
'  payment split table. Technically, there should not any PI in the purchase invoice table. because, when
'  users create a PI, system immidiately exports the transaction to the PP and PP_split table automatically.
'08/08/2012
      Exit Function

'      szSQL = "UPDATE tblPurInv " & _
'              "SET    TrfPayment = FALSE " & _
'              "WHERE  MY_ID NOT IN (SELECT PI AS MY_ID FROM tlbPayment WHERE PI <> '') AND " & _
'                  "TOTAL_AMOUNT <> 0;"
   'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
     szSQL = "UPDATE tblPurInv P LEFT JOIN tlbPayment M ON P.MY_ID = M.PI   SET   P.TrfPayment = FALSE WHERE    TOTAL_AMOUNT <> 0 AND  PI is NULL;"
'Debug.Print szSQL
      Conn1.Execute szSQL

      MigratePIIntoPayment Conn1
   End If
   Rst1.Close

'   Add a column ULC on 27/10/2011 DemandTypes
'###############################################################################################################
ADD_ULC_DemandTypes:
   On Error GoTo CHANGE_ADD_ULC_DemandTypes

   Rst1.Open "SELECT ULC FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo FIX_DEMANDTYPE_LLEASE

CHANGE_ADD_ULC_DemandTypes:
   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN ULC BIT;"
   UpdateDatabase1 = 1
   Exit Function

'   Check the system for wrong demand type setup in the lease.
'   If there is any transaction found, then system will warn user and if user wants
'   they will able to print the list of lease need to be fix manually.
'   27/10/2011
'   The reports are generating only DemandType and Lease tables entry.
'   But still system can generating this warning by checking the demand and split table
'   Please check szSQL___
'   08/01/2014
'###############################################################################################################
FIX_DEMANDTYPE_LLEASE:
   On Error GoTo BUG_FIX_DEMANDTYPE_LLEASE

   If bFixingDT Then GoTo MODIFY_TypeIE_NominalLedger

   szSQL = "SELECT ID " & _
           "FROM (" & _
               "SELECT R.BRDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
               "FROM DemandTypes AS T, LRentCharges AS R, LeaseDetails AS L, Units AS U " & _
               "Where T.ID = R.BRDemandType And R.LeaseID = L.LeaseID And " & _
                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
               "GROUP BY R.BRDemandType, T.PropertyID, U.PropertyID " & _
           ") AS Q " & _
           "Where PropertyID <> LeaseProperty"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   szSQL_ = "SELECT ID " & _
           "FROM (" & _
               "SELECT R.SCDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
               "FROM DemandTypes AS T, LServiceCharges AS R, LeaseDetails AS L, Units AS U " & _
               "Where T.ID = R.SCDemandType And R.LeaseID = L.LeaseID And " & _
                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
               "GROUP BY R.SCDemandType, T.PropertyID, U.PropertyID " & _
           ") AS Q " & _
           "WHERE PropertyID <> LeaseProperty"
'Debug.Print szSQL
   Rst2.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   szSQL__ = "SELECT ID " & _
           "FROM (" & _
               "SELECT R.InsuranceDemandType AS ID, T.PropertyID, U.PropertyID AS LeaseProperty " & _
               "FROM DemandTypes AS T, LInsuranceCharges AS R, LeaseDetails AS L, Units AS U " & _
               "Where T.ID = R.InsuranceDemandType And R.LeaseID = L.LeaseID And " & _
                     "L.UnitNumber = U.UnitNumber AND L.Status " & _
               "GROUP BY R.InsuranceDemandType, T.PropertyID, U.PropertyID " & _
           ") AS Q " & _
           "WHERE PropertyID <> LeaseProperty"
'Debug.Print szSQL
   Rst3.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
' check other charges
   szSQL___ = "SELECT ID, P1, P2 " & _
           "FROM (" & _
               "SELECT S.TypeOfDemand AS ID, U.PropertyID AS P1, DT.PropertyID AS P2 " & _
               "FROM DemandRecords AS D, DemandSplitRecords AS S, Tenants AS T, LeaseDetails AS L, Units AS U, DemandTypes AS DT " & _
               "WHERE L.Status AND D.DemandID = S.DemandID AND D.SageAccountNumber = T.SageAccountNumber AND " & _
                     "T.SageAccountNumber = L.SageAccountNumber AND L.UnitNumber = U.UnitNumber AND " & _
                     "S.TypeOfDemand = DT.ID " & _
               "GROUP BY S.TypeOfDemand, U.PropertyID, DT.PropertyID " & _
           ") AS Q " & _
           "WHERE Q.P1 <> Q.P2"
'Debug.Print szSQL
   Rst4.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Or Not Rst2.EOF Or Not Rst3.EOF Or Not Rst4.EOF Then
      If MsgBox("There are some lease demand type need to fix." & _
                "Do you want to do it now?", vbCritical + vbYesNo, "Demand Types") = vbNo Then

         Rst1.Close
         Rst2.Close
         Rst3.Close
         Rst4.Close
         GoTo MODIFY_TypeIE_NominalLedger
      End If
'      FIX_MODE__DT = True

      szSQL = "UPDATE DemandTypes " & _
              "SET ULC = TRUE " & _
              "WHERE ID IN (" & szSQL & ");"
'Debug.Print szSQL
      Conn1.Execute szSQL
      szSQL = "UPDATE DemandTypes " & _
              "SET ULC = TRUE " & _
              "WHERE ID IN (" & szSQL_ & ");"
'Debug.Print szSQL
      Conn1.Execute szSQL
      szSQL = "UPDATE DemandTypes " & _
              "SET ULC = TRUE " & _
              "WHERE ID IN (" & szSQL__ & ");"
'Debug.Print szSQL
      Conn1.Execute szSQL
      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_RC.rpt"
      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_SC.rpt"
      ShowReport App.Path & "\CompanyReports\FixDemandTypeInLease_IC.rpt"

      bFixingDT = True
   Else
      szSQL = "UPDATE DemandTypes " & _
              "SET ULC = TRUE " & _
              "WHERE ULC = FALSE;"
      Conn1.Execute szSQL
   End If
   Rst1.Close
   Rst2.Close
   Rst3.Close
   Rst4.Close

   GoTo MODIFY_TypeIE_NominalLedger

BUG_FIX_DEMANDTYPE_LLEASE:
   MsgBox "System could not fix your data. Please contact with PCM.", vbInformation + vbOKOnly, "Demand Types"
   UpdateDatabase1 = 1
   Exit Function
'
''   Modify DATA TYPE NLTypeCode on 01/11/11 NLType
''###############################################################################################################
'MODIFY_NLTypeCode_NLType:
'   On Error GoTo CHANGE_MODIFY_NLTypeCode_NLType
'
'   Rst1.Open "SELECT NLTypeCode FROM NLType;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.Fields(0).Type = 202 Then
'      Rst1.Close
'   Else
'      Rst1.Close
'      Conn1.Execute "ALTER TABLE NominalLedger DROP CONSTRAINT NLTypeNominalLedger;"
'      Conn1.Execute "ALTER TABLE NLType ALTER COLUMN NLTypeCode TEXT(50);"
'      Conn1.Execute "ALTER TABLE NominalLedger ALTER COLUMN Type TEXT(50);"
'   End If
'
'   GoTo MODIFY_TypeIE_NominalLedger
'
'CHANGE_MODIFY_NLTypeCode_NLType:
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
'   UpdateDatabase1 = 1
'   Exit Function

'   Modify DATA TYPE TypeIE on 01/11/11 NominalLedger
'###############################################################################################################
MODIFY_TypeIE_NominalLedger:

   On Error GoTo CHANGE_MODIFY_TypeIE_NominalLedger

   Rst1.Open "SELECT TypeIE FROM NominalLedger WHERE TypeIE IN ('1', '2', '3');", Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'IN' WHERE TypeIE = '1';"
      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'EX' WHERE TypeIE = '2';"
      Conn1.Execute "UPDATE NominalLedger SET TypeIE = 'BS' WHERE TypeIE = '3';"
   End If
   Rst1.Close

   GoTo MODIFY_Code_SecondaryCode

CHANGE_MODIFY_TypeIE_NominalLedger:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase1 = 1
   Exit Function

'   Modify DATA Code on 01/11/11 SecondaryCode
'###############################################################################################################
MODIFY_Code_SecondaryCode:
   On Error GoTo CHANGE_MODIFY_Code_SecondaryCode

   Rst1.Open "SELECT Code FROM SecondaryCode WHERE Code IN ('1', '2', '3');", Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Conn1.Execute "UPDATE SecondaryCode SET Code = 'IN' WHERE Code = 'INC' AND PrimaryCode = 'IE';"
      Conn1.Execute "UPDATE SecondaryCode SET Code = 'EX' WHERE Code = 'EXP' AND PrimaryCode = 'IE';"
   End If
   Rst1.Close

   GoTo ADD_EmailSC_Tenants

CHANGE_MODIFY_Code_SecondaryCode:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column EmailSC on 16/11/2011 Tenants
'###############################################################################################################
ADD_EmailSC_Tenants:
   On Error GoTo CHANGE_ADD_EmailSC_Tenants

   Rst1.Open "SELECT EmailSC FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_EmailSC_tlbDRCurrentPrint

CHANGE_ADD_EmailSC_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN EmailSC BIT;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column EmailSC on 28/11/2011 tlbDRCurrentPrint
'###############################################################################################################
ADD_EmailSC_tlbDRCurrentPrint:
   On Error GoTo CHANGE_ADD_EmailSC_tlbDRCurrentPrint

   Rst1.Open "SELECT EmailSC FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_SentStName_DemandRecords

CHANGE_ADD_EmailSC_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN EmailSC BIT;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column SentStName on 09/12/2011 DemandRecords
'###############################################################################################################
ADD_SentStName_DemandRecords:
   On Error GoTo CHANGE_ADD_SentStName_DemandRecords

   Rst1.Open "SELECT SentStName FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_ReportedBy_PropertyMaintHistory

CHANGE_ADD_SentStName_DemandRecords:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN SentStName TEXT(100);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new columns ReportedBy.. on 15/12/2011 PropertyMaintHistory
'###############################################################################################################
ADD_ReportedBy_PropertyMaintHistory:
   On Error GoTo CHANGE_ADD_ReportedBy_PropertyMaintHistory

   Rst1.Open "SELECT ReportedBy FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo LeaseDetails_Lessee_Duplicate

CHANGE_ADD_ReportedBy_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedBy TEXT(50);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN AssignedIL TEXT(1);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedIS TEXT(1);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Urgent TEXT(1);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Instruction MEMO;"
   UpdateDatabase1 = 1
   Exit Function
'***************************************************************************************************************
'                  Check System data - System will check lease details table for multiple lessee's account     '
'                  If system found any lessee more than one time in the lease table as active leases           '
'                  then system will stop user to use the system and push them to report it to us               '
'                                                                                                              '
'                                               21/12/2011                                                     '
'###############################################################################################################
LeaseDetails_Lessee_Duplicate:
   Rst1.Open "SELECT * " & _
             "From " & _
             "( " & _
              "SELECT COUNT(SageAccountNumber) AS A, SageAccountNumber " & _
              "From LeaseDetails " & _
              "Where Status " & _
              "GROUP BY SageAccountNumber " & _
              ") AS Q " & _
             "Where Q.a > 1;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      GoTo ADD_TemplateID_tlbLetterReports
   Else
      MsgBox "The system found an inconsistency in your database. Please contact PCM Consulting Support.", vbCritical + vbOKOnly, "Err. Multiple Lessee"
      Rst1.Close
      UpdateDatabase1 = -1
      Exit Function
   End If

'   Add new column TemplateID on 21/12/2011 tlbLetterReports
'###############################################################################################################
ADD_TemplateID_tlbLetterReports:
   On Error GoTo CHANGE_ADD_TemplateID_tlbLetterReports

   Rst1.Open "SELECT TemplateID FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_REC_TEMP_TYPE

CHANGE_ADD_TemplateID_tlbLetterReports:
   Conn1.Execute "ALTER TABLE tlbLetterReports ADD COLUMN TemplateID LONG;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new record TEMP_TYPE on 22/12/11 PrimaryCode
'###############################################################################################################
ADDNEW_REC_TEMP_TYPE:
   On Error GoTo MissingData_ADDNEW_REC_TEMP_TYPE

   With Rst1
      .Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'TEMP_TYPE';", Conn1, adOpenStatic, adLockReadOnly

      If .EOF Then
         .Close
         .Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item("Code").Value = "TEMP_TYPE"
         .Fields.Item("Value").Value = "LETTER TEMPLATE TYPE"
         .Fields.Item("Flexible").Value = False
         .Update
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
         .Fields.Item("Code").Value = "LT"
         .Fields.Item("Value").Value = "Letter Template"
         .Update
         .AddNew
         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
         .Fields.Item("Code").Value = "RT"
         .Fields.Item("Value").Value = "Reminder Template"
         .Update
         .AddNew
         .Fields.Item("PrimaryCode").Value = "TEMP_TYPE"
         .Fields.Item("Code").Value = "OT"
         .Fields.Item("Value").Value = "Other Template"
         .Update
      End If
      .Close
   End With

   GoTo ADD_TempType_Template

MissingData_ADDNEW_REC_TEMP_TYPE:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - TEMP_TYPE) - tlbReceipt"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column TempType on 22/12/2011 Template
'###############################################################################################################
ADD_TempType_Template:

   On Error GoTo CHANGE_ADD_TempType_Template

   Rst1.Open "SELECT TempType FROM Template;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_ReportedFrom_PropertyMaintHistory

CHANGE_ADD_TempType_Template:
   Conn1.Execute "ALTER TABLE Template ADD COLUMN TempType TEXT(10);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new columns ReportedFrom on 21/01/2012 PropertyMaintHistory
'###############################################################################################################
ADD_ReportedFrom_PropertyMaintHistory:
   On Error GoTo CHANGE_ADD_ReportedFrom_PropertyMaintHistory

   Rst1.Open "SELECT ReportedFrom FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_Fund_TenantDeposit

CHANGE_ADD_ReportedFrom_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ReportedFrom TEXT(1);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column FundID on 26/01/2012 TenantDeposit
'###############################################################################################################
ADD_Fund_TenantDeposit:
   On Error GoTo CHANGE_ADD_Fund_TenantDeposit

   Rst1.Open "SELECT FundID FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo MODIFY_RefundRef_TenantDeposit

CHANGE_ADD_Fund_TenantDeposit:
   Conn1.Execute "ALTER TABLE TenantDeposit ADD COLUMN FundID Long;"
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN TenantDeposit TEXT(50);"
   Conn1.Execute "ALTER TABLE TenantDeposit ALTER COLUMN DepositID TEXT(50);"
   UpdateDatabase1 = 1
   Exit Function

'   Modify DATA TYPE RefundRef on 26/01/2012 TenantDeposit
'###############################################################################################################
MODIFY_RefundRef_TenantDeposit:
   On Error GoTo CHANGE_MODIFY_RefundRef_TenantDeposit

   Rst1.Open "SELECT RefundRef FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = 202 Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE TenantDeposit ALTER COLUMN RefundRef TEXT(50);"
      GoTo CHANGE_MODIFY_RefundRef_TenantDeposit
   End If

   GoTo ALTER_TENANT_ADDRESS_SecondaryCode

CHANGE_MODIFY_RefundRef_TenantDeposit:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase1 = 1
   Exit Function

'   Altered Data TenantAddress on 23/12/2011 SecondaryCode
'###############################################################################################################
ALTER_TENANT_ADDRESS_SecondaryCode:

   Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'TENANT ADDRESS';", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'TENANT ADDRESS';", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.Fields.Item("Value").Value = "Lessee Address"
      Rst1.Update
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'INVADD' AND Value = 'ALTERNATIVE ADDRESS';", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.Fields.Item("Value").Value = "Alternative Address"
      Rst1.Update
   End If
   Rst1.Close

'   Add new column TYPE on 08/05/2013 Supplier
'###############################################################################################################
ADDNEW_COL_TYPE_Supplier:
   On Error GoTo MISSING_ADDNEW_COL_TYPE_Supplier

   Rst1.Open "SELECT TYPE FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo DUPLICATE_ID_LCLSA

MISSING_ADDNEW_COL_TYPE_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN TYPE TEXT(20);"
   Conn1.Execute "UPDATE Supplier SET TYPE = 'SUPPLIER';"
   UpdateDatabase1 = 1
   Exit Function

'   Check Duplicate ID on 12/01/2012 Lessee, Client, Landlord, Supplier & MA
'###############################################################################################################
DUPLICATE_ID_LCLSA:
   szSQL = "SELECT ID "
   szSQL = szSQL & "FROM (SELECT ID, COUNT(ID) AS C "
   szSQL = szSQL & "FROM ("
   szSQL = szSQL & "SELECT SupplierID AS ID "
   szSQL = szSQL & "FROM Supplier WHERE TYPE = 'SUPPLIER' UNION ALL "
   szSQL = szSQL & "SELECT ClientID AS ID "
   szSQL = szSQL & "FROM Client UNION ALL "
   szSQL = szSQL & "SELECT SageAccountNumber AS ID "
   szSQL = szSQL & "FROM Tenants UNION ALL "
   szSQL = szSQL & "SELECT AgentID AS ID "
   szSQL = szSQL & "FROM Agent UNION ALL "
   szSQL = szSQL & "SELECT LandlordID AS ID "
   szSQL = szSQL & "From Landlord "
   szSQL = szSQL & ") "
   szSQL = szSQL & " GROUP BY ID "
   szSQL = szSQL & ") "
   szSQL = szSQL & "WHERE C > 1;"

'Debug.Print szSQL

   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then                                  'Duplicate ID found
      Conn1.Execute "DELETE Tenants.* " & _
                    "FROM Tenants, Landlord " & _
                    "WHERE Tenants.SageAccountNumber = Landlord.LandlordID;"
      Rst1.Close
      Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
      If Not Rst1.EOF Then                                  'Duplicate ID found
         szSQL = SQL2String(Rst1, 0)
         Rst1.Close

         MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
      Else
         Rst1.Close
      End If
   Else
      Rst1.Close
   End If
'
''  Export Clients into Supplier table on 08/05/2013         ~~~NEVER REMOVED~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''###############################################################################################################
'   Rst1.Open "SELECT * " & _
'             "FROM Client " & _
'             "WHERE ClientID NOT IN (" & _
'               "SELECT SupplierID FROM Supplier WHERE TYPE = 'CLIENT');", Conn1, adOpenStatic, adLockReadOnly
'   If Not Rst1.EOF Then
'      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
'      While Not Rst1.EOF
'         Rst2.AddNew
'         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("ClientID").Value
'         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("ClientName").Value
'         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("ClientAddressLine1").Value
'         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("ClientAddressLine2").Value
'         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("ClientAddressLine3").Value
'         Rst2.Fields.Item("SupplierPostCode").Value = IIf(IsNull(Rst1.Fields.Item("ClientPostCode").Value), "", Rst1.Fields.Item("ClientPostCode").Value)
'         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
'         Rst2.Fields.Item("TYPE").Value = "CLIENT"
'         Rst2.Update
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close
'
''  Export Agents into Supplier table on 08/05/2013         ~~~NEVER REMOVED~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''###############################################################################################################
'   Rst1.Open "SELECT * " & _
'             "FROM Agent " & _
'             "WHERE AgentID NOT IN (" & _
'               "SELECT SupplierID FROM Supplier WHERE TYPE = 'AGENT');", Conn1, adOpenStatic, adLockReadOnly
'   If Not Rst1.EOF Then
'      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
'      While Not Rst1.EOF
'         Rst2.AddNew
'         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("AgentID").Value
'         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("AgentName").Value
'         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("AgentAddressLine1").Value
'         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("AgentAddressLine2").Value
'         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("AgentAddressLine3").Value
'         Rst2.Fields.Item("SupplierPostCode").Value = Rst1.Fields.Item("AgentPostCode").Value
'         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
'         Rst2.Fields.Item("TYPE").Value = "AGENT"
'         Rst2.Update
'         Rst1.MoveNext
'      Wend
'      Rst2.Close
'   End If
'   Rst1.Close

'   Add new column TEmail on 06/02/2012 tlbLetterReports
'###############################################################################################################
ADD_TEmail_tlbLetterReports:

   On Error GoTo CHANGE_ADD_TEmail_tlbLetterReports

   Rst1.Open "SELECT TEmail FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo RESIZE_SageSuppAC_Supplier

CHANGE_ADD_TEmail_tlbLetterReports:
   Conn1.Execute "ALTER TABLE tlbLetterReports ADD COLUMN TEmail TEXT(40);"
   UpdateDatabase1 = 1
   Exit Function

'   Extend the field size SageSuppAC on 29/02/2012 Supplier
'###############################################################################################################
RESIZE_SageSuppAC_Supplier:

   Rst1.Open "SELECT SageSuppAC FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("SageSuppAC").DefinedSize = 10 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SageSuppAC TEXT(50)"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   Extend the field size of Field8 on 29/02/2012 ShoppingCentre
'###############################################################################################################
   Rst1.Open "SELECT Field8 FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Field8").DefinedSize = 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN Field8 TEXT(255)"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   Add new record CT on 09/02/2012 Client
'###############################################################################################################
ADD_CT_Client:
   On Error GoTo MissingData_ADD_CT_Client

   Rst1.Open "SELECT CT FROM Client;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_FileLoc__tlbClientBanks
'GoTo UPDATE_AMOUNT_2_DECIMAL

MissingData_ADD_CT_Client:
   Conn1.Execute "ALTER TABLE Client ADD COLUMN CT TEXT(25);"
   Conn1.Execute "UPDATE Client SET CT = 'Property Management';"
   UpdatePI_ClientID
   UpdateDatabase1 = 1
   Exit Function

'   Add new column FileLoc_ on 06/03/2012 tlbClientBanks
'###############################################################################################################
ADD_FileLoc__tlbClientBanks:
   On Error GoTo CHANGE_ADD_FileLoc__tlbClientBanks

   Rst1.Open "SELECT FileLoc_ FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo tlbPaymentSplit

CHANGE_ADD_FileLoc__tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN FileLoc_ TEXT(255);"
   UpdateDatabase1 = 1
   Exit Function

'   New table on 14/02/2012 tlbPaymentSplit
'###############################################################################################################
tlbPaymentSplit:
   On Error GoTo MissingTable_tlbPaymentSplit

   Rst1.Open "SELECT * FROM tlbPaymentSplit;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount = 0 Then
'     We have to check here: is there any PP partially allocated?
'     If any PP found then user has to fix it before upgrade the system to PP split.
'     System will generate a report for user to fix PP.
      Rst2.Open "SELECT * FROM tlbPayment " & _
                "WHERE Type > 7 AND OSAmount > 0 AND " & _
                      "OSAmount < Amount;", Conn1, adOpenStatic, adLockReadOnly
      If Not Rst2.EOF Then
         Rst2.Close
         Rst1.Close
         ShowReport App.Path & szReportPath & "\PP_PartialAlloc.rpt"
         MsgBox "Please print this report and unallocate these transactions." & Chr(13) & _
                "After upgrade the system, re-allocate these transactions.", _
                vbInformation + vbOKOnly, "Partially Allocated Purchase Payment"

         GoTo FixDataByUser
      Else
         Rst2.Close
      End If

      If Not UpdateTlbPaymentSplit Then
         UpdateDatabase1 = -1
         Exit Function
      End If
   End If

   Rst1.Close

   GoTo ADDNEW_REC_LL

MissingTable_tlbPaymentSplit:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbPaymentSplit"
'Debug.Print ERR.description

   UpdateDatabase1 = -1
   Exit Function

'   Add new record Landlord on 09/03/2012 SecondaryCode
'###############################################################################################################
ADDNEW_REC_LL:
   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'SCODE' AND Code = 'LL';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "SCODE"
         !Code = "LL"
         !Value = "Landlord"
         .Update
      End With
   End If
   Rst1.Close

'   Add new column RAS on 25/04/2012 Property
'###############################################################################################################
ADDNEW_COL_RAS_Property:
   On Error GoTo MISSING_ADDNEW_COL_RAS_Property

   Rst1.Open "SELECT RAS FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo FIX_LeaseRef_DemandRecords

MISSING_ADDNEW_COL_RAS_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN RAS TEXT(1);"
   UpdateDatabase1 = 1
   Exit Function

'  Check DemandRecords table 'LeaseRef' column.
'  If 'LeaseRef' is empty then system will update according LeaseDetails table.
'  If the lease is expired then system will update 'LeaseRef' with latest expired lease
'###############################################################################################################
FIX_LeaseRef_DemandRecords:
   On Error GoTo MISSING_FIX_LeaseRef_DemandRecords

   Rst1.Open "SELECT D.* " & _
             "FROM DemandRecords AS D INNER JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
             "WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef));", Conn1, adOpenDynamic, adLockOptimistic

   If Rst1.EOF Then
      Rst1.Close
      GoTo UPDATE_UNITNUMBER_RECEIPT
   Else
      Conn1.Execute "UPDATE DemandRecords AS D, " & _
                        "[SELECT * FROM LeaseDetails WHERE Status = TRUE]. AS L, " & _
                        "[SELECT D.LeaseRef, D.DemandID " & _
                        " FROM DemandRecords AS D INNER JOIN " & _
                        "     DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
                    "WHERE D.DemandID = SQ.DemandID AND L.Status AND " & _
                        "L.SageAccountNumber = D.SageAccountNumber;"

      Conn1.Execute "UPDATE DemandRecords AS D, LeaseDetails AS L, " & _
                        "[SELECT D.LeaseRef, D.DemandID " & _
                        " FROM DemandRecords AS D INNER JOIN " & _
                              "DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
                        " WHERE (D.LeaseRef = '' OR ISNULL(D.LeaseRef)) AND " & _
                        "     S.A_M = 'M' AND S.SplitID = 1 " & _
                        "]. AS SQ SET D.LeaseRef = L.LeaseID " & _
                    "WHERE D.DemandID = SQ.DemandID AND " & _
                        "L.SageAccountNumber = D.SageAccountNumber;"
   End If

   Rst1.Close
   GoTo UPDATE_UNITNUMBER_RECEIPT

MISSING_FIX_LeaseRef_DemandRecords:
   Debug.Print Err.description
   MsgBox "This company database needs to be updated. Please contact with PCM Consulting.", vbCritical + vbOKOnly, "Lease Reference in Demand"

'###############################################################################################################
'  Update the UnitNumber in the transactions for expired lessee
'  Previously we used NOC as unit number if user creates SI after expired the lease
'  Now, system will update NOC with a lastest unit number, which will be found in the lease details
'     table where latest expired lease
UPDATE_UNITNUMBER_RECEIPT:
   On Error GoTo ERROR_UPDATE_UNITNUMBER_RECEIPT

   Rst1.Open "SELECT * " & _
             "FROM tlbReceipt AS R " & _
             "WHERE R.UnitID = 'NOC' OR ISNULL(R.UnitID);", Conn1, adOpenDynamic, adLockOptimistic

   While Not Rst1.EOF
      Rst2.Open "SELECT L.UnitNumber, L.Status " & _
                "FROM  LeaseDetails AS L " & _
                "WHERE L.SageAccountNumber = '" & Rst1.Fields.Item("SageAccountNumber").Value & "' " & _
                "ORDER BY L.Status, L.EndDate DESC;", Conn1, adOpenStatic, adLockReadOnly
'Debug.Print "SELECT L.UnitNumber, L.Status " & _
                "FROM  LeaseDetails AS L " & _
                "WHERE L.SageAccountNumber = '" & Rst1.Fields.Item("SageAccountNumber").Value & "' " & _
                "ORDER BY L.Status, L.EndDate DESC;"
      If Not Rst2.EOF Then
         Rst1.Fields.Item("UnitID").Value = Rst2.Fields.Item(0).Value
         Rst1.Update
      End If

      Rst1.MoveNext
      Rst2.Close
   Wend
   Rst1.Close

   GoTo UPDATE_AMOUNT_2_DECIMAL

ERROR_UPDATE_UNITNUMBER_RECEIPT:
   MsgBox "Unit information could not be updated automatically", vbCritical + vbOKOnly, "Unit Id for expired lessee in the receipt."
   UpdateDatabase1 = 1
   Exit Function

'###############################################################################################################
'  THE FOLLOWING UPDATE SQL WILL FIX AMOUNT FIGURE INTO 2 DECIMAL PLACE.
'  THIS UPDATE PROCESS WILL RUN IN MY PC. I WILL UPDATE ALL DATABASE IN THE PCM
'  WHEN I WILL CONFIRM ALL TRANSACTION SAVING PROCESS TO 2 DECIMAL PROCESS,
'  THIS UPDATE ROCESS WILL NOT BE REQUIRED ANY MORE.
UPDATE_AMOUNT_2_DECIMAL:
   On Error GoTo ERROR_UPDATE_AMOUNT_2_DECIMAL

   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
GoTo RESIZE_Email1_Tenants
      Update2DecimalPlace "DemandSplitRecords", "DSR", "Amount"
      Update2DecimalPlace "DemandSplitRecords", "DSR", "TotalAmount"
      Update2DecimalPlace "DemandSplitRecords", "DSR", "VATAmount"

      Update2DecimalPlace "tlbReceipt", "TransactionID", "Amount"
      Update2DecimalPlace "tlbReceipt", "TransactionID", "OSAmount"
      Update2DecimalPlace "tlbReceipt", "TransactionID", "ReceiptAmount"

      Update2DecimalPlace "tlbReceiptSplit", "TransactionID", "Amount"
      Update2DecimalPlace "tlbReceiptSplit", "TransactionID", "OSAmount"

      Update2DecimalPlace "tlbPayment", "TransactionID", "Amount"
      Update2DecimalPlace "tlbPayment", "TransactionID", "OSAmount"

      Update2DecimalPlace "tlbPaymentSplit", "TransactionID", "Amount"
      Update2DecimalPlace "tlbPaymentSplit", "TransactionID", "OSAmount"

      Update2DecimalPlace "PayTransactions", "TransactionID", "PaymentAmount"
      Update2DecimalPlace "RptTransactions", "TransactionID", "ReceiptAmount"

      Update2DecimalPlace "tblPurInv", "MY_ID", "TOTAL_AMOUNT"

      Update2DecimalPlace "tblPurInvSRec", "MY_ID", "NET_AMOUNT"
      Update2DecimalPlace "tblPurInvSRec", "MY_ID", "TOTAL_AMOUNT"

      Update2DecimalPlace "LServiceCharges", "ServiceCharge", "SCTotal"
      Update2DecimalPlace "LServiceCharges", "ServiceCharge", "SCAmount"
      Update2DecimalPlace "LRentCharges", "RentCharges", "BRTotal"
      Update2DecimalPlace "LRentCharges", "RentCharges", "BRAmount"
      Update2DecimalPlace "LInsuranceCharges", "InsCharges", "InsuranceEachPeriod"
      Update2DecimalPlace "LInsuranceCharges", "InsCharges", "TotalYearlyInsurance"
'
'AFTER RUNNING THIS SCRIPT, CHECK BY: SELECT R.TransactionID, R.Amount, R.DemandRef, S.DT FROM tlbReceipt AS R, (SELECT D.DemandID,  SUM(S.TotalAmount) AS DT FROM DemandRecords AS D LEFT JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID GROUP BY D.DemandID) AS S WHERE R.Type = 1 AND R.DemandRef = S.DemandID AND ROUND(R.Amount, 2) <> ROUND(CCUR(IIF(ISNULL(S.DT),'0',S.DT)), 2);
'IF THIS SQL GETS ANY RECORDS, THEN FIX THE DEMANDSPLIT TABLE
'
   End If

   GoTo RESIZE_Email1_Tenants

ERROR_UPDATE_AMOUNT_2_DECIMAL:
   MsgBox "System could not update all transactions to two decimal places", vbCritical + vbOKOnly, "Update Transactions"
   UpdateDatabase1 = -1
   Exit Function

'   Extend the field size Email on 12/06/2012 Tenants
'###############################################################################################################
RESIZE_Email1_Tenants:

   On Error GoTo MissingTable_RESIZE_Email1_Tenants

   Rst1.Open "SELECT Email1 FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Email1").DefinedSize = 40 Or Rst1.Fields.Item("Email1").DefinedSize = 99 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN Email1 TEXT(100)"
      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN Email2 TEXT(100)"
      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientOfficeEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientPersonalEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentOfficeEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentPersonalEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierPersonalEmail TEXT(100)"
      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SageSuppAC TEXT(100)"
   End If
   Rst1.Close
   Set Rst1 = Nothing

'   Exit Function
   GoTo BUGFIX_DemandSplitRecords_SplitID

MissingTable_RESIZE_Email1_Tenants:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - Email1 of DSR"
   UpdateDatabase1 = 1
   Exit Function

'BUGFIX: DEMAND SPLIT ID
'        When users delete the first line of demand record split which SplitID=1
'        System does not reschedule the split id
'        This creates the problem in the receipt unallocation
'        To solve, ..
BUGFIX_DemandSplitRecords_SplitID:
   Rst1.Open "SELECT Q1.* " & _
             "FROM ( " & _
               "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
               "FROM DemandSplitRecords AS DS " & _
               "GROUP BY DS.DemandID " & _
             ") AS Q1 INNER JOIN " & _
             "( " & _
               "SELECT DS.DemandID, COUNT(DSR) AS C " & _
               "FROM DemandSplitRecords AS DS " & _
               "GROUP BY DS.DemandID " & _
             ") AS Q2 ON Q2.DemandID = Q1.DemandID " & _
             "WHERE Q1.M <> Q2.C", Conn1, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then                                                 'There are demand split without SplitID
      Rst1.Close
      Conn1.Execute "UPDATE DemandSplitRecords " & _
                    "Set SplitID = SplitID - 1 " & _
                    "WHERE DSR IN ( " & _
                    "SELECT DS.DSR " & _
                    "FROM ((" & _
                       "SELECT DS.DemandID, Max(DS.SplitID) AS M " & _
                       "FROM DemandSplitRecords AS DS " & _
                       "GROUP BY DS.DemandID " & _
                    ") AS Q1 INNER JOIN " & _
                    "( " & _
                       "SELECT DS.DemandID, COUNT(DSR) AS C " & _
                       "FROM DemandSplitRecords AS DS " & _
                       "GROUP BY DS.DemandID " & _
                    ") AS Q2 ON Q2.DemandID = Q1.DemandID) INNER JOIN " & _
                       "DemandSplitRecords AS DS ON Q1.DemandID = DS.DemandID " & _
                    "WHERE Q1.M <> Q2.C);"

         Conn1.Execute "UPDATE tlbReceiptSplit AS R, DemandSplitRecords AS S " & _
                       "SET R.SplitID = S.SplitID " & _
                       "WHERE R.AllocTranID = S.DSR AND R.Amount = S.TotalAmount;"

         UpdateDatabase1 = 1

         Exit Function
   End If
   Rst1.Close

'   Add new column SelFund on 01/08/2012 Fund
'###############################################################################################################
ADD_SelFund_Fund:
   On Error GoTo CHANGE_ADD_SelFund_Fund

   Rst1.Open "SELECT SelFund FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_szFundID_Fund

CHANGE_ADD_SelFund_Fund:
   Conn1.Execute "ALTER TABLE Fund ADD COLUMN SelFund TEXT(10);"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN SelBanks TEXT(10);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column szFundID on 10/08/2012 Fund
'###############################################################################################################
ADD_szFundID_Fund:

   On Error GoTo CHANGE_ADD_szFundID_Fund

   Rst1.Open "SELECT szFundID FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo CREAT_SPLIT_tblPurInv

CHANGE_ADD_szFundID_Fund:
   Conn1.Execute "ALTER TABLE Fund ADD COLUMN szFundID TEXT(10);"
   Conn1.Execute "UPDATE Fund SET szFundID = CSTR(FundID);"
   Conn1.Execute "ALTER TABLE Fund ADD COLUMN FundList TEXT(255);"
   UpdateDatabase1 = 1
   Exit Function

'   Check the system PI Transactions which have not been exported to payment table.
'   If there is any transaction found, then system will export them
'###############################################################################################################
CREAT_SPLIT_tblPurInv:
' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
'   szSQL = "SELECT * " & _
'           "FROM   tblPurInv " & _
'           "WHERE MY_ID NOT IN (SELECT ParentID FROM tblPurInvSRec GROUP BY ParentID);"
    szSQL = "SELECT A.* FROM   tblPurInv A Left JOIN tblPurInvSRec B ON A.MY_ID=B.ParentID WHERE A.MY_ID<>B.ParentID;"
'Debug.Print szSql
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      szSQL = "SELECT * FROM tblPurInvSRec"
      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic

      While Not Rst1.EOF
         With Rst2
            .AddNew
            .Fields.Item("MY_ID").Value = UniqueID()
            .Fields.Item("ParentID").Value = Rst1.Fields.Item("MY_ID").Value
            .Fields.Item("TRAN_ID").Value = 1
            .Fields.Item("TRANS").Value = Rst1.Fields.Item("PropertyID").Value
            .Fields.Item("NOMINAL_CODE").Value = "0000"
            .Fields.Item("DEPT_ID").Value = 0
            .Fields.Item("description").Value = "DELETED PURCHASE TRANSACTIONS"
            .Fields.Item("NET_AMOUNT").Value = 0
            .Fields.Item("TAX_CODE").Value = "T9"
            .Fields.Item("VAT").Value = 0
            .Fields.Item("TOTAL_AMOUNT").Value = 0
            .Fields.Item("RecoverablePt").Value = 0

            .Update
         End With

         Rst1.MoveNext
      Wend
      Rst2.Close
   End If
   Rst1.Close
' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
'   szSQL = "SELECT * " & _
'           "From tlbPayment " & _
'           "WHERE TransactionID NOT IN (" & _
'               "SELECT PayHeader FROM tlbPaymentSplit GROUP BY PayHeader);"
 szSQL = "SELECT A.* From tlbPayment A LEFT JOIN tlbPaymentSplit B ON A.TransactionID= B.PayHeader WHERE  B.PayHeader is NULL;"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'Debug.Print szSql

   If Not Rst1.EOF Then
      szSQL = "SELECT * FROM tlbPaymentSplit;"
      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic

      While Not Rst1.EOF
         With Rst2
            .AddNew
            .Fields.Item("TransactionID").Value = UniqueID()
            .Fields.Item("PayHeader").Value = Rst1.Fields.Item("TransactionID").Value
            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
            .Fields.Item("OSAmount").Value = Rst1.Fields.Item("OSAmount").Value
'Debug.Print Rst1.Fields.Item("Amount").Value
            .Fields.Item("SplitID").Value = 1
            .Fields.Item("DueDate").Value = Rst1.Fields.Item("DDate").Value
            .Fields.Item("Description").Value = "DELETED PURCHASE TRANSACTIONS"
            .Update
         End With
         Rst1.MoveNext
      Wend
      Rst2.Close
   End If
   Rst1.Close
'********************************************************************************************
'   There are some PI found in the header table, which amounts don't match with split total
'********************************************************************************************
   szSQL = "SELECT P.*, S.ST " & _
           "FROM tblPurInv AS P, ( " & _
               "SELECT ParentID, SUM(TOTAL_AMOUNT) AS ST " & _
               "FROM tblPurInvSRec " & _
               "GROUP BY ParentID) AS S " & _
           "WHERE P.MY_ID = S.ParentID And P.TOTAL_AMOUNT <> S.ST"
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      While Not Rst1.EOF
         If Val(Rst1.Fields.Item("TOTAL_AMOUNT").Value) = 0 Then
            Conn1.Execute "UPDATE tblPurInvSRec AS S " & _
                          "SET    NET_AMOUNT = 0, VAT = 0, TOTAL_AMOUNT = 0 " & _
                          "WHERE  S.ParentID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
         Else
            Conn1.Execute "UPDATE tblPurInv AS P " & _
                          "SET    P.TOTAL_AMOUNT = " & Rst1.Fields.Item("ST").Value & " " & _
                          "WHERE  P.MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
         End If

         Rst1.MoveNext
      Wend
   End If
   Rst1.Close

'   New table on 03/09/2012 tlbBankReconcilation
'###############################################################################################################
NEW_TABLE_TlbBankReconcilation:

   On Error GoTo MissingTable_tlbBankReconcilation

   Rst1.Open "SELECT * FROM tlbBankReconcilation;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Rst1.Open "SELECT * FROM tlbBankReconcilation WHERE MY_ID = 'SAMRAT';", Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Conn1.Execute "DELETE * FROM tlbBankReconcilation;"
      CreateBankReconSplits
   End If
   Rst1.Close

   GoTo NEW_TABLE_tlbBankReconClosingBal

MissingTable_tlbBankReconcilation:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbBankReconcilation"
   UpdateDatabase1 = -1
   Exit Function

'   New table on 03/09/2012 tlbBankReconClosingBal
'###############################################################################################################
NEW_TABLE_tlbBankReconClosingBal:

   On Error GoTo MissingTable_tlbBankReconClosingBal

   Rst1.Open "SELECT * FROM tlbBankReconClosingBal;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADD_szTransactionID_tlbReceipt

MissingTable_tlbBankReconClosingBal:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbBankReconClosingBal"
   UpdateDatabase1 = -1
   Exit Function

'   Add new column szTransactionID on 11/09/2012 tlbReceipt
'###############################################################################################################
ADD_szTransactionID_tlbReceipt:
   On Error GoTo CHANGE_ADD_szTransactionID_tlbReceipt

   Rst1.Open "SELECT szTransactionID FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_szTransactionID_tlbPayment

CHANGE_ADD_szTransactionID_tlbReceipt:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN szTransactionID TEXT(20);"
   Conn1.Execute "UPDATE tlbReceipt SET szTransactionID = CSTR(TransactionID)"

   UpdateDatabase1 = 1
   Exit Function

'   Add new column szTransactionID on 11/09/2012 tlbPayment
'###############################################################################################################
ADD_szTransactionID_tlbPayment:
   On Error GoTo CHANGE_ADD_szTransactionID_tlbPayment

   Rst1.Open "SELECT szTransactionID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   'GoTo ADD_RRR_Fund
   GoTo FIXING_SplitID_tlbPaymentSplit

CHANGE_ADD_szTransactionID_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN szTransactionID TEXT(20);"
   Conn1.Execute "UPDATE tlbPayment SET szTransactionID = CSTR(TransactionID)"
   UpdateDatabase1 = 1
   Exit Function

'   Add a column RRR on 30/10/2012 Fund
'###############################################################################################################
'ADD_RRR_Fund:
'   On Error GoTo CHANGE_ADD_RRR_Fund
'
'   Rst1.Open "SELECT RRR FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
'   Rst1.Close
'
'   GoTo FIXING_SplitID_tlbPaymentSplit
'
'CHANGE_ADD_RRR_Fund:
'   Conn1.Execute "ALTER TABLE Fund ADD COLUMN RRR TEXT(10);"
'   UpdateDatabase1 = 1
'   Exit Function

'********************************************************************************************
'   There are some PI found which FundID don't match with split FundID
'********************************************************************************************
FIXING_SplitID_tlbPaymentSplit:
   szSQL = "SELECT A.PayHeader, A.X " & _
           "FROM ( " & _
                 "SELECT PayHeader, SplitID, COUNT(SplitID) AS X " & _
                 "From tlbPaymentSplit " & _
                 "GROUP BY PayHeader, SplitID " & _
           ") AS A " & _
           "WHERE A.X > 1;"
'
'            Debug.Print szSQL
'             Debug.Print time
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'   Debug.Print time
   If Not Rst1.EOF Then
      'keep a log in the error log table
      Conn1.Execute "Insert into SpareTable5(ClientID,Code,CC) values('Login','" & Date & "' ,'tlbPaymentSplit SplitID duplicated " & Rst1.Fields("PayHeader").Value & "' )"
      While Not Rst1.EOF
         Rst2.Open "SELECT * FROM tlbPaymentSplit AS S " & _
                   "WHERE S.PayHeader = " & Rst1.Fields.Item(0).Value & ";", Conn1, adOpenDynamic, adLockOptimistic

         For i = 1 To RecordCount(Rst2) 'Rst1.Fields.Item("X").Value 'fixed by anol 20181119
            Rst2.Fields.Item("SplitID").Value = i
            Rst2.Update
            Rst2.MoveNext
         Next i

         Rst1.MoveNext
         Rst2.Close
      Wend
   End If
   Rst1.Close
'   Debug.Print time

'********************************************************************************************
'   FIXING DATA: There are some PI found which FundID don't match with split FundID
'********************************************************************************************
''FIXING_FUNDID_tlbPayment_tlbPaymentSplit:
''   szSQL = "SELECT P.* " & _
''           "FROM tlbPaymentSplit AS S, tlbPayment AS P " & _
''           "WHERE P.TransactionID = S.PayHeader AND " & _
''                 "P.FundID <> S.FundID AND S.Splitid = 1;"
''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''
''   If Not Rst1.EOF Then
''   'we are not doing this anymore by anol 20181119
'''      Conn1.Execute "UPDATE tlbPayment AS P, tlbPaymentSplit AS S " & _
'''                    "SET    P.FundID = S.FundID " & _
'''                    "WHERE  P.TransactionID = S.PayHeader AND " & _
'''                           "S.Splitid = 1;"
''   End If
''   Rst1.Close

'********************************************************************************************
'   FIXING DATA: tlbReceiptSplit table has some transaction with FundID = 0
'********************************************************************************************
FIXING_FUNDID_tlbReceiptSplit:
   szSQL = "SELECT S.* " & _
           "FROM tlbReceiptSplit AS S " & _
           "WHERE S.FundID = 0;"
   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   If Not Rst1.EOF Then
      Conn1.Execute "UPDATE tlbReceiptSplit AS S1, tlbReceiptSplit AS S2, " & _
                        "tlbReceipt AS R1,tlbReceipt AS R2, RptTransactions AS T " & _
                    "Set S1.FundID = S2.FundID, S1.SplitID = S2.SplitID " & _
                    "WHERE S1.FundID = 0 AND R1.TransactionID = T.FromTran AND " & _
                        "T.ToTran = R2.TransactionID AND R1.TransactionID = S1.RptHeader AND " & _
                        "R2.TransactionID = S2.RptHeader  AND VAL(S1.AllocTranID) = S2.RptHeader;"
   End If
   Rst1.Close

'********************************************************************************************
'   FIXING DATA: tlbPayment, tlbReceipt & tlbBankReconcilation tables have some transactions
'                which have amount transaction method is 1 (CHQ, DD, etc)
'********************************************************************************************
FIXING_AMT_MTH_tlbPayment:
   szSQL = "UPDATE tlbPayment " & _
           "SET PayAmtType = 'CHQ' " & _
           "WHERE PayAmtType = '1';"
   Conn1.Execute szSQL

FIXING_AMT_MTH_tlbReceipt:
   szSQL = "UPDATE tlbReceipt " & _
           "SET RptAmtType = 'CHQ' " & _
           "WHERE RptAmtType = '1';"
   Conn1.Execute szSQL

FIXING_AMT_MTH_tlbBankReconcilation:
   szSQL = "UPDATE tlbBankReconcilation " & _
           "SET TranMth = 'CHQ' " & _
           "WHERE TranMth = '1';"
   Conn1.Execute szSQL

'********************************************************************************************
'   FIXING DATA: There are some reconciled transactions found dated after their
'                reconciled statement date. which tran dt is later then the bank reconcilation statement date
'********************************************************************************************
'Below procedure is rem out by anol 27 07 2016
''FIXING_TRANS_DT_tlbPayment:
''   Dim sChoice As Single
''
''   szSQL = "SELECT * " & _
''           "FROM  tlbBankReconcilation " & _
''           "WHERE TDate > ReconDate "
''
''   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      sChoice = MsgBox("There are some reconciled transactions found dated after their reconciled statement date." + Chr(13) + _
''                    "Please contact PCM Support to correct these transactions before proceeding further." + Chr(13) + _
''                    "Click OK to print a list of these transactions.", vbCritical + vbOKCancel, _
''                    "Purchase Payment")
''      If sChoice = vbOK Then
''         Rst1.Close
''         ShowReport App.Path & szReportPath & "\TranDt_BankRecDate.rpt"
''         UpdateDatabase1 = 0
''         Exit Function
''      End If
''   End If
''   Rst1.Close
   'end of rem
'Update tlbBankReconcilation
'SET TDate = ReconDate, DDate = ReconDate
'Where TDate > ReconDate And (TransactionType = 11 OR TransactionType = 12 OR TransactionType = 8 OR TransactionType = 9 OR TransactionType = 3 OR TransactionType = 4)
'--------------------------------------------------------------------
'Update tlbBankPayment
'Set TRAN_DATE = CDate(Left(ReconNow, 10))
'Where TRAN_DATE > CDate(Left(ReconNow, 10)) And (TransactionType = 11 OR TransactionType = 12) and ReconNow <> "" and not isnull(ReconNow)
'--------------------------------------------------------------------
'Update tlbPayment
'Set PDATE = CDate(Left(ReconNow, 10))
'WHERE PDATE > CDATE(LEFT(ReconNow, 10))  and (Type = 8 OR Type = 9) and ReconNow <> "" and not isnull(ReconNow)
'--------------------------------------------------------------------
'Update tlbReceipt
'SET RDATE = CDATE(LEFT(ReconNow, 10)), DDATE = CDATE(LEFT(ReconNow, 10))
'WHERE RDATE > CDATE(LEFT(ReconNow, 10))  and (Type = 3 or Type = 4) and ReconNow <> "" and not isnull(ReconNow)




'********************************************************************************************
'   FIXING DATA: There are some SRR found in the receipt and split table, their split OS <> header OS
'********************************************************************************************
FIXING_SRR_OS_Split:
   szSQL = "UPDATE  tlbReceiptSplit AS S, tlbReceipt AS R " & _
           "SET     R.OSAmount = S.OSAmount " & _
           "WHERE   R.TransactionID = S.RptHeader AND " & _
                   "R.Type = 23 AND " & _
                   "R.OSAmount <> S.OSAmount;"
   Conn1.Execute szSQL

   szSQL = "UPDATE  tlbReceiptSplit AS S " & _
           "SET     S.SplitID = 1 " & _
           "WHERE   S.SplitID = -1;"
   Conn1.Execute szSQL

'********************************************************************************************
'   FIXING DATA: There are some SA found in the receipt table without split in the split table
'********************************************************************************************
' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
'   szSQL = "SELECT * " & _
'           "FROM  tlbReceipt " & _
'           "WHERE Type = 4 AND Amount <> 0 AND " & _
'               "TransactionID NOT IN (" & _
'                  "SELECT RptHeader " & _
'                  "FROM   tlbReceiptSplit " & _
'                  "GROUP BY RptHeader);"
   szSQL = "SELECT A.* FROM  tlbReceipt A Left join tlbReceiptSplit B ON A.TransactionID =B.RptHeader WHERE B.RptHeader IS NULL AND Type = 4 AND A.Amount <> 0 ;"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      szSQL = "SELECT * FROM tlbReceiptSplit;"
      Rst2.Open szSQL, Conn1, adOpenDynamic, adLockPessimistic

      While Not Rst1.EOF
         With Rst2
            .AddNew
            .Fields.Item("TransactionID").Value = UniqueID()
            .Fields.Item("RptHeader").Value = Rst1.Fields.Item("TransactionID").Value
            .Fields.Item("FundID").Value = Rst1.Fields.Item("FundID").Value
            .Fields.Item("Amount").Value = Rst1.Fields.Item("Amount").Value
            .Fields.Item("OSAmount").Value = .Fields.Item("OSAmount").Value
            .Fields.Item("SplitID").Value = 1
            .Fields.Item("DueDate").Value = Format(Rst1.Fields.Item("DDate").Value, "dd mmmm yyyy")
            .Fields.Item("Description").Value = "Receipt on Account"
            .Update
         End With

         Rst1.MoveNext
      Wend
      Rst2.Close
   End If

   Rst1.Close

'********************************************************************************************
'   FIXING DATA:
'********************************************************************************************
FIXING_HEADER_n_SPLIT_SI_n_PI:
   Call Pi_Check_pre(Conn1) 'written by anol issue 791 Batch Payments (Support - WPM) 2019-07-25

   Call SiPi_Check(Conn1, "PI", "25875")
   Call SiPi_Check(Conn1, "SI", "15875")

'How to Fix:
'     Look at the total of tlbPayment header and split amount total
'     With the help of PayTransaction table, determine the split's amount

'********************************************************************************************
'   There are some PP found in the header table, which O/S amounts don't match with split O/S total
'********************************************************************************************
'   szSQL = "SELECT P.TransactionID, S.ST " & _
'           "FROM tlbPayment AS P, ( " & _
'               "SELECT PayHeader, SUM(OSAmount) AS ST " & _
'               "FROM tlbPaymentSplit " & _
'               "GROUP BY PayHeader) AS S " & _
'           "WHERE P.TransactionID = S.PayHeader And P.OSAmount <> S.ST"
'   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
''      szSQL = SQL2String(Rst1, 0)
''Debug.Print szSQL
''      MsgBox "The following ID(s) are duplicating: " & szSQL & ". Please contact with PCM Consulting.", vbCritical & vbOKOnly, "Data need to fix"
'      While Not Rst1.EOF
'         szSQL = "UPDATE tlbPayment " & _
'                 "SET OSAmount = " & Rst1.Fields.Item(1).Value & " " & _
'                 "WHERE TransactionID = " & Rst1.Fields.Item(0).Value & ";"
''Debug.Print szSQL
'         Conn1.Execute szSQL
'         Rst1.MoveNext
'      Wend

'      While Not Rst1.EOF
'         If Val(Rst1.Fields.Item("TOTAL_AMOUNT").Value) = 0 Then
'            Conn1.Execute "UPDATE tblPurInvSRec AS S " & _
'                          "SET    NET_AMOUNT = 0, VAT = 0, TOTAL_AMOUNT = 0 " & _
'                          "WHERE  S.ParentID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         Else
'            Conn1.Execute "UPDATE tblPurInv AS P " & _
'                          "SET    P.TOTAL_AMOUNT = " & Rst1.Fields.Item("ST").Value & " " & _
'                          "WHERE  P.MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
'         End If
'
'         Rst1.MoveNext
'      Wend
'   End If
'   Rst1.Close

'   Add a columns RRTotal on 08/08/2011 LRentCharges
'###############################################################################################################
ADD_RRTotal_LRentCharges:
   On Error GoTo CHANGE_ADD_RRTotal_LRentCharges

   Rst1.Open "SELECT RRTotal FROM LRentCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo LeaseHistory

CHANGE_ADD_RRTotal_LRentCharges:
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRTotal DOUBLE;"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRAmount DOUBLE;"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RRPrint TEXT(1);"
   UpdateDatabase1 = 1
   Exit Function

'   New table on 10/12/2012 LeaseHistory
'###############################################################################################################
LeaseHistory:
   On Error GoTo MissingTable_LeaseHistory

   Rst1.Open "SELECT * FROM LeaseHistory;", Conn1, adOpenStatic, adLockReadOnly
   lSlNumber = Rst1.RecordCount
   Rst1.Close

   If lSlNumber = 0 Then
      Rst1.Open "SELECT * FROM LeaseHistory;", Conn1, adOpenDynamic, adLockPessimistic
      Rst2.Open "SELECT * FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

      While Not Rst2.EOF
         Rst1.AddNew
         Rst1.Fields.Item("HistoryID").Value = UniqueID()
         For i = 1 To Rst1.Fields.Count - 1
            For lSlNumber = 0 To Rst2.Fields.Count - 1
               If Rst1.Fields.Item(i).Name = Rst2.Fields.Item(lSlNumber).Name Then
                  Rst1.Fields.Item(i).Value = Rst2.Fields.Item(lSlNumber).Value
                  Exit For
               End If
            Next lSlNumber
         Next i
         Rst1.Update
         Rst2.MoveNext
      Wend

      Rst1.Close
      Rst2.Close
   End If

   GoTo MODIFY_Code_FREQ_SecondaryCode

MissingTable_LeaseHistory:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - LeaseHistory"
   UpdateDatabase1 = -1
   Exit Function


'   Amend DATA Code FREQ on 30/01/13 SecondaryCode
'###############################################################################################################
MODIFY_Code_FREQ_SecondaryCode:
   On Error GoTo CHANGE_Code_FREQ_SecondaryCode

   Rst1.Open "SELECT Code FROM SecondaryCode WHERE Code = 'DAILY' AND PrimaryCode = 'FREQ';", Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Code = 'QTR', SC.Value = 'QUARTERLY', SC.Description = '4' WHERE SC.Code = 'DAILY' AND SC.PrimaryCode = 'FREQ';"
      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Code = 'HY', SC.Value = 'HALF YEARLY', SC.Description = '2' WHERE SC.Code = 'MTHADV' AND SC.PrimaryCode = 'FREQ';"
      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Description = '12' WHERE SC.Code = 'MONTHLY' AND SC.PrimaryCode = 'FREQ';"
      Conn1.Execute "UPDATE SecondaryCode AS SC SET SC.Description = '52' WHERE SC.Code = 'WEEKLY' AND SC.PrimaryCode = 'FREQ';"
   End If
   Rst1.Close

   GoTo NEW_TABLE_FinancialYear

CHANGE_Code_FREQ_SecondaryCode:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase1 = 1
   Exit Function

'   New table on 30/01/13 FinancialYear
'###############################################################################################################
NEW_TABLE_FinancialYear:
   On Error GoTo MissingTable_FinancialYear

   szEmail = "FinancialYear"
   Rst1.Open "SELECT * FROM " & szEmail & ";", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   szEmail = "Periods"
   Rst1.Open "SELECT * FROM " & szEmail & ";", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_LChildsRef_DemandSplPreview

MissingTable_FinancialYear:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - " & szEmail
   UpdateDatabase1 = -1
   Exit Function

FixDataByUser:
'System will jump here without modifying further. System wants user to do some job.
'Check the caller point for further information.
   UpdateDatabase1 = 0

'   Add a columns LChildsRef on 08/03/2013 DemandSplPreview
'###############################################################################################################
ADD_LChildsRef_DemandSplPreview:
   On Error GoTo CHANGE_ADD_LChildsRef_DemandSplPreview

   Rst1.Open "SELECT LChildsRef FROM DemandSplPreview;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_FinancialYear_GlobalSC

CHANGE_ADD_LChildsRef_DemandSplPreview:
   Conn1.Execute "ALTER TABLE DemandSplPreview ADD COLUMN LChildsRef TEXT(25);"
   UpdateDatabase1 = 1
   Exit Function

'   Add a columns FinancialYear on 19/04/2013 GlobalSC
'###############################################################################################################
ADD_FinancialYear_GlobalSC:
   On Error GoTo CHANGE_ADD_FinancialYear_GlobalSC

   Rst1.Open "SELECT FinancialYear FROM GlobalSC;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_CBY_Property

CHANGE_ADD_FinancialYear_GlobalSC:
   Conn1.Execute "ALTER TABLE GlobalSC ADD COLUMN FinancialYear TEXT(20);"
   Conn1.Execute "ALTER TABLE GlobalInsurance ADD COLUMN FinancialYear TEXT(20);"
   Conn1.Execute "ALTER TABLE GlobalRC ADD COLUMN FinancialYear TEXT(20);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new column CBY on 22/04/2013 Property
'###############################################################################################################
ADDNEW_COL_CBY_Property:
   On Error GoTo MISSING_ADDNEW_COL_CBY_Property

   Rst1.Open "SELECT CBY FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_REC_DTKW

MISSING_ADDNEW_COL_CBY_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN CBY TEXT(20);"
   UpdateDatabase1 = 1
   Exit Function

'   Add new record EMAIL_DEMAND_TEMPLATE on 30/04/2013 PrimaryCode
'###############################################################################################################
ADDNEW_REC_DTKW:
   With Rst1
      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'DTKW';", Conn1, adOpenStatic, adLockReadOnly

      If .EOF Then
         .Close
         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         !Code = "DTKW"
         !Value = "DEMAND TEMPLATE KEYWORD"
         !Flexible = False
         .Update
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item(0).Value = "DTKW"
         .Fields.Item(1).Value = "CN"
         .Fields.Item(2).Value = "CLIENT NAME"
         .Fields.Item(3).Value = "<CLIENT NAME>"
         .Update
         .AddNew
         .Fields.Item(0).Value = "DTKW"
         .Fields.Item(1).Value = "LN"
         .Fields.Item(2).Value = "LESSEE NAME"
         .Fields.Item(3).Value = "<LESSEE NAME>"
         .Update
      End If
      .Close
   End With
   GoTo NEW_TABLE_NJ_Header

'   ----SKIPPED-----
'   Control account has been moved into the NominalLedger table.
'   New table on 07/05/2013 SpareTable1         ~~~~ Control Account ~~~~
'###############################################################################################################
NEW_TABLE_SpareTable1:
   On Error GoTo MissingTable_SpareTable1

   Rst1.Open "SELECT * FROM SpareTable1;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Rst1.Open "SELECT * FROM SpareTable1;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Count = 1 Or Rst1.Fields.Count = 7 Then
      If Rst1.Fields.Count = 1 Then
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN CAName    TEXT(50);"
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN NCode     TEXT(10);"
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN NName     TEXT(100);"
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN ClientID  TEXT(10);"
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Fixed     BIT;"
         Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN DisOrder  Single;"
      End If
      Rst1.Close
      Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Posting   BIT;"
      Conn1.Execute "ALTER TABLE SpareTable1 ADD COLUMN Type      TEXT(1);"

      Rst2.Open "SELECT * FROM Client;", Conn1, adOpenStatic, adLockReadOnly
      While Not Rst2.EOF
         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
                       "Values ('" & UniqueID() & "', 'Sales Ledger Control', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 1, FALSE, 'S');"
         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
                       "Values ('" & UniqueID() & "', 'Purchase Ledger Control', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 2, FALSE, 'P');"
         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
                       "Values ('" & UniqueID() & "', 'Input VAT', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 3, FALSE, 'I');"
         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
                       "Values ('" & UniqueID() & "', 'Output VAT', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 4, FALSE, 'O');"
         Conn1.Execute "INSERT INTO SpareTable1 (RecordID, CAName, ClientID, Fixed, DisOrder, Posting, Type) " & _
                       "Values ('" & UniqueID() & "', 'Retained Earnings', '" & Rst2.Fields.Item("ClientID").Value & "', TRUE, 5, FALSE, 'R');"
         Rst2.MoveNext
      Wend
      Rst2.Close
   Else
      Rst1.Close
   End If

   GoTo NEW_TABLE_NJ_Header

MissingTable_SpareTable1:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - SpareTable1"
   UpdateDatabase1 = -1
   Exit Function

'   New table on 14/05/2013 NJ_Header            ~~~~  Nominal Journal Header  ~~~~
'###############################################################################################################
NEW_TABLE_NJ_Header:
   On Error GoTo MissingTable_NJ_Header

   Rst1.Open "SELECT * FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo NEW_TABLE_NJ_Split

MissingTable_NJ_Header:
   Conn1.Execute _
      "CREATE TABLE NJ_Header " & _
         "(" & _
            "RecordID      LONG NOT NULL PRIMARY KEY, " & _
            "ClientID      TEXT(10) NOT NULL, " & _
            "PropertyID    TEXT(4), " & _
            "NJDate        TEXT(20), " & _
            "NJTitle       TEXT(100), " & _
            "History       BIT, " & _
            "Posted        BIT, " & _
            "PrintThis     BIT" & _
         ");"

   UpdateDatabase1 = 1
   Exit Function

'   New table on 14/05/2013 NJ_Split         ~~~~  Nominal Journal Splits  ~~~~
'###############################################################################################################
NEW_TABLE_NJ_Split:
   On Error GoTo MissingTable_NJ_Split

   Rst1.Open "SELECT * FROM NJ_Split;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo NEW_TABLE_ReportCategory

MissingTable_NJ_Split:
   Conn1.Execute _
      "CREATE TABLE NJ_Split " & _
         "(" & _
            "RecordID      TEXT(50) NOT NULL PRIMARY KEY, " & _
            "ParentID      LONG NOT NULL, " & _
            "NC            TEXT(10), " & _
            "SpLineDes     TEXT(200), " & _
            "FundID        LONG, " & _
            "TYPE_ID       BYTE, " & _
            "NetAmt        CURRENCY, " & _
            "VAT_CODE      TEXT(5), " & _
            "VATAmt        CURRENCY, " & _
            "TotalAmt      CURRENCY" & _
         ");"

   UpdateDatabase1 = 1
   Exit Function

'   New table on 17/05/2013 ReportCategory         ~~~~  Report Category  ~~~~
'###############################################################################################################
NEW_TABLE_ReportCategory:
   On Error GoTo MissingTable_ReportCategory

   Rst1.Open "SELECT * FROM ReportCategory;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_ClientID_tlbPayment

MissingTable_ReportCategory:
   Conn1.Execute _
      "CREATE TABLE ReportCategory " & _
         "(" & _
            "RecordID      TEXT(50) NOT NULL PRIMARY KEY, " & _
            "ClientID      TEXT(10) NOT NULL, " & _
            "CategoryCode  TEXT(8), " & _
            "CategoryName  TEXT(100), " & _
            "CatDesc       TEXT(255)" & _
         ");"

   UpdateDatabase1 = 1
   Exit Function

'   Add new column ClientID on 24/05/2013 tlbPayment
'###############################################################################################################
ADD_ClientID_tlbPayment:
   On Error GoTo ERROR_ADD_ClientID_tlbPayment

   Rst1.Open "SELECT ClientID FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo VIEWS_NJ
ERROR_ADD_ClientID_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN ClientID TEXT(10);"

'  Update ClientID if there is existing PP

   Conn1.Execute "UPDATE tlbPayment AS P, tlbClientBanks AS B " & _
                 "SET P.ClientID = B.CLIENT_ID " & _
                 "WHERE NOT ISNULL(P.NominalCode) AND P.NominalCode <> '' AND " & _
                     "P.NominalCode =  B.NominalCode;"

   UpdateDatabase1 = 1
   Exit Function

'  CREATE a View NJ_HeaderTotal ON 28/05/2013
'###############################################################################################################
VIEWS_NJ:
   On Error GoTo ERROR_VIEWS_NJ

   Rst1.Open "SELECT * FROM NJ_HeaderTotal;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_REC_CAT

ERROR_VIEWS_NJ:
   Conn1.Execute "CREATE VIEW NJ_HeaderTotal " & _
                  "AS SELECT D.RecordID, D.ClientID, D.PropertyID, D.NJDate, D.NJTitle, D.History, SUM(S.TotalAmt) AS TotalAmt " & _
                  "FROM NJ_Header AS D INNER JOIN NJ_Split AS S ON D.RecordID=S.ParentID " & _
                  "GROUP BY D.RecordID, D.ClientID, D.PropertyID, D.NJDate, D.NJTitle, D.History;"
'   Conn1.Execute "CREATE VIEW NJ_ID " & _
'                  "AS SELECT MAX(RecordID)+1 AS NJ " & _
'                  "FROM NJ_Header;"
   UpdateDatabase1 = 1
   Exit Function

'   Add new record CAT on 27/06/13 PrimarySecondaryCode     --> Control Account Type
'###############################################################################################################
ADDNEW_REC_CAT:
   On Error GoTo MissingRec_REC_CAT

   With Rst1

      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'CAT';", Conn1, adOpenStatic, adLockReadOnly

      If .EOF Then
         .Close
         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         !Code = "CAT"
         !Value = "Control Account Type"
         .Update
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item(0).Value = "CAT"
         .Fields.Item(1).Value = "S"
         .Fields.Item(2).Value = "Sales"
         .Fields.Item(3).Value = "Sales Control Account"
         .Update
         .AddNew
         .Fields.Item(0).Value = "CAT"
         .Fields.Item(1).Value = "P"
         .Fields.Item(2).Value = "Purchase"
         .Fields.Item(3).Value = "Purchase Control Account"
         .Update
         .AddNew
         .Fields.Item(0).Value = "CAT"
         .Fields.Item(1).Value = "I"
         .Fields.Item(2).Value = "Input VAT"
         .Fields.Item(3).Value = "Input VAT"
         .Update
         .AddNew
         .Fields.Item(0).Value = "CAT"
         .Fields.Item(1).Value = "O"
         .Fields.Item(2).Value = "Output VAT"
         .Fields.Item(3).Value = "Output VAT"
         .Update
         .AddNew
         .Fields.Item(0).Value = "CAT"
         .Fields.Item(1).Value = "R"
         .Fields.Item(2).Value = "Retained Earnings"
         .Fields.Item(3).Value = "Retained Earnings"
         .Update
         .Close
      End If
      .Close
   End With
   GoTo Modify_TABLE_PropertyAnalysis
   'Exit Function

MissingRec_REC_CAT:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase1 = 1
   Exit Function
Modify_TABLE_PropertyAnalysis:
   On Error GoTo Mod__PropertyAnalysis
   Rst1.Open "SELECT Reference FROM PropertyAnalysis;", Conn1, adOpenStatic, adLockReadOnly
      If Rst1.State = 1 Then
         Rst1.Close
      End If
      Exit Function
Mod__PropertyAnalysis:
      Conn1.Execute "ALTER TABLE PropertyAnalysis add COLUMN Reference text(255);"
      Conn1.Execute "ALTER TABLE PropertyAnalysis add COLUMN AnalysisValue1 Number;"
      Exit Function

   UpdateDatabase1 = 1
   Exit Function
End Function

Private Function UpdateDatabase2(Conn1 As ADODB.Connection)
   Dim szTemp()      As String
   Dim Count         As Long
   Dim szStr         As String
   Dim Rst2          As New ADODB.Recordset
   Dim iCount1       As Long
'DoEvents
''   Debug.Print " databaseupdate2 start " & time
''   UpdateDatabase2 = 0
''
'''   Resolved by BOSL. Issue 0000476. By Asif
'''   Add new column Posting on 18 Nov 2014 NominalLedger
'''###############################################################################################################
''Add_TransactionRef_NLPosting:
''   On Error GoTo ERROR_Add_TransactionRef_NLPosting
''
''   Rst1.Open "SELECT TRANSACTION_REF FROM NLPosting;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_Posting_NominalLedger
''
''ERROR_Add_TransactionRef_NLPosting:
''   Conn1.Execute "ALTER TABLE NLPosting ADD COLUMN TRANSACTION_REF TEXT(10);"
'''   Conn1.Execute "UPDATE NominalLedger SET Posting = TRUE;"
''   UpdateDatabase2 = 1
''   Exit Function
''
''
'''   Add new column Posting on 28/06/2013 NominalLedger
'''###############################################################################################################
''Add_Posting_NominalLedger:
''   On Error GoTo ERROR_Add_Posting_NominalLedger
''
''   Rst1.Open "SELECT Posting FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_NLPost_tlbReceipt
''
''ERROR_Add_Posting_NominalLedger:
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN Posting BIT;"
''   Conn1.Execute "UPDATE NominalLedger SET Posting = TRUE;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column NLPost on 11/07/2013 tlbReceipt
'''###############################################################################################################
''Add_NLPost_tlbReceipt:
''   On Error GoTo ERROR_Add_NLPost_tlbReceipt
''
''   Rst1.Open "SELECT NLPost FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_NLPost_tlbBankPayment
''
''ERROR_Add_NLPost_tlbReceipt:
''   Conn1.Execute "DELETE * FROM NLPosting;"
''   Conn1.Execute "ALTER TABLE NLPosting DROP COLUMN UNIQUE_REFERENCE_NO;"
''   Conn1.Execute "ALTER TABLE NLPosting ADD  COLUMN UNIQUE_REFERENCE_NO COUNTER;"
'''   Conn1.Execute "ALTER TABLE NLPosting ADD  COLUMN ClientID TEXT(10) NOT NULL;"
''
''   Conn1.Execute "ALTER TABLE tlbReceipt     ADD COLUMN NLPost BIT;"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column NLPost on XX/XX/201X tlbBankPayment
'''###############################################################################################################
''Add_NLPost_tlbBankPayment:
''   On Error GoTo ERROR_Add_NLPost_tlbBankPayment
''
''   Rst1.Open "SELECT NLPost FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_NLPost_tlbPayment
''
''ERROR_Add_NLPost_tlbBankPayment:
''   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN NLPost BIT;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column NLPost on XX/XX/201X tlbPayment
'''###############################################################################################################
''ADD_NLPost_tlbPayment:
''   On Error GoTo ERROR_Add_NLPost_tlbPayment
''
''   Rst1.Open "SELECT NLPost FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_NLPost_tblPurInv
''
''ERROR_Add_NLPost_tlbPayment:
''   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN NLPost BIT;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column NLPost on XX/XX/201X tblPurInv
'''###############################################################################################################
''Add_NLPost_tblPurInv:
''   On Error GoTo ERROR_Add_NLPost_tblPurInv
''
''   Rst1.Open "SELECT NLPost FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_ClientID_NominalLedger
''
''ERROR_Add_NLPost_tblPurInv:
''   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN NLPost BIT;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column ClientID on 19/07/2013 NominalLedger           'ORIGINALLY THIS ADD COLUMN WAS IN THE NL MODULE
'''###############################################################################################################
''Add_ClientID_NominalLedger:
''   On Error GoTo ERROR_Add_ClientID_NominalLedger
''
''   Rst1.Open "SELECT ClientID FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo NEW_TABLE_NJ_CC
''
''ERROR_Add_ClientID_NominalLedger:
'''Creating the composite key Code+ClientID
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN ClientID TEXT(10) NOT NULL;"
''   Conn1.Execute "UPDATE NominalLedger SET ClientID = 'NONE';"
''   Conn1.Execute "DROP INDEX PrimaryKey ON NominalLedger;"
''   Conn1.Execute "ALTER TABLE NominalLedger ADD CONSTRAINT pk_XXX PRIMARY KEY (Code, ClientID);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   New table on 31/07/2013 NJ_CC            ~~~~  Nominal Journal && Category Code  ~~~~
'''###############################################################################################################
''NEW_TABLE_NJ_CC:
''   On Error GoTo MissingTable_NJ_CC
''
''   Rst1.Open "SELECT * FROM NJ_CC;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo NEW_TABLE_NLSubTypes
''
''MissingTable_NJ_CC:
''   Conn1.Execute _
''      "CREATE TABLE NJ_CC " & _
''         "(" & _
''            "RecordID      AUTOINCREMENT, " & _
''            "ClientID      TEXT(10) NOT NULL, " & _
''            "Code          TEXT(15) NOT NULL, " & _
''            "CC            TEXT(50) NOT NULL " & _
''         ");"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   New table on 13/08/2013 NLSubTypes            ~~~~  Nominal Journal Sub Type  ~~~~
'''###############################################################################################################
''NEW_TABLE_NLSubTypes:
''
''   On Error GoTo MissingTable_NLSubTypes
''
''   Rst1.Open "SELECT * FROM NLSubTypes;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo SubType_4_NominalLedger
''
''MissingTable_NLSubTypes:
''   Conn1.Execute _
''      "CREATE TABLE NLSubTypes " & _
''         "(" & _
''            "STCode TEXT(5) NOT NULL PRIMARY KEY, " & _
''            "STName TEXT(50) NOT NULL, " & _
''            "STDescription TEXT(200)" & _
''         ");"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('FA', 'Fixed Assets', 'Fixed Assets');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CA', 'Current Assets', 'Current Assets');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CDS', 'Creditors - Short', 'Current Liabilities');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CDL', 'Creditors - Long', 'Long term liabilities');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CR', 'Capital and Reserve', 'Capital and Reserve');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CC1', 'Income', 'Rent and Service Charge');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('CC2', 'Other Income', 'Other Income');"
''   Conn1.Execute _
''      "INSERT INTO NLSubTypes (STCode, STName, STDescription) " & _
''      "VALUES ('EX1', 'Overheads', 'Expenditure');"
''
'''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - NLSubTypes"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column SubType on 30/08/2013 NominalLedger
'''###############################################################################################################
''SubType_4_NominalLedger:
''   On Error GoTo ERROR_SubType_4_NominalLedger
''
''   Rst1.Open "SELECT SubType FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_ODM_ShoppingCentre
''
''ERROR_SubType_4_NominalLedger:
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN SubType TEXT(5);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column ODM on 30/08/2013 ShoppingCentre;     ONCE DATABASE MODIFICATION
'''###############################################################################################################
''ADDNEW_ODM_ShoppingCentre:
''   On Error GoTo ERROR_ADDNEW_ODM_ShoppingCentre
''
''   Rst1.Open "SELECT ODM FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo DELETE_RS_tlbLetterReports_Template__
''
''ERROR_ADDNEW_ODM_ShoppingCentre:
''   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN ODM TEXT(250);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Delete relationship between tlbLetterReports and Template__ on 23/09/08
'''###############################################################################################################
''DELETE_RS_tlbLetterReports_Template__:
''''   On Error Resume Next
'''
'''   Dim ws As WorkSpace
'''   Dim db As DAO.Database               'Creates a location to open a database to
'''   Dim rs As DAO.Recordset
'''
'''
'''Dim strConnection As String
'''
'''Set db = CurrentDb()
'''
'''Set ws = DBEngine.Workspaces(0)
'''Let strConnection = "ODBC;DSN=" & Adsn & ";UID=;PWD=RDSWKDPP"
'''Set db = ws.OpenDatabase("", False, False, strConnection)
'''
'''   Rst1.Open "SELECT szRelationship FROM Msysrelationships WHERE szObject = tlbLetterReports and szReferencedObject = Template__", Conn1, adOpenStatic, adLockReadOnly
'''
'''   If Not Rst1.EOF Then
'''MsgBox Rst1.Fields.Item("szRelationship").Value
'''      Conn1.Execute "ALTER TABLE  tlbLetterReports DROP CONSTRAINT " & Rst1.Fields.Item("szRelationship").Value
'''      Conn1.Execute "ALTER TABLE  tlbLetterReports DROP CONSTRAINT Template__tlbLetterReports"
'''   End If
'''   Rst1.Close
'''   Set Rst1 = Nothing
''
''REMOVE_CONSTRAINT_tlbLetterReports_TemplateID:
''   On Error GoTo ERROR_CONSTRAINT_tlbLetterReports_TemplateID
''
''   Rst1.Open "SELECT ODM FROM ShoppingCentre;", Conn1, adOpenDynamic, adLockOptimistic
''   If IsNull(Rst1.Fields.Item(0).Value) Then
''      Conn1.Execute "DROP INDEX TemplateID ON tlbLetterReports;"
''      Rst1.Fields.Item(0).Value = "tlbLetterReports:TemplateID"
''      Rst1.Update
''      Rst1.Close
''   Else
''      If InStr(Rst1.Fields.Item(0).Value, "tlbLetterReports:TemplateID") = 0 Then
''         Rst1.Fields.Item(0).Value = Rst1.Fields.Item(0).Value & "#tlbLetterReports:TemplateID"
''         Rst1.Update
''         Rst1.Close
''         Conn1.Execute "DROP INDEX TemplateID ON tlbLetterReports;"
''      Else
''         Rst1.Close
''      End If
''   End If
''
''   GoTo ADDNEW_COL_PostingDate_DemandRecords
''
''ERROR_CONSTRAINT_tlbLetterReports_TemplateID:
''   Rst1.Fields.Item(0).Value = Rst1.Fields.Item(0).Value & "#tlbLetterReports:TemplateID"
''   Rst1.Update
''   Rst1.Close
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PostingDate on 01/10/2013 DemandRecords
'''###############################################################################################################
''ADDNEW_COL_PostingDate_DemandRecords:
''   On Error GoTo MISSING_ADDNEW_COL_PostingDate_DemandRecords
''
''   Rst1.Open "SELECT PostingDate FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM DemandRecords WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE DemandRecords SET PostingDate = FORMAT(IssueDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''   GoTo Add_ClientID_NLPosting
''
''MISSING_ADDNEW_COL_PostingDate_DemandRecords:
''   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE DemandRecords SET PostingDate = FORMAT(IssueDate, 'DD MMMM YYYY')"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column ClientID on 08/04/2014 NLPosting
'''###############################################################################################################
''Add_ClientID_NLPosting:
''   On Error GoTo ERROR_Add_ClientID_NLPosting
''
''   Rst1.Open "SELECT ClientID FROM NLPosting;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo Add_DeleteFlag_NLPosting
''
''ERROR_Add_ClientID_NLPosting:
'''Creating the composite key Code+DeleteFlag
''   Conn1.Execute "ALTER TABLE NLPosting ADD COLUMN ClientID TEXT(10);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column DeleteFlag on 03/10/2013 NLPosting
'''###############################################################################################################
''Add_DeleteFlag_NLPosting:
''   On Error GoTo ERROR_Add_DeleteFlag_NLPosting
''
''   Rst1.Open "SELECT DeleteFlag FROM NLPosting;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_PostingDate_tlbReceipt
''
''ERROR_Add_DeleteFlag_NLPosting:
'''Creating the composite key Code+DeleteFlag
''   Conn1.Execute "ALTER TABLE NLPosting ADD COLUMN DeleteFlag BIT;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PostingDate on 17/10/2013 tlbReceipt
'''###############################################################################################################
''ADDNEW_COL_PostingDate_tlbReceipt:
''   On Error GoTo MISSING_ADDNEW_COL_PostingDate_tlbReceipt
''
''   Rst1.Open "SELECT PostingDate FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM tlbReceipt WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tlbReceipt SET PostingDate = FORMAT(RDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   GoTo ADDNEW_COL_PostingDate_tblPurInv
''
''MISSING_ADDNEW_COL_PostingDate_tlbReceipt:
''   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tlbReceipt SET PostingDate = FORMAT(RDate, 'DD MMMM YYYY')"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PostingDate on 21/10/2013 tblPurInv
'''###############################################################################################################
''ADDNEW_COL_PostingDate_tblPurInv:
''   On Error GoTo MISSING_ADDNEW_COL_PostingDate_tblPurInv
''
''   Rst1.Open "SELECT PostingDate FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM tblPurInv WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tblPurInv SET PostingDate = FORMAT(TRAN_DATE, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
'''Below line has been modified by anol 25 Nov 2015
''   Rst1.Open "SELECT * FROM tlbPayment WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tlbPayment SET PostingDate = FORMAT(PDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM tblBtRptTran WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tblBtRptTran SET PostingDate = FORMAT(DueDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM tblBatchTransaction WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tblBatchTransaction SET PostingDate = FORMAT(DueDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   GoTo ADDNEW_COL_PostingDate_tlbBankPayment
''
''MISSING_ADDNEW_COL_PostingDate_tblPurInv:
''   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tblPurInv SET PostingDate = FORMAT(TRAN_DATE, 'DD MMMM YYYY')"
'''  tlbPayment
''   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tlbPayment SET PostingDate = FORMAT(PDate, 'DD MMMM YYYY')"
'''  tblBtRptTran
''   Conn1.Execute "ALTER TABLE tblBtRptTran ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tblBtRptTran SET PostingDate = FORMAT(RptDt, 'DD MMMM YYYY')"
'''  tblBatchTransaction
''   Conn1.Execute "ALTER TABLE tblBatchTransaction ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tblBatchTransaction SET PostingDate = FORMAT(DueDate, 'DD MMMM YYYY')"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PostingDate on 21/10/2013 tlbBankPayment
'''###############################################################################################################
''ADDNEW_COL_PostingDate_tlbBankPayment:
''   On Error GoTo MISSING_ADDNEW_COL_PostingDate_tlbBankPayment
''
''   Rst1.Open "SELECT PostingDate FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM tlbBankPayment WHERE ISNULL(PostingDate);", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE tlbBankPayment SET PostingDate = FORMAT(TRAN_DATE, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   GoTo ADDNEW_COL_PostingDate_NJ_Header
''
''MISSING_ADDNEW_COL_PostingDate_tlbBankPayment:
''   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE tlbBankPayment SET PostingDate = FORMAT(TRAN_DATE, 'DD MMMM YYYY')"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PostingDate on 21/10/2013 NJ_Header
'''###############################################################################################################
''ADDNEW_COL_PostingDate_NJ_Header:
''   On Error GoTo MISSING_ADDNEW_COL_PostingDate_NJ_Header
''
''   Rst1.Open "SELECT PostingDate FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   Rst1.Open "SELECT * FROM NJ_Header WHERE ISNULL(PostingDate) ;", Conn1, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      Conn1.Execute "UPDATE NJ_Header SET PostingDate = FORMAT(NJDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''
''   GoTo FIX_POSTED_DATE_NLPosting
''
''MISSING_ADDNEW_COL_PostingDate_NJ_Header:
''   Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN PostingDate TEXT(20);"
''   Conn1.Execute "UPDATE NJ_Header SET PostingDate = FORMAT(NJDate, 'DD MMMM YYYY')"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Fix POSTED_DATE in the NLPosting on 07/11/2013
'''   POSTED_DATE should be earlier or same as TRANSACTION_DATE.
'''###############################################################################################################
''FIX_POSTED_DATE_NLPosting:
''   On Error GoTo MISSING_FIX_POSTED_DATE_NLPosting
''
''   Rst1.Open "SELECT Field6 " & _
''             "FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields.Item(0).Value <> "DONE" Then
''      Conn1.Execute "UPDATE NLPosting SET POSTED_DATE = TRANSACTION_DATE;"
''      Conn1.Execute "UPDATE ShoppingCentre SET Field6 = 'DONE';"
''   End If
''
''   Rst1.Close
''
'''Exit Function
''
''   GoTo ADDNEW_COL_CAName_NominalLedger
''
''MISSING_FIX_POSTED_DATE_NLPosting:
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new columns CAName on 11/11/2013 NominalLedger
'''###############################################################################################################
''ADDNEW_COL_CAName_NominalLedger:
''   On Error GoTo MISSING_ADDNEW_COL_CAName_NominalLedger
''
''   Rst1.Open "SELECT CAName FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_AllowOverDraft_tlbClientBanks
''
''MISSING_ADDNEW_COL_CAName_NominalLedger:
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN CAName TEXT(50);"
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN CAFixed BIT;"
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN CADisOrder Single;"
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN CAPosting BIT;"
''   Conn1.Execute "ALTER TABLE NominalLedger ADD COLUMN CAType TEXT(1);"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add a columns AllowOverDraft on 15/11/2013 tlbClientBanks
'''###############################################################################################################
''ADD_AllowOverDraft_tlbClientBanks:
''   On Error GoTo ERROR_ADD_AllowOverDraft_tlbClientBanks
''
''   Rst1.Open "SELECT AllowOverDraft FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   GoTo ADD_FundID_PropertyMaintHistory
''
''ERROR_ADD_AllowOverDraft_tlbClientBanks:
''   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN AllowOverDraft BIT;"
''   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN OverDraftLimit CURRENCY;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add a columns FundID on 16/11/2013 PropertyMaintHistory
'''###############################################################################################################
''ADD_FundID_PropertyMaintHistory:
''   On Error GoTo ERROR_ADD_FundID_PropertyMaintHistory
''
''   Rst1.Open "SELECT FundID FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   GoTo ADD_SelInFund_Fund
''
''ERROR_ADD_FundID_PropertyMaintHistory:
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN FundID LONG;"
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN OverrideBudget BIT;"
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN FYrID TEXT(20);"
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN BudgetPassed BIT;"
''   Conn1.Execute "UPDATE PropertyMaintHistory SET BudgetPassed = TRUE;"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column SelInFund on 12/12/2013 Fund
'''###############################################################################################################
''ADD_SelInFund_Fund:
''   On Error GoTo CHANGE_ADD_SelInFund_Fund
''
''   Rst1.Open "SELECT SelInFund FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_DateFrom_LUtilityUsage
''
''CHANGE_ADD_SelInFund_Fund:
''   Conn1.Execute "ALTER TABLE Fund ADD COLUMN SelInFund TEXT(10);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column DateFrom on 19/12/2013 LUtilityUsage
'''###############################################################################################################
''ADD_DateFrom_LUtilityUsage:
''   On Error GoTo CHANGE_ADD_DateFrom_LUtilityUsage
''
''   Rst1.Open "SELECT DateFrom FROM LUtilityUsage;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_DateTo_LUtilityUsage
''
''CHANGE_ADD_DateFrom_LUtilityUsage:
''   Conn1.Execute "ALTER TABLE LUtilityUsage ADD COLUMN DateFrom TEXT(20);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column DateTo on 19/12/2013 LUtilityUsage
'''###############################################################################################################
''ADD_DateTo_LUtilityUsage:
''   On Error GoTo CHANGE_ADD_DateTo_LUtilityUsage
''
''   Rst1.Open "SELECT DateTo FROM LUtilityUsage;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_StatementTemplate_DemandTypes
''
''CHANGE_ADD_DateTo_LUtilityUsage:
''   Conn1.Execute "ALTER TABLE LUtilityUsage ADD COLUMN DateTo TEXT(20);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add a column StatementTemplate on 10/01/2014 DemandTypes
'''###############################################################################################################
''ADD_StatementTemplate_DemandTypes:
''   On Error GoTo CHANGE_ADD_StatementTemplate_DemandTypes
''
''   Rst1.Open "SELECT StatementTemplate FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADD_EditRecord_DemandRecords
''
''CHANGE_ADD_StatementTemplate_DemandTypes:
''   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN StatementTemplate TEXT(100);"
''   Conn1.Execute "UPDATE DemandTypes SET StatementTemplate = 'InvDemandStatement.rpt' WHERE ISNULL(StatementTemplate);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add a column EditRecord on 17/01/2014 DemandRecords
'''###############################################################################################################
''ADD_EditRecord_DemandRecords:
''   On Error GoTo CHANGE_ADD_EditRecord_DemandRecords
''
''   Rst1.Open "SELECT EditRecord FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_TRAN_TYPE_PO
''
''CHANGE_ADD_EditRecord_DemandRecords:
''   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN EditRecord BIT;"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new record on 09/12/2013 tlbTransactionTypes
'''###############################################################################################################
''ADDNEW_TRAN_TYPE_PO:
''   Rst1.Open "SELECT * FROM tlbTransactionTypes WHERE CONSTANT = 'sdoPO';", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.EOF Then
''      Rst1.Close
''      Rst1.Open "SELECT * FROM tlbTransactionTypes;", Conn1, adOpenDynamic, adLockOptimistic
''      With Rst1
''         .AddNew
''         !TYPE_ID = 25
''         !description = "Purchase Order"
''         !CONSTANT = "sdoPO"
''         !ACCOUNTS_LINE = "Prestige"
''         .Update
''         .AddNew
''         !TYPE_ID = 26
''         !description = "Purchase Quote"
''         !CONSTANT = "sdoPQ"
''         !ACCOUNTS_LINE = "Prestige"
''         .Update
''      End With
''   End If
''   Rst1.Close
''
'''   New table on 30/01/2014 NLType
'''###############################################################################################################
''NEW_TABLE_NLType:
''   On Error GoTo MissingTable_NLType
''
''   Rst1.Open "SELECT * FROM NLType;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_RunningAutoDemand_Property
''   GoTo NEW_TABLE_POrders
''
''MissingTable_NLType:
''   Conn1.Execute _
''      "CREATE TABLE NLType " & _
''         "(" & _
''            "NLTypeCode    LONG NOT NULL PRIMARY KEY, " & _
''            "TypeValue     TEXT(100) NOT NULL " & _
''         ");"
''   Conn1.Execute "INSERT INTO NLType (NLTypeCode, TypeValue) " & _
''                 "Values (0, 'Not Defined');"
''   Conn1.Execute "INSERT INTO NLType (NLTypeCode, TypeValue) " & _
''                 "Values (1, 'Balance Sheet');"
''   Conn1.Execute "INSERT INTO NLType (NLTypeCode, TypeValue) " & _
''                 "Values (2, 'Profit and Loss');"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   New table on 09/12/2014 POrders            ~~~~  Purchase Order/Quote  ~~~~
'''###############################################################################################################
''NEW_TABLE_POrders:
''   On Error GoTo MissingTable_POrders
''
''   Rst1.Open "SELECT * FROM POrders;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
'''   Rst1.Open "SELECT * FROM POrderSplit;", Conn1, adOpenStatic, adLockReadOnly
'''   Rst1.Close
''
''   Exit Function
''
''MissingTable_POrders:
''   Conn1.Execute _
''      "CREATE TABLE POrders " & _
''         "(" & _
''            "MY_ID         TEXT(25) PRIMARY KEY, " & _
''            "SlNumber      LONG NOT NULL, " & _
''            "SupplierID    TEXT(10) NOT NULL, " & _
''            "Q_DATE        DATETIME, " & _
''            "Exp_DATE      DATETIME, " & _
''            "OrderType     BYTE, " & _
''            "PI_ID         TEXT(25), " & _
''            "Status        BIT, " & _
''            "TOTAL_AMOUNT  Currency, " & _
''            "History       BIT, " & _
''            "ClientID      TEXT(10), " & _
''            "PropertyID    TEXT(4) " & _
''         ");"
''
''   Conn1.Execute _
''      "CREATE TABLE POrderSplit " & _
''         "(" & _
''            "MY_ID         TEXT(25) PRIMARY KEY, " & _
''            "SlNumber      LONG NOT NULL, " & _
''            "SupplierID    TEXT(10) NOT NULL, " & _
''            "Q_DATE        DATETIME, " & _
''            "Exp_DATE      DATETIME, " & _
''            "OrderType     BYTE, " & _
''            "PI_ID         TEXT(25), " & _
''            "Status        BIT, " & _
''            "History       BIT, " & _
''            "TOTAL_AMOUNT  Currency, " & _
''            "ClientID      TEXT(10), " & _
''            "PropertyID    TEXT(4) " & _
''         ");"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column RunningAutoDemand on 24/09/2010 Property
'''###############################################################################################################
''ADDNEW_COL_RunningAutoDemand_Property:
''   On Error GoTo MISSING_ADDNEW_COL_RunningAutoDemand_Property
''
''   Rst1.Open "SELECT RunningAutoDemand FROM Property;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   GoTo NEW_TABLE_Contacts
''
''MISSING_ADDNEW_COL_RunningAutoDemand_Property:
''   Conn1.Execute "ALTER TABLE Property ADD COLUMN RunningAutoDemand TEXT(50);"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   New table on 20/02/2014 Contacts
'''###############################################################################################################
''NEW_TABLE_Contacts:
''   On Error GoTo MissingTable_Contacts
''
''   Rst1.Open "SELECT * FROM Contacts;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_FundCode_Fund
''
''MissingTable_Contacts:
''   Conn1.Execute _
''      "CREATE TABLE Contacts " & _
''         "(" & _
''            "ContactID           TEXT(25) PRIMARY KEY, " & _
''            "WhosContact         TEXT(1 ) NOT NULL, " & _
''            "HeadContact         TEXT(30) NOT NULL, " & _
''            "ContactName         TEXT(50), " & _
''            "CompanyName         TEXT(100), " & _
''            "OfficeAddressLine1  TEXT(100), " & _
''            "OfficeAddressLine2  TEXT(100), " & _
''            "OfficeAddressLine3  TEXT(100), " & _
''            "OfficeAddressLine4  TEXT(100), " & _
''            "OfficePostCode      TEXT(10), " & _
''            "OfficeTel           TEXT(40), " & _
''            "OfficeMobile        TEXT(40), " & _
''            "Mobile              TEXT(40), " & _
''            "PersonalEmail       TEXT(100), " & _
''            "AddressLine1        TEXT(100), " & _
''            "AddressLine2        TEXT(100), " & _
''            "AddressLine3        TEXT(100), " & _
''            "AddressLine4        TEXT(100), " & _
''            "PostCode            TEXT(10), " & _
''            "HomeTel             TEXT(40), " & _
''            "OfficeEmail         TEXT(100) " & _
''         ");"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column FundCode on 01/04/2014 Fund
'''###############################################################################################################
''ADDNEW_COL_FundCode_Fund:
''   On Error GoTo MISSING_ADDNEW_COL_FundCode_Fund
''
''   Rst1.Open "SELECT FundCode FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
''
''   Rst1.Close
''
''   GoTo ADDNEW_COL_ScheduleCode_Schedule
''
''MISSING_ADDNEW_COL_FundCode_Fund:
''   Conn1.Execute "ALTER TABLE Fund ADD COLUMN FundCode TEXT(12);"
''   Rst1.Open "SELECT FundCode, FundName FROM Fund;", Conn1, adOpenDynamic, adLockOptimistic
''   If RecordCount(Rst1) > 0 Then
''      While Not Rst1.EOF
''         szStr = UCase(Left(Replace(Rst1.Fields.Item(1).Value, " ", ""), 12))
''         Rst2.Open "SELECT * FROM Fund WHERE FundCode = '" & szStr & "';", Conn1, adOpenStatic, adLockReadOnly
''
''         If Rst2.EOF Then
''            Rst1.Fields.Item(0).Value = szStr
''            Rst2.Close
''         Else
''            Rst2.Close
''            count = 1
''            Do
''               szStr = Left(szStr, Len(szStr) - Len(CStr(count))) & CStr(count)
''               Rst2.Open "SELECT * FROM Fund WHERE FundCode = '" & szStr & "';"
''               If Rst2.EOF Then
''                  Rst1.Fields.Item(0).Value = szStr
''                  Rst2.Close
''                  Exit Do
''               Else
''                  Rst2.Close
''                  count = count + 1
''               End If
''            Loop
''         End If
''         Rst1.Update
''         Rst1.MoveNext
''      Wend
''   End If
''   Rst1.Close
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column ScheduleCode on 11/04/2014 Schedule
'''###############################################################################################################
''ADDNEW_COL_ScheduleCode_Schedule:
''   On Error GoTo MISSING_ADDNEW_COL_ScheduleCode_Schedule
''
''   Rst1.Open "SELECT ScheduleCode FROM Schedule;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
'''Exit Function
''   GoTo RESIZE_UNITID_EVERYWHERE
''
''MISSING_ADDNEW_COL_ScheduleCode_Schedule:
''   Conn1.Execute "ALTER TABLE Schedule ADD COLUMN ScheduleCode TEXT(15);"
''   Conn1.Execute "UPDATE Schedule SET ScheduleCode = UCASE(LEFT(ScheduleName, 4));"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Extend the field size on 07/04/2014 Units
'''###############################################################################################################
''RESIZE_UNITID_EVERYWHERE:
''   On Error GoTo MissingTable_RESIZE_UNITID_EVERYWHERE
''
'''  FIRST DELETE RELATIONSHIP AMONG ALL TABLES; DELETING RELATIONSHIPS
''   Dim cn         As New ADODB.Connection
''   Dim cat        As New ADOX.Catalog
''   Dim tbl        As New ADOX.table
''
''Dim n As Integer
''
'''n = FreeFile()
'''Open "C:\test.txt" For Append As #n
'''Print #n, FullDatabasePath ' write to file
'''Close #n
'''Exit Function
''
''   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FullDatabasePath & ";Jet OLEDB:Database Password=" & accessDBPws & ";"
'''Print #n, "1"
''   Set cat.ActiveConnection = cn
'''Print #n, "2   "
''   Set Rst1 = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
'''Print #n, "3   "
''   Do While Not Rst1.EOF
'''Print #n, szStr
''      szStr = Rst1!TABLE_NAME
''
''      Set tbl = cat.Tables(szStr)
''
''      For count = tbl.Keys.count - 1 To 0 Step -1
''         If tbl.Keys(count).Type = adKeyForeign Then tbl.Keys.Delete tbl.Keys(count).Name
''      Next
''      Rst1.MoveNext
''   Loop
'''Print #n, "#################### clear all the relationship #########################"
''   Rst1.Close
''   cn.Close
''   Set Rst1 = Nothing
''   Set cn = Nothing
''
''   Rst1.Open "SELECT UnitNumber FROM Units;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields.Item("UnitNumber").DefinedSize = 8 Then
''      Rst1.Close
''      Conn1.Execute "ALTER TABLE Units ALTER COLUMN UnitNumber TEXT(12)"
''      Conn1.Execute "ALTER TABLE LeaseDetails ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE UnitMaintHistory ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE UnitAnalysis ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE UnitInsurance ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbRechargePre ALTER COLUMN UNIT_ID TEXT(25)"
''      Conn1.Execute "ALTER TABLE UnitUtilities ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE UnitSafety ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN UnitID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbPaymentSplit ALTER COLUMN UNIT_ID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN UnitID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbCreditNote ALTER COLUMN UNIT_ID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbBankReconcilation ALTER COLUMN UnitID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN UNIT_ID TEXT(25)"
''      Conn1.Execute "ALTER TABLE TemplateUnitSelection ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE tblPurInvSRec ALTER COLUMN UNIT_ID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tblPrevGLU ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE tblPoA ALTER COLUMN UnitID TEXT(25)"
''      Conn1.Execute "ALTER TABLE tblBtRptTran ALTER COLUMN UnitID TEXT(25)"
''      Conn1.Execute "ALTER TABLE DemandRecords ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE DemandRecPreview ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE LeaseHistory ALTER COLUMN UnitNumber TEXT(25)"
''      Conn1.Execute "ALTER TABLE NLPosting ALTER COLUMN UNIT_ID TEXT(25)"
''   Else
''      Rst1.Close
''   End If
''   Set Rst1 = Nothing
''
'''Print #n, "######################### finish chaning unitid field size  ##########################"
'''Close #n
''
''   GoTo ADDNEW_REC_MTYS
''
''MissingTable_RESIZE_UNITID_EVERYWHERE:
'''   MsgBox ERR.description, vbInformation + vbOKOnly, "Col Size - Description of DSR"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new record MTYS on 17/04/14 PrimaryCode
'''###############################################################################################################
''ADDNEW_REC_MTYS:
''   On Error GoTo MissingRec_ADDNEW_MTYS
''
''   With Rst1
''      .Open "SELECT * FROM PRIMARYCODE WHERE CODE = 'MTYS';", Conn1, adOpenDynamic, adLockOptimistic
''
''      If .EOF Then
''         .AddNew
''         .Fields.Item("Code").Value = "MTYS"
''         .Fields.Item("Value").Value = "MAINTENANCE RECORD STATUS"
''         .Update
''         .Close
''         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
''         .AddNew
''         !PrimaryCode = "MTYS"
''         !Code = "PENDING"
''         !Value = "PENDING"
''         .Update
''         .AddNew
''         !PrimaryCode = "MTYS"
''         !Code = "COMPLETED"
''         !Value = "COMPLETED"
''         .Update
''         .AddNew
''         !PrimaryCode = "MTYS"
''         !Code = "URGENT"
''         !Value = "URGENT"
''         .Update
''      End If
''      Rst1.Close
''   End With
''
''   GoTo RESIZE_PropertyName_Property
''
''MissingRec_ADDNEW_MTYS:
'''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Extend the field size PropertyName on 24/04/2014 Property
'''###############################################################################################################
''RESIZE_PropertyName_Property:
'''MsgBox "14"
''   Rst1.Open "SELECT PropertyName FROM Property;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields.Item("PropertyName").DefinedSize < 100 Then
''      Rst1.Close
''      Set Rst1 = Nothing
''
''      Conn1.Execute "ALTER TABLE Property ALTER COLUMN PropertyName TEXT(100)"
''      Conn1.Execute "ALTER TABLE Property ALTER COLUMN ProAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Property ALTER COLUMN ProAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Property ALTER COLUMN ProAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Property ALTER COLUMN ProAddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientOfficeAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientOfficeAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN ClientOfficeAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ADD COLUMN ClientAddressLine4 TEXT(70);"
''      Conn1.Execute "ALTER TABLE Client ADD COLUMN ClientOfficeAddressLine4 TEXT(70);"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN RegAdd1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN RegAdd2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ALTER COLUMN RegAdd3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd4 TEXT(70);"
''
''      Conn1.Execute "ALTER TABLE Units ALTER COLUMN UnitAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Units ALTER COLUMN UnitAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Units ALTER COLUMN UnitAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Units ALTER COLUMN UnitAddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN HOAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN HOAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN HOAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN HOAddressLine4 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN BillAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN BillAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN BillAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN BillAddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordOfficeAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordOfficeAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordOfficeAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Landlord ADD COLUMN LandlordAddressLine4 TEXT(70);"
''      Conn1.Execute "ALTER TABLE Landlord ADD COLUMN LandlordOfficeAddressLine4 TEXT(70);"
''
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentOfficeAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentOfficeAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentOfficeAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Agent ADD COLUMN AgentAddressLine4 TEXT(70);"
''      Conn1.Execute "ALTER TABLE Agent ADD COLUMN AgentOfficeAddressLine4 TEXT(70);"
''
''      Conn1.Execute "ALTER TABLE tlbBank ALTER COLUMN BANK_ADDRESS1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbBank ALTER COLUMN BANK_ADDRESS2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbBank ALTER COLUMN BANK_ADDRESS3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN BANK_ADDRESS4 TEXT(70);"
''
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierAddressLine4 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeAddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeAddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeAddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN SupplierOfficeAddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN AddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN AddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN AddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN AddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE tlbLetterReports ALTER COLUMN AddressLine1 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbLetterReports ALTER COLUMN AddressLine2 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbLetterReports ALTER COLUMN AddressLine3 TEXT(70)"
''      Conn1.Execute "ALTER TABLE tlbLetterReports ALTER COLUMN AddressLine4 TEXT(70)"
''
''      Conn1.Execute "ALTER TABLE LServiceCharges ALTER COLUMN SCDesc TEXT(100)"
''      Conn1.Execute "ALTER TABLE LRentCharges ALTER COLUMN RentDesc TEXT(100)"
''      Conn1.Execute "ALTER TABLE LInsuranceCharges ALTER COLUMN InsDesc TEXT(100)"
''   Else
''      Rst1.Close
''      Set Rst1 = Nothing
''   End If
''
'''  Export Clients into Supplier table on 08/05/2013         ~~~NEVER REMOVED~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''###############################################################################################################
''' fixed by anol 2017 03 20 anol issue 327 procedure was taking long time to run
'''   Rst1.Open "SELECT * " & _
'''             "FROM Client " & _
'''             "WHERE ClientID NOT IN (" & _
'''               "SELECT SupplierID FROM Supplier WHERE TYPE = 'CLIENT');", Conn1, adOpenStatic, adLockReadOnly
''  Rst1.Open "SELECT A.* FROM Client A Left join Supplier B ON A.ClientID=B.SupplierID WHERE ClientID <>SupplierID AND TYPE = 'CLIENT';", Conn1, adOpenStatic, adLockReadOnly
'''MsgBox "15               "
''   If Not Rst1.EOF Then
''      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
''      While Not Rst1.EOF
''         Rst2.AddNew
''         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("ClientID").Value
''         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("ClientName").Value
''         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("ClientAddressLine1").Value
''         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("ClientAddressLine2").Value
''         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("ClientAddressLine3").Value
''         Rst2.Fields.Item("SupplierPostCode").Value = IIf(IsNull(Rst1.Fields.Item("ClientPostCode").Value), "", Rst1.Fields.Item("ClientPostCode").Value)
''         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
''         Rst2.Fields.Item("TYPE").Value = "CLIENT"
''         Rst2.Update
''         Rst1.MoveNext
''      Wend
''      Rst2.Close
''   End If
''   Rst1.Close
'''
''''  Export Agents into Supplier table on 08/05/2013         ~~~ removed to the agent module.  Samrat 05/08/2014
''''###############################################################################################################
'''   Rst1.Open "SELECT * " & _
'''             "FROM Agent " & _
'''             "WHERE AgentID NOT IN (" & _
'''               "SELECT SupplierID FROM Supplier WHERE TYPE = 'AGENT');", Conn1, adOpenStatic, adLockReadOnly
''''MsgBox "16               "
'''   If Not Rst1.EOF Then
'''      Rst2.Open "SELECT * FROM Supplier", Conn1, adOpenDynamic, adLockOptimistic
'''      While Not Rst1.EOF
'''         Rst2.AddNew
'''         Rst2.Fields.Item("SupplierID").Value = Rst1.Fields.Item("AgentID").Value
'''         Rst2.Fields.Item("SupplierName").Value = Rst1.Fields.Item("AgentName").Value
'''         Rst2.Fields.Item("SupplierAddressLine1").Value = Rst1.Fields.Item("AgentAddressLine1").Value
'''         Rst2.Fields.Item("SupplierAddressLine2").Value = Rst1.Fields.Item("AgentAddressLine2").Value
'''         Rst2.Fields.Item("SupplierAddressLine3").Value = Rst1.Fields.Item("AgentAddressLine3").Value
'''         Rst2.Fields.Item("SupplierPostCode").Value = Rst1.Fields.Item("AgentPostCode").Value
'''         Rst2.Fields.Item("VATReg").Value = Rst1.Fields.Item("VATReg").Value
'''         Rst2.Fields.Item("TYPE").Value = "AGENT"
'''         Rst2.Update
'''         Rst1.MoveNext
'''      Wend
'''      Rst2.Close
'''   End If
'''   Rst1.Close
'''   Set Rst2 = Nothing
''
'''   Extend the field size PropertyID on 30/04/2014 PropertyInsurance
'''###############################################################################################################
''EXTEND_COL_PropertyID_Schedule:
''
''   Rst1.Open "SELECT PropertyID FROM PropertyInsurance;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields.Item("PropertyID").DefinedSize < 12 Then
''      Rst1.Close
''      Set Rst1 = Nothing
''      Conn1.Execute "ALTER TABLE PropertyInsurance ALTER COLUMN PropertyID TEXT(12);"
''   Else
''      Rst1.Close
''      Set Rst1 = Nothing
''   End If
''
'''   Add new record JOBS on 09/05/14 PrimaryCode
'''###############################################################################################################
''ADDNEW_REC_JOBS:
''   On Error GoTo MissingRec_ADDNEW_JOBS
'''MsgBox "17"
''   With Rst1
''      .Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'TTP' AND Value = 'JOBS';", Conn1, adOpenDynamic, adLockOptimistic
''
''      If .EOF Then
''         .AddNew
''         .Fields.Item("PrimaryCode").Value = "TTP"
''         .Fields.Item("Code").Value = "4"
''         .Fields.Item("Value").Value = "JOBS"
''         .Update
''      End If
''      .Close
''   End With
''
''   GoTo ADDNEW_REC_NewPO
''
''MissingRec_ADDNEW_JOBS:
'''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new record PO_NEW on 09/05/14 PrimaryCode
'''###############################################################################################################
''ADDNEW_REC_NewPO:
''   On Error GoTo MissingRec_ADDNEW_NewPO
'''MsgBox "18"
''   With Rst1
''      .Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'TTP' AND Value = 'NewPO';", Conn1, adOpenDynamic, adLockOptimistic
''
''      If .EOF Then
''         .AddNew
''         .Fields.Item("PrimaryCode").Value = "TTP"
''         .Fields.Item("Code").Value = "5"
''         .Fields.Item("Value").Value = "NewPO"
''         .Update
''      End If
''      .Close
''   End With
''
''   GoTo ADDNEW_COL_PO_tblPurInv
''
''MissingRec_ADDNEW_NewPO:
'''   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column PO on 24/08/2014 tblPurInv
'''###############################################################################################################
''ADDNEW_COL_PO_tblPurInv:
''   On Error GoTo MISSING_ADDNEW_COL_PO_tblPurInv
''
''   Rst1.Open "SELECT PO FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_UnitNumber_PropertyMaintHistory
''
''MISSING_ADDNEW_COL_PO_tblPurInv:
''   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN PO TEXT(25);"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''   Add new column UnitNumber on 24/08/2014 PropertyMaintHistory
'''###############################################################################################################
''ADDNEW_COL_UnitNumber_PropertyMaintHistory:
''   On Error GoTo MISSING_ADDNEW_COL_UnitNumber_PropertyMaintHistory
''
''   Rst1.Open "SELECT UnitNumber FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_Invoiced_tblPurInvSRec
''
''MISSING_ADDNEW_COL_UnitNumber_PropertyMaintHistory:
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN UnitNumber TEXT(12);"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''Invoiced (Yes/No) need to add in the blank database            19/09/2014    Samrat
'''   Add new column Invoiced on 19/09/2014 tblPurInvSRec
'''###############################################################################################################
''ADDNEW_COL_Invoiced_tblPurInvSRec:
''   On Error GoTo MISSING_ADDNEW_COL_Invoiced_tblPurInvSRec
''
''   Rst1.Open "SELECT Invoiced FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
''   GoTo ADDNEW_COL_PoPiCross_tblPurInvSRec
''
''MISSING_ADDNEW_COL_Invoiced_tblPurInvSRec:
''   Conn1.Execute "ALTER TABLE tblPurInvSRec ADD COLUMN Invoiced BIT;"
''
''   UpdateDatabase2 = 1
''   Exit Function
''
'''PoPiCross (Text 25) need to add in the blank database            20/09/2014    Samrat
'''   Add new column PoPiCross on 20/09/2014 tblPurInvSRec
'''###############################################################################################################
''ADDNEW_COL_PoPiCross_tblPurInvSRec:
''   On Error GoTo MISSING_ADDNEW_COL_PoPiCross_tblPurInvSRec
''
''   Rst1.Open "SELECT PoPiCross FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''
'' GoTo Modify_TABLE_tlbPayment
''
''MISSING_ADDNEW_COL_PoPiCross_tblPurInvSRec:
''   Conn1.Execute "ALTER TABLE tblPurInvSRec ADD COLUMN PoPiCross TEXT(25);"
''
''   UpdateDatabase2 = 1
''   Exit Function
'''Resolved by BOSL LIVE
'''Resolved by BOSL
'''0000468: Posting dates not implemented correctly
'''468 note 792
'''added by anol anol 23 Nov 2014
''Modify_TABLE_tlbPayment:
''   On Error GoTo Mod_TABLE_tlbPayment
''   Rst1.Open "SELECT PostingDate FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''
''      GoTo NEW_TABLE_MemoDetails
''   Else
''Mod_TABLE_tlbPayment:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "Update tlbpayment set postingdate=PDate where isdate(postingdate)=false and len(postingdate)>0"
''      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN PostingDate DateTime;"
''      GoTo NEW_TABLE_MemoDetails
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
'''Resolved by BOSL
'''issue 474:  Adding Memo at Jobs and maintenance
'''added by anol  01 Dec 2014
''NEW_TABLE_MemoDetails:
''   On Error GoTo MissingTable_MemoDetails
''
''   Rst1.Open "SELECT * FROM MemoDetails;", Conn1, adOpenStatic, adLockReadOnly
''      If Rst1.Fields("SageAccountNumber").DefinedSize <> 14 Then
''            Rst1.Close
''            Conn1.Execute "Drop TABLE MemoDetails;"
''              Conn1.Execute _
''            "CREATE TABLE MemoDetails " & _
''               "(" & _
''                  "MemoID      NUMBER, " & _
''                  "MemoType    Text(100), " & _
''                  "SageAccountNumber Text(14), " & _
''                  "MemoDescription    memo, " & _
''                   "userName    Text(100), " & _
''                    "UpdateTime    DATETIME, " & _
''                  "PRIMARY KEY(MemoID,SageAccountNumber) " & _
''               ");"
''               'below two lines were added by anol 07 Jan 2015
''               UpdateDatabase2 = 1
''               Exit Function
''      Else
''         Rst1.Close
''      End If
''
''   GoTo Modify_TABLE_DemandRecords
''
''MissingTable_MemoDetails:
''   Conn1.Execute _
''      "CREATE TABLE MemoDetails " & _
''         "(" & _
''            "MemoID      NUMBER, " & _
''            "MemoType    Text(100), " & _
''            "SageAccountNumber Text(14), " & _
''            "MemoDescription    memo, " & _
''             "userName    Text(100), " & _
''              "UpdateTime    DATETIME, " & _
''            "PRIMARY KEY(MemoID,SageAccountNumber) " & _
''         ");"
''
''   UpdateDatabase2 = 1
''   'below one lines were added by anol 07 Jan 2015
''   Exit Function
''   GoTo Modify_TABLE_DemandRecords
''ADD_PropertyAnalysis1:
''   Exit Function
''
'''Resolved by BOSL
'''0000468: Posting dates not implemented correctly
'''468 note 792
'''DemandRecords,tlbReceipt,tblPurInv,tlbBankPayment,tblBatchReceipt,tblBtRptTran,tblBatchReceipt,NJ_Header
'''added by anol 07 Dec 2014
''Modify_TABLE_DemandRecords:
''   On Error GoTo Mod_TABLE_DemandRecords
''   Rst1.Open "SELECT PostingDate FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly
''
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_tlbReceipt
''   Else
''Mod_TABLE_DemandRecords:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE DemandRecords ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tlbReceipt
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''   '468 note 792 tlbReceipt
''Modify_TABLE_tlbReceipt:
''   On Error GoTo Mod_TABLE_tlbReceipt
''   Rst1.Open "SELECT PostingDate FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_tblPurInv
''   Else
''Mod_TABLE_tlbReceipt:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tblPurInv
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''    '468 note 792 tblPurInv
''Modify_TABLE_tblPurInv:
''   On Error GoTo Mod_TABLE_tblPurInv
''   Rst1.Open "SELECT PostingDate FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_tlbBankPayment
''   Else
''Mod_TABLE_tblPurInv:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE tblPurInv ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tlbBankPayment
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''    '468 note 792 tlbBankPayment
''Modify_TABLE_tlbBankPayment:
''   On Error GoTo Mod_TABLE_tlbBankPayment
''   Rst1.Open "SELECT PostingDate FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_tblBatchReceipt
''   Else
''Mod_TABLE_tlbBankPayment:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tblBatchReceipt
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''   '468 note 792 tblBatchReceipt
''Modify_TABLE_tblBatchReceipt:
''   On Error GoTo Mod_TABLE_tblBatchReceipt
''   'Here the error comes
''   Rst1.Open "SELECT PostingDate FROM tblBatchReceipt;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_tblBtRptTran
''   Else
'''Mod_TABLE_tblBatchReceipt:
'''modifed by anol 11 Dec 2014
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      'fixed
''      Conn1.Execute "Update tblBatchReceipt set postingdate=BRDate where isDate(postingdate)=false and len(postingdate)>0"
''      Conn1.Execute "ALTER TABLE tblBatchReceipt ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tblBtRptTran
''Mod_TABLE_tblBatchReceipt:
''      Conn1.Execute "ALTER TABLE tblBatchReceipt add COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tblBtRptTran
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''   '468 note 792 tblBtRptTran
''Modify_TABLE_tblBtRptTran:
''   On Error GoTo Mod_TABLE_tblBtRptTran
''   Rst1.Open "SELECT PostingDate FROM tblBtRptTran;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_TABLE_NJ_Header
''   Else
''Mod_TABLE_tblBtRptTran:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      'fixed
''      Conn1.Execute "Update tblBtRptTran set postingdate=DueDate where isDate(postingdate)=false and len(postingdate)>0"
''      Conn1.Execute "ALTER TABLE tblBtRptTran ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_NJ_Header
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''   '' DemandRecords,tlbReceipt,tblPurInv,tlbBankPayment,tblBatchReceipt,tblBtRptTran,tblBatchReceipt,
''   'NJ_Header,tblBatchTransaction
''   '468 note 792 NJ_Header
''Modify_TABLE_NJ_Header:
''   On Error GoTo Mod__NJ_Header
''   Rst1.Open "SELECT PostingDate FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''       GoTo Modify_TABLE_tblBatchTransaction
''   Else
''Mod__NJ_Header:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE NJ_Header ALTER COLUMN PostingDate DateTime;"
''      GoTo Modify_TABLE_tblBatchTransaction
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''    '468 note 792 tblBatchTransaction
''Modify_TABLE_tblBatchTransaction:
''   On Error GoTo Mod__tblBatchTransaction
''   Rst1.Open "SELECT PostingDate FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly
''   If Rst1.Fields(0).Type = 135 Then
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      GoTo Modify_table_PropertyMaintHistory
''      'Exit Function
''   Else
''Mod__tblBatchTransaction:
''      If Rst1.State = 1 Then
''         Rst1.Close
''      End If
''      Conn1.Execute "ALTER TABLE tblBatchTransaction ALTER COLUMN PostingDate DateTime;"
''      Exit Function
''      'GoTo Modify_TABLE_PropertyAnalysis
''   End If
''   UpdateDatabase2 = 1
''   Exit Function
''Modify_table_PropertyMaintHistory:
''   On Error GoTo Mod_table_PropertyMaintHistory
''   Rst1.Open "SELECT LastModified FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''   GoTo Add_Column_Cap
''Mod_table_PropertyMaintHistory:
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory add COLUMN LastModified Text(100);"
''   Conn1.Execute "ALTER TABLE PropertyMaintHistory add COLUMN ModifiedBy Text(100);"
''   UpdateDatabase2 = 1
''   Exit Function
''Add_Column_Cap:
''   On Error GoTo Missing_Column_Cap
''   Rst1.Open "SELECT CapAmount FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''   GoTo Missing_Column_PIRef_tlbPaymentSplit
''   Exit Function
''Missing_Column_Cap:
''   Conn1.Execute "ALTER TABLE LeaseDetails Add column CapAmount Number"
''    UpdateDatabase2 = 1
''    Exit Function
''    'Bellow parts has been added by anol 29 Apr 2015
''    'Purchase report was not showing correct entries for single payments.
''Missing_Column_PIRef_tlbPaymentSplit:
''   On Error GoTo Mod_Column_PIRef_tlbPaymentSplit
''   Rst1.Open "SELECT PIRef FROM tlbPaymentSplit;", Conn1, adOpenStatic, adLockReadOnly
''   Rst1.Close
''   GoTo Missing_Column_Consolidated_tlbClientBank
''Mod_Column_PIRef_tlbPaymentSplit:
''    Conn1.Execute "ALTER TABLE tlbPaymentSplit Add column PIRef Text(100)"
''    UpdateDatabase2 = 1
''    Exit Function
''    'Bellow parts has been added by anol 04 May 2015
''    'Adding a column at tlbClientBank for consolidated YES/No.
''Missing_Column_Consolidated_tlbClientBank:
''    On Error GoTo Mod_Column_Consolidated_tlbClientBank
''    Rst1.Open "SELECT consolidated FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
''    Rst1.Close
''    GoTo tlbLetterReports_coulmnL_Check
''    Exit Function
''Mod_Column_Consolidated_tlbClientBank:
''    Conn1.Execute "ALTER TABLE tlbClientBanks Add column consolidated BIT"
''    UpdateDatabase2 = 1
''    Exit Function
''    'Bellow parts has been added by anol 02 Jun 2015
''    'Increasing column size of Unitno on tlbLetterReports table
tlbLetterReports_coulmnL_Check:
   ' On Error GoTo Mod_Column_tlbLetterReports
    Rst1.Open "SELECT UnitNo FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly
    If Rst1.Fields.Item("UnitNo").DefinedSize < 12 Then
        Rst1.Close
        Conn1.Execute "ALTER TABLE tlbLetterReports ALTER COLUMN UnitNo TEXT(12)"
        UpdateDatabase2 = 1
        Exit Function
    Else
        Rst1.Close
        GoTo MissingRec_MAGENT
    End If
'adding Primary Code MAGENT and LLORD
'Written by anol 1 July 2015
'5. A primary code of "MAGENT" with the value "MANAGING AGENT" should be created for the managing agent.'
'This should be added to the database under supplier type every time a managing agent is created the "TYPE"'
'field in the supplier table will be populated with "AGENT".'
'6. A primary code of "LLORD" with the value "LANDLORD" should be created for the Landlord. This should be'
'added to the database under supplier type every time a landlord is created the "TYPE" field in the'
'supplier table will be populated with "LLORD".
MissingRec_MAGENT:
   With Rst1
      .Open "SELECT Code FROM PrimaryCode WHERE Code = 'MAGENT';", Conn1, adOpenStatic, adLockReadOnly
      If .EOF Then
         .Close
         .Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         !Code = "MAGENT"
         !Value = "MANAGING AGENT"
         .Update
         .AddNew
         !Code = "LLORD"
         !Value = "LANDLORD"
         .Update
         .Close
         .Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields.Item(0).Value = "MAGENT"
         .Fields.Item(1).Value = "MA"
         .Fields.Item(2).Value = "MANAGING AGENT"
         .Fields.Item(3).Value = "MANAGING AGENT"
         .Update
         .AddNew
         .Fields.Item(0).Value = "LLORD"
         .Fields.Item(1).Value = "LL"
         .Fields.Item(2).Value = "LANDLORD"
         .Fields.Item(3).Value = "LANDLORD"
         .Update
         .Close
         UpdateDatabase2 = 1
        Exit Function
      Else
         .Close
        GoTo DataLenChangeDirectLine2
        
      End If
     End With
DataLenChangeDirectLine2:
   Rst1.Open "Select DirectLine2 from tenants", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("DirectLine2").DefinedSize = 20 Then
        Rst1.Close
        Conn1.Execute "ALTER TABLE tenants ALTER COLUMN DirectLine2 text(40)"
        Conn1.Execute "ALTER TABLE tenants ALTER COLUMN HOTelephone text(40)"
        Conn1.Execute "ALTER TABLE tenants ALTER COLUMN BillFax text(40)"
        Conn1.Execute "ALTER TABLE tenants ALTER COLUMN HOFax text(40)"

        UpdateDatabase2 = 1
        Exit Function
  Else
    Rst1.Close
    GoTo ChangeDataType_DemandSplPreview
   End If
   'Below part has been added by anol 02 Nov 2015
ChangeDataType_DemandSplPreview:
   Rst1.Open "Select TypeOfDemand from DemandSplPreview", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("TypeOfDemand").DefinedSize = 1 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE DemandSplPreview ALTER COLUMN TypeOfDemand INTEGER"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
       GoTo Change_sizeDirectLine1_tenant
   End If
    'Below part has been added by anol 03 Nov 2015
Change_sizeDirectLine1_tenant:
   Rst1.Open "Select DirectLine1 from Tenants", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("DirectLine1").DefinedSize = 20 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN DirectLine1 text(40)"
       Conn1.Execute "ALTER TABLE Tenants ALTER COLUMN BillTelephone text(40)"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
   End If
    'Below part has been added by anol 18 Jan 2016
Modify_TABLE_DemandRecPreview:
   On Error GoTo Mod_TABLE_DemandRecPreview
   Rst1.Open "SELECT spare1 FROM DemandRecPreview;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).Type <> 202 Then
      If Rst1.State = 1 Then
         Rst1.Close
      End If
   Else
Mod_TABLE_DemandRecPreview:
      If Rst1.State = 1 Then
         Rst1.Close
      End If
      Conn1.Execute "ALTER TABLE DemandRecPreview ALTER COLUMN spare1 LONG;"
      UpdateDatabase2 = 1
   End If
Modify_TABLE_tlbDRCURRENTPRINT:
   On Error GoTo Mod_TABLE_tlbDRCURRENTPRINT
   Rst1.Open "SELECT Typeofdemand  FROM tlbDRCURRENTPRINT ;", Conn1, adOpenStatic, adLockReadOnly
    If Rst1.Fields.Item("TypeOfDemand").DefinedSize = 1 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE tlbDRCURRENTPRINT ALTER COLUMN Typeofdemand INTEGER;"
       UpdateDatabase2 = 1
       Exit Function
    End If
Mod_TABLE_tlbDRCURRENTPRINT:
    If Rst1.State = 1 Then
         Rst1.Close
    End If
ChangeDataType_tblPrevGLU:
   Rst1.Open "Select DmdType from tblPrevGLU", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("DmdType").DefinedSize = 1 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE tblPrevGLU ALTER COLUMN DmdType INTEGER"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
       GoTo ADD_NLPost_tlbPayment1
       Exit Function
   End If
   'added by anol 20160415
ADD_NLPost_tlbPayment1:
    On Error GoTo Mod_ADD_NLPost_tlbPayment
     Rst1.Open "Select NLPost from tlbPayment", Conn1, adOpenKeyset, adLockReadOnly
    If Rst1.Fields.Item("NLPost").DefinedSize = 2 Then
        Rst1.Close
        GoTo ADD_FUND_FundList
    Else
Mod_ADD_NLPost_tlbPayment:
        If Rst1.State = 1 Then
               Rst1.Close
        End If
        Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN NLPost BIT"
        UpdateDatabase2 = 1
        Exit Function
   End If
    'added by anol 20160429
ADD_FUND_FundList:
     On Error GoTo MOD_FUND_FundList
     Rst1.Open "Select FundList from FUND", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_Tenants_SLControl
MOD_FUND_FundList:
      Conn1.Execute "ALTER Table FUND ADD column FundList text(255);"
      UpdateDatabase2 = 1
      Exit Function
ADD_Tenants_SLControl:
     On Error GoTo MOD_Tenants_SLControl
     Rst1.Open "Select SLControl from Tenants", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_Tenants_DefaultNC
MOD_Tenants_SLControl:
      Conn1.Execute "ALTER Table Tenants ADD column SLControl text(15);"
      UpdateDatabase2 = 1
      Exit Function
ADD_Tenants_DefaultNC:
     On Error GoTo MOD_Tenants_DefaultNC
     Rst1.Open "Select DefaultNC from Tenants", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_Tenants_VAT_CODE
MOD_Tenants_DefaultNC:
      Conn1.Execute "ALTER Table Tenants ADD column DefaultNC text(15);"
      UpdateDatabase2 = 1
      Exit Function
ADD_Tenants_VAT_CODE:
    On Error GoTo MOD_Tenants_VAT_CODE
     Rst1.Open "Select VAT_CODE from Tenants", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_ClientID_tlbReceipt

MOD_Tenants_VAT_CODE:
     Conn1.Execute "ALTER Table Tenants ADD column VAT_CODE text(5);"
     UpdateDatabase2 = 1
     Exit Function
ADD_ClientID_tlbReceipt:
    On Error GoTo MOD_ClientID_tlbReceipt
     Rst1.Open "Select ClientID from tlbReceipt", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_ClientID_tlbClientBanks
     Exit Function
MOD_ClientID_tlbReceipt:
     Conn1.Execute "ALTER Table tlbReceipt ADD column ClientID text(10);"
     Conn1.Execute "Update tlbReceipt,Units,Property SET tlbReceipt.ClientID= Property.ClientID where tlbReceipt.UnitID=Units.UnitNumber AND Units.PropertyID=Property.PropertyID "
     UpdateDatabase2 = 1
     Exit Function
ADD_ClientID_tlbClientBanks:
     On Error GoTo MOD_ClientID_tlbClientBanks
     Rst1.Open "Select GroupCode from Client;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_SO_DemandRecords
     Exit Function
MOD_ClientID_tlbClientBanks:
     Conn1.Execute "ALTER Table Client ADD column GroupCode text(10);"
     UpdateDatabase2 = 1
     Exit Function
'implementation of standing order boolean field
ADD_SO_DemandRecords:
     On Error GoTo MOD_SO_DemandRecords
     Rst1.Open "Select SO from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_UserID_DemandRecords
     Exit Function
MOD_SO_DemandRecords:
    Conn1.Execute "ALTER TABLE DemandRecords ADD SO BIT"
    Conn1.Execute "Update DemandRecords Set SO=false"
    UpdateDatabase2 = 1
    Exit Function
'ADD_table_RecordLocking:
'     On Error GoTo MOD_table_RecordLocking
'     Rst1.Open "Select ClientID from RecordLocking;", Conn1, adOpenKeyset, adLockReadOnly
'     Rst1.Close
'     GoTo ADD_UserID_DemandRecords
'     Exit Function
'MOD_table_RecordLocking:
'    Conn1.Execute "Create TABLE RecordLocking " & _
'         "(" & _
'            "SCREEN TEXT(100), " & _
'            "WorkStation     TEXT(250), " & _
'            "USER     TEXT(250), " & _
'            "TableName      TEXT(100), " & _
'            "ClientID  TEXT(100), " & _
'            "BankCode      TEXT(20) " & _
'            ");"
'     UpdateDatabase2 = 1
'     Exit Function
ADD_UserID_DemandRecords:
     On Error GoTo MOD_UserID_DemandRecords
     Rst1.Open "Select UserID from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_SSL_ShoppingCenter
     Exit Function
MOD_UserID_DemandRecords:
    Conn1.Execute "ALTER TABLE DemandRecords ADD UserID text(100);"
    Conn1.Execute "ALTER TABLE DemandRecords ADD SystemID text(100);"
    UpdateDatabase2 = 1
    Exit Function
ADD_SSL_ShoppingCenter:
    'added by anol 20161125
     On Error GoTo MOD_SSL_ShoppingCenter
     Rst1.Open "Select SSL from ShoppingCentre;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Add_OurRef_template
     Exit Function
MOD_SSL_ShoppingCenter:
    Conn1.Execute "ALTER TABLE ShoppingCentre ADD SSL BIT;"
    Conn1.Execute "Update ShoppingCentre Set SSL=false"
    UpdateDatabase2 = 1
    Exit Function
Add_OurRef_template:
    'added by anol 20170219 letter template does not have our ref
     On Error GoTo Mod_OurRef_template
     Rst1.Open "Select OurRef from Template;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Fix_TEmail_tlbLetterReports
     Exit Function
Mod_OurRef_template:
     Conn1.Execute "ALTER TABLE Template ADD OurRef text(200);"
     UpdateDatabase2 = 1
   'added by anol 20170219 letter template Email address short issue 335
Fix_TEmail_tlbLetterReports:
'     On Error GoTo Mod_OurRef_template
     Rst1.Open "Select TEmail from tlbLetterReports;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields.Item("TEmail").DefinedSize = 40 Then
        Rst1.Close
        GoTo Mod_TEmail_tlbLetterReports
     Else
        Rst1.Close
     End If
     GoTo Fix_Value_tblPrevGLU
Mod_TEmail_tlbLetterReports:
     Conn1.Execute "ALTER TABLE tlbLetterReports Alter Column TEmail text(100);"
     UpdateDatabase2 = 1
     Exit Function
  'added by anol 20170609 update lessee service charge amount was not supporting decimal places up to 8,issue 403
Fix_Value_tblPrevGLU:
     Rst1.Open "Select Value from tblPrevGLU;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields.Item("Value").Precision = 19 Then
        Rst1.Close
        GoTo Mod_Value_tblPrevGLU
     Else
        Rst1.Close
        GoTo ADD_MaxPurChaseHist_ShoppingCenter
     End If
     Exit Function
Mod_Value_tblPrevGLU:
     Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column [Value] Double;"
      Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column ExValue Double;"
       Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column NewValue Double;"
        Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column ExYrTotal Double;"
         Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column NewYrTotal Double;"
          Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column ExDueEachPeriod Double;"
           Conn1.Execute "ALTER TABLE tblPrevGLU Alter Column NewDueEachPeriod Double;"
     UpdateDatabase2 = 1
     Exit Function
ADD_MaxPurChaseHist_ShoppingCenter:
    'added by anol 20171111 issue 504
     On Error GoTo MOD_MaxPurChaseHist_ShoppingCenter
     Rst1.Open "Select MaxPurChaseHist from ShoppingCentre;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Modify_tlbDRCurrentPrint_DUEDATE
     'Exit Function
MOD_MaxPurChaseHist_ShoppingCenter:
    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxPurChaseHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxPurChaseHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxPurPaymentHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxPurPaymentHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxSupplierHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxSupplierHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxDemandHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxDemandHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxReceiptHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxReceiptHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxLesseeHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxLesseeHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxCashBookHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxCashBookHist=0"

    Conn1.Execute "ALTER TABLE ShoppingCentre ADD MaxNominalHist Long;"
    Conn1.Execute "Update ShoppingCentre Set MaxNominalHist=0"
    UpdateDatabase2 = 1
    Exit Function
Modify_tlbDRCurrentPrint_DUEDATE:
   On Error GoTo Mod__tlbDRCurrentPrint_DUEDATE
   Rst1.Open "SELECT DUEDATE FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).Type = 135 Then
      If Rst1.State = 1 Then
         Rst1.Close
      End If
      GoTo Modify_LeaseBreaches_DeleteFlag
   Else
Mod__tlbDRCurrentPrint_DUEDATE:
      If Rst1.State = 1 Then
         Rst1.Close
      End If
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN DUEDATE DateTime;"
   End If
   UpdateDatabase2 = 1
   Exit Function
Modify_LeaseBreaches_DeleteFlag:
    On Error GoTo Mod_LeaseBreaches_DeleteFlag
    Rst1.Open "SELECT DeleteFlag FROM LeaseBreaches;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_tlbDRCurrentPrint_MINDUEDATE
Mod_LeaseBreaches_DeleteFlag:
    Conn1.Execute "ALTER TABLE LeaseBreaches ADD DeleteFlag text(200);"
    UpdateDatabase2 = 1
    Exit Function
    'added by anol 2018 06 15 issue 601
Modify_tlbDRCurrentPrint_MINDUEDATE:
    On Error GoTo Mod_tlbDRCurrentPrint_MINDUEDATE
    Rst1.Open "SELECT MINDUEDATE FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_LeaseBreaches_LeaseMemo
Mod_tlbDRCurrentPrint_MINDUEDATE:
    Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD MINDUEDATE DATETIME;"
    UpdateDatabase2 = 1
    Exit Function
Modify_LeaseBreaches_LeaseMemo:
    On Error GoTo Mod_LeaseBreaches_LeaseMemo
    Rst1.Open "SELECT LeaseMemo FROM LeaseBreaches;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_Tenants_RCCComments2
Mod_LeaseBreaches_LeaseMemo:
    Conn1.Execute "ALTER TABLE LeaseBreaches ADD LeaseMemo Memo;"
    UpdateDatabase2 = 1
    Exit Function
Modify_Tenants_RCCComments2:
    On Error GoTo Mod_Tenants_RCCComments2
    Rst1.Open "SELECT RCCComments2 FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_Client_Comments1
Mod_Tenants_RCCComments2:
    Conn1.Execute "ALTER TABLE Tenants ADD COLUMN RCCComments2 TEXT(250);"
    UpdateDatabase2 = 1
    Exit Function
Modify_Client_Comments1:
    On Error GoTo Mod_Client_Comments1
    Rst1.Open "SELECT Comments1 FROM Client;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_Tenants_isCurrent
Mod_Client_Comments1:
    Conn1.Execute "ALTER TABLE Client ADD COLUMN Comments1 TEXT(250);"
    Conn1.Execute "ALTER TABLE Client ADD COLUMN Comments2 TEXT(250);"
    UpdateDatabase2 = 1
    Exit Function
Modify_Tenants_isCurrent: 'added on 2018/10/21 issue 656
    On Error GoTo Mod_Tenants_isCurrent
    Rst1.Open "SELECT isCurrent FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo Modify_ShoppingCenter_PIShowBal
Mod_Tenants_isCurrent:
    Conn1.Execute "ALTER TABLE Tenants ADD COLUMN isCurrent BIT;"
    Conn1.Execute "ALTER TABLE Tenants ADD COLUMN CurrUnit TEXT(12);"
    UpdateDatabase2 = 1
    Exit Function
Modify_ShoppingCenter_PIShowBal: 'added on 2018/11/12 issue 677
    On Error GoTo Mod_ShoppingCenter_PIShowBal
    Rst1.Open "SELECT PIShowBal FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_table_ReportLAChistory
Mod_ShoppingCenter_PIShowBal:
    Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN PIShowBal BIT;"
    Conn1.Execute "Update ShoppingCentre Set PIShowBal = True ;"
    Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN CurrentVersion TEXT(100);"
    UpdateDatabase2 = 1
    Exit Function
ADD_table_ReportLAChistory:
     On Error GoTo MOD_table_ReportLAChistory
     Rst1.Open "Select * from ReportLAChistory;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_ActualINV_ReportSAChistory
MOD_table_ReportLAChistory:
    Conn1.Execute "Create TABLE ReportLAChistory " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "SIGN text(1), " & _
            "transactionID LONG, " & _
            "Type Number, " & _
            "Type_desc text(255), " & _
            "PF text(255), " & _
            "slnumber LONG, " & _
            "Rdate datetime, " & _
            "Details text(255), " & _
            "extref text(255), " & _
            "amount currency, " & _
            "Osamount  currency, " & _
            "Balance currency, " & _
            "isMaster number " & _
            ");"
    Conn1.Execute "Create TABLE ReportSAChistory " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "SIGN text(1), " & _
            "transactionID LONG, " & _
            "Type Number, " & _
            "Type_desc text(255), " & _
            "PF text(255), " & _
            "Pdate datetime, " & _
            "Details text(255), " & _
            "extref text(255), " & _
            "amount currency, " & _
            "Osamount  currency, " & _
            "Balance  currency, " & _
            "flag number, " & _
            "isMaster number, " & _
            "ClientID Text(10) " & _
            ");"
      UpdateDatabase2 = 1
      Exit Function
ADD_ActualINV_ReportSAChistory:
      On Error GoTo Mod_ActualINV_ReportSAChistory
      Rst1.Open "Select ActualINV from ReportSAChistory;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo ADD_ClientOfficeAddressLine4_Client
      Exit Function
Mod_ActualINV_ReportSAChistory:
      Conn1.Execute "ALTER TABLE ReportSAChistory ADD COLUMN ActualINV TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function
ADD_ClientOfficeAddressLine4_Client: 'added 20190105 issue 702
      On Error GoTo Mod_ClientOfficeAddressLine4_Client
      Rst1.Open "Select ClientOfficeAddressLine4 from Client;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo ADD_AgentAddressLine4_Agent
Mod_ClientOfficeAddressLine4_Client:
      Conn1.Execute "ALTER TABLE Client ADD COLUMN ClientOfficeAddressLine4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function

ADD_AgentAddressLine4_Agent: 'added 20190105 issue 702
      On Error GoTo Mod_AgentAddressLine4_Agent
      Rst1.Open "Select AgentAddressLine4 from Agent;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo ADD_AgentOfficeAddressLine4_Agent
Mod_AgentAddressLine4_Agent:
      Conn1.Execute "ALTER TABLE Agent ADD COLUMN AgentAddressLine4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function

ADD_AgentOfficeAddressLine4_Agent:
      On Error GoTo Mod_AgentOfficeAddressLine4_Agent
      Rst1.Open "Select AgentOfficeAddressLine4 from Agent;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
     GoTo ADD_LandlordOfficeAddressLine4_Landlord
Mod_AgentOfficeAddressLine4_Agent:
      Conn1.Execute "ALTER TABLE Agent ADD COLUMN AgentOfficeAddressLine4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function

ADD_LandlordOfficeAddressLine4_Landlord: 'added 20190105 issue 702
      On Error GoTo Mod_LandlordOfficeAddressLine4_Landlord
      Rst1.Open "Select LandlordOfficeAddressLine4 from Landlord;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo ADD_LandlordAddressLine4_Landlord
Mod_LandlordOfficeAddressLine4_Landlord:
      Conn1.Execute "ALTER TABLE Landlord ADD COLUMN LandlordOfficeAddressLine4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function
ADD_LandlordAddressLine4_Landlord:       'added 20190105 issue 702
      On Error GoTo Mod_LandlordAddressLine4_Landlord
      Rst1.Open "Select LandlordAddressLine4 from Landlord;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo ADD_RegAdd4_Client
Mod_LandlordAddressLine4_Landlord:
      Conn1.Execute "ALTER TABLE Landlord ADD COLUMN LandlordAddressLine4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function
ADD_RegAdd4_Client:       'added 20190105 issue 702
      On Error GoTo Mod_RegAdd4_Client
      Rst1.Open "Select RegAdd4 from Client;", Conn1, adOpenKeyset, adLockReadOnly
      Rst1.Close
      GoTo Modify_NJHeader_NJDate
Mod_RegAdd4_Client:
      Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd4 TEXT(255);"
      UpdateDatabase2 = 1
      Exit Function
Modify_NJHeader_NJDate:
     On Error GoTo Mod_NJHeader_NJDate
     Rst1.Open "SELECT NJDate FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
     If Rst1.Fields(0).Type = 135 Then '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Modify_tlbPayment_isPrint
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Mod_NJHeader_NJDate
     End If
     Exit Function
Mod_NJHeader_NJDate:
    Conn1.Execute "ALTER TABLE NJ_Header ALTER COLUMN NJDate Datetime;"
    UpdateDatabase2 = 1
    Exit Function

Modify_tlbPayment_isPrint: 'this field is needed for Print payment list report in PI form
     On Error GoTo Mod_tlbPayment_isPrint
     Rst1.Open "SELECT isPrint FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     GoTo Modify_Landlord_LandlordID
     Exit Function
     
Mod_tlbPayment_isPrint:
    Debug.Print time
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN isPrint BIT;"
    UpdateDatabase2 = 1
    Exit Function
Modify_Landlord_LandlordID:     'this field is needed for Print payment list report in PI form
     'On Error GoTo Mod_Landlord_LandlordID ' this is no fixed by occuring error
     Rst1.Open "SELECT LandlordID FROM Landlord;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     If Rst1.Fields.Item("LandlordID").DefinedSize = 10 Then  '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Mod_Landlord_LandlordID
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Modify_Agent_AgentID
     End If
     Exit Function
Mod_Landlord_LandlordID:
    Conn1.Execute "ALTER TABLE Landlord ALTER COLUMN LandlordID Text(15);"
    UpdateDatabase2 = 1
    Exit Function
    
Modify_Agent_AgentID:         'this field is needed for Print payment list report in PI form
     'On Error GoTo Mod_Agent_AgentID' this is no fixed by occuring error
     Rst1.Open "SELECT AgentID FROM Agent;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     If Rst1.Fields.Item("AgentID").DefinedSize = 10 Then  '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Mod_Agent_AgentID
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If
'         GoTo Mod_NJHeader_NJDate 'everthing is fine just finishing this procedure
        GoTo Modify_tlbPayment_DateTimeStamp
     End If
     Exit Function
Mod_Agent_AgentID:
    Conn1.Execute "ALTER TABLE Agent ALTER COLUMN AgentID Text(15);"
    UpdateDatabase2 = 1
    Exit Function
    
Modify_tlbPayment_DateTimeStamp: 'this fields shall be used to solve server multiuser conflict problem issue 749
     On Error GoTo Mod_tlbPayment_DateTimeStamp
     Rst1.Open "SELECT DateTimeStamp FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     GoTo Modify_DemandRecords_DateTimeStamp
     Exit Function
Mod_tlbPayment_DateTimeStamp:
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN DateTimeStamp text(100);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN Module text(100);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN UserSessionID text(25);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN WindowsUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN MachineName text(100);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN PrestigeUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN ServerIPaddress text(100);"
    UpdateDatabase2 = 1
    Exit Function

Modify_DemandRecords_DateTimeStamp: 'this fields shall be used to solve server multiuser conflict problem issue 749
On Error GoTo Mod_DemandRecords_DateTimeStamp
     Rst1.Open "SELECT DateTimeStamp FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     GoTo Modify_NJ_Header_DateTimeStamp
     Exit Function
Mod_DemandRecords_DateTimeStamp:
    'we do not need those locking in demandrecords we are doing it by tlbReceipt
    
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN DateTimeStamp text(100);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN Module text(100);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN UserSessionID text(25);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN WindowsUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN MachineName text(100);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN PrestigeUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN ServerIPaddress text(100);"
    
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN DateTimeStamp text(100);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN Module text(100);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN UserSessionID text(25);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN WindowsUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN MachineName text(100);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN PrestigeUserName text(100);"
    Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN ServerIPaddress text(100);"
    
  
    UpdateDatabase2 = 1
    Exit Function
Modify_NJ_Header_DateTimeStamp:     'this fields shall be used to solve server multiuser conflict problem issue 749
On Error GoTo Mod_NJ_Header_DateTimeStamp
     Rst1.Open "SELECT DateTimeStamp FROM NJ_Header;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     GoTo Modify_tlbReceipt_DateTimeType
     Exit Function
Mod_NJ_Header_DateTimeStamp:
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN DateTimeStamp text(100);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN Module text(100);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN UserSessionID text(25);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN WindowsUserName text(100);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN MachineName text(100);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN PrestigeUserName text(100);"
        Conn1.Execute "ALTER TABLE NJ_Header ADD COLUMN ServerIPaddress text(100);"
        UpdateDatabase2 = 1
Modify_tlbReceipt_DateTimeType:
        Rst1.Open "SELECT DateTimeStamp FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly
        If Rst1.Fields(0).Type = 135 Then '202 means type text and 135 means datetime
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo Mod_tlbReceipt_DateTimeType
        Else
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo Modify_tlbBankPayment_DateTimeType
        End If
        Exit Function
Mod_tlbReceipt_DateTimeType:
        Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN DateTimeStamp text(100);"
        UpdateDatabase2 = 1
        Exit Function
Modify_tlbBankPayment_DateTimeType:
        Rst1.Open "SELECT DateTimeStamp FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly
        If Rst1.Fields(0).Type = 135 Then '202 means type text and 135 means datetime
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo Mod_tlbBankPayment_DateTimeType
        Else
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo Modify_GlobalInsurance_ID
        End If
       ' Exit Function
Mod_tlbBankPayment_DateTimeType:
        Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN DateTimeStamp text(100);"
        UpdateDatabase2 = 1
        Exit Function
Modify_GlobalInsurance_ID:
        Rst1.Open "SELECT ID FROM GlobalInsurance;", Conn1, adOpenStatic, adLockReadOnly
        If Rst1.Fields(0).Type = 17 Then '3 means type long integer and 17 means Byte
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo Mod_GlobalInsurance_ID
        Else
            If Rst1.State = 1 Then
               Rst1.Close
            End If
            GoTo ADD_ProcessFileLoc_tlbClientBanks
        End If
        Exit Function
Mod_GlobalInsurance_ID:
        Conn1.Execute "ALTER TABLE GlobalInsurance ALTER COLUMN ID Int;"
        UpdateDatabase2 = 1
        Exit Function
        '*******adding field ProcessFileLoc 2020-02-12 by anol
ADD_ProcessFileLoc_tlbClientBanks:
    On Error GoTo MOD_ProcessFileLoc_tlbClientBanks
    Rst1.Open "SELECT ProcessFileLoc FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_table_BACSPaymentRun

MOD_ProcessFileLoc_tlbClientBanks:
    Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN ProcessFileLoc TEXT(255);"
    UpdateDatabase2 = 1
    Exit Function
  '*******adding field ProcessFileLoc 2020-02-23 by anol
ADD_table_BACSPaymentRun:
     On Error GoTo MOD_table_BACSPaymentRun
     Rst1.Open "Select * from BACSPaymentRun;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo NEW_TABLE_ConsolidatedBankList
     Exit Function
MOD_table_BACSPaymentRun:
    Conn1.Execute "Create TABLE BACSPaymentRun " & _
         "(" & _
            "RunNo LONG NOT NULL, " & _
            "RunDate Date," & _
            "LineNo LONG NOT NULL, " & _
            "EB INT, " & _
            "Description Text(255) " & _
            ");"
     UpdateDatabase2 = 1
     Exit Function
'###############################################################################################################
NEW_TABLE_ConsolidatedBankList:
   On Error GoTo CreateTable_ConsolidatedBankList

   Rst1.Open "SELECT * FROM ConsolidatedBankList;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo ADD_vatOptionEnabled_GlobalData
   Exit Function
   
CreateTable_ConsolidatedBankList:
   Conn1.Execute _
      "CREATE TABLE ConsolidatedBankList " & _
         "(" & _
            "conBankID      LONG NOT NULL PRIMARY KEY, " & _
            "BankCode      TEXT(10) NOT NULL, " & _
            "BankName      TEXT(255) NOT NULL, " & _
            "BankACNumber      TEXT(255) NOT NULL, " & _
            "SortCode    TEXT(255) NOT NULL, " & _
            "StatementDate    TEXT(255), " & _
            "ClosingBal    Currency, " & _
            "SOB   Currency " & _
         ");"
         
        Conn1.Execute "ALTER TABLE  tlbClientBanks ADD COLUMN ConsolidatedBankID LONG;"
        Conn1.Execute "ALTER TABLE  tlbClientBanks ADD COLUMN ConsBankACNumber TEXT(255);"
        Conn1.Execute "ALTER TABLE  tlbClientBanks ADD COLUMN ConsSortCode TEXT(255);"
        Conn1.Execute "ALTER TABLE  tlbClientBanks ADD COLUMN conBankCode TEXT(10);"
        Conn1.Execute "ALTER TABLE  tlbClientBanks ADD COLUMN conBankReadOnly LONG;"
    
        UpdateDatabase2 = 1
   Exit Function
ADD_vatOptionEnabled_GlobalData:
    On Error GoTo MOD_vatOptionEnabled_GlobalData
    Rst1.Open "SELECT vatOptionEnabled FROM GlobalData;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_OnReturn_tlbVatCode
    Exit Function
MOD_vatOptionEnabled_GlobalData:
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN vatOptionEnabled LONG;"
    Conn1.Execute "Update GlobalData set vatOptionEnabled=1;"
    UpdateDatabase2 = 1
    Exit Function
    
ADD_OnReturn_tlbVatCode:
    On Error GoTo MOD_OnReturn_tlbVatCode
    Rst1.Open "SELECT OnReturn FROM tlbVatCode;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_Supplier_OptedtoTax
    Exit Function
MOD_OnReturn_tlbVatCode:
    Conn1.Execute "ALTER TABLE tlbVatCode ADD COLUMN OnReturn BIT;"
    Conn1.Execute "Update tlbVatCode set OnReturn=true;"
    UpdateDatabase2 = 1
    Exit Function
ADD_Supplier_OptedtoTax:
    On Error GoTo MOD_Supplier_OptedtoTax
    Rst1.Open "SELECT OptedtoTax FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_FinancialYear_Setascurrent
    Exit Function
MOD_Supplier_OptedtoTax:
    Conn1.Execute "ALTER TABLE Supplier ADD COLUMN OptedtoTax BIT;"
    Conn1.Execute "Update Supplier set OptedtoTax=true;"
     Conn1.Execute "Update Supplier set OptedtoTax=false where VATCode ='';"
     Conn1.Execute "Update Supplier set OptedtoTax=false where VATCode=true ;"
      Conn1.Execute "Update Supplier set OptedtoTax=false where VATCode is null ;"
     
'    Conn1.Execute "Update Supplier set OptedtoTax=false where VATCode ='' OR VATCode=-1 or VATCode is null;"
    UpdateDatabase2 = 1
    Exit Function
    
ADD_FinancialYear_Setascurrent:
    On Error GoTo MOD_FinancialYear_Setascurrent
    Rst1.Open "SELECT setascurrent FROM FinancialYear;", Conn1, adOpenStatic, adLockReadOnly
    Rst1.Close
    GoTo ADD_FundMatrix
    Exit Function
MOD_FinancialYear_Setascurrent:
    Conn1.Execute "ALTER TABLE FinancialYear ADD COLUMN setascurrent BIT;"
    UpdateDatabase2 = 1
    Exit Function
ADD_FundMatrix:
        On Error GoTo CreateTable_FundMatrix
        Rst1.Open "SELECT * FROM FundMatrix;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_ShoppingCentre_isFundassign
        Exit Function
CreateTable_FundMatrix:
        Conn1.Execute _
         "create table FundMatrix" & _
         "(  ID AUTOINCREMENT," & _
         "   ClientId text(10)," & _
         "   PropertyID text(4)," & _
         "   FundID Number," & _
         "   FundName text(100)," & _
         "   FundCategory text(100)," & _
         "   FundCode  text(12)," & _
         "   isDeleted BIT" & _
         ");"
         UpdateDatabase2 = 1
   Exit Function
ADD_ShoppingCentre_isFundassign:
        On Error GoTo MOD_ShoppingCentre_isFundassign
        Rst1.Open "SELECT isFundassign FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_ShoppingCentre_isRestrictedtoBudget
        Exit Function
MOD_ShoppingCentre_isFundassign:
        Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN isFundassign BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_ShoppingCentre_isRestrictedtoBudget:
        On Error GoTo MOD_ShoppingCentre_isRestrictedtoBudget
        Rst1.Open "SELECT isRestrictedtoBudget FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandType_Consolidated
        Exit Function
MOD_ShoppingCentre_isRestrictedtoBudget:
        Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN isRestrictedtoBudget BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_DemandType_Consolidated:
        On Error GoTo MOD_DemandType_Consolidated
        Rst1.Open "SELECT Consolidated FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_LServiceCharges_SCYEDate
        Exit Function
MOD_DemandType_Consolidated:
        Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN Consolidated BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_LServiceCharges_SCYEDate:
        On Error GoTo MOD_LServiceCharges_SCYEDate
        Rst1.Open "SELECT SCYEDate FROM LServiceCharges;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_ChargeTypes_propertyID
        Exit Function
MOD_LServiceCharges_SCYEDate:
        Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN SCYEDate Date;"
        UpdateDatabase2 = 1
        Exit Function
ADD_ChargeTypes_propertyID:
        On Error GoTo MOD_ChargeTypes_PropertyID
        Rst1.Open "Select PropertyID from ChargeTypes;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_PayableTypes_PropertyID
        Exit Function
MOD_ChargeTypes_PropertyID:
        Conn1.Execute "ALTER TABLE ChargeTypes ADD COLUMN PropertyID Text(4);"
        Conn1.Execute "ALTER TABLE ChargeTypes ADD COLUMN ClientID Text(10);"
        Conn1.Execute "ALTER TABLE ChargeTypes ADD COLUMN GroupID INT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_PayableTypes_PropertyID:
        On Error GoTo MOD_PayableTypes_PropertyID
        Rst1.Open "Select PropertyID from PayableTypes;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_ClientProAgr_agreementstartdate
        Exit Function
MOD_PayableTypes_PropertyID:
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN PropertyID Text(4);"
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN ClientID Text(10);"
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN GroupID INT;"
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN isUseControlAccount BIT;"
        
        UpdateDatabase2 = 1
        Exit Function
'added on 2020-12-15
ADD_ClientProAgr_agreementstartdate:
        On Error GoTo MOD_ClientProAgr_agreementstartdate
        Rst1.Open "Select agreementstartdate from ClientProAgr;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_ClientGlobalData_dueDate
        Exit Function
MOD_ClientProAgr_agreementstartdate:
        Conn1.Execute "ALTER TABLE ClientProAgr ADD COLUMN agreementStartDate Date;"
        Conn1.Execute "ALTER TABLE ClientProAgr ADD COLUMN agreementEndDate Date;"
        UpdateDatabase2 = 1
        Exit Function
ADD_ClientGlobalData_dueDate:
        On Error GoTo MOD_ClientGlobalData_dueDate
        Rst1.Open "Select dueDate from ClientGlobalData;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbPayable_clientLandlordID
        Exit Function
MOD_ClientGlobalData_dueDate:
        Conn1.Execute "ALTER TABLE ClientGlobalData ADD COLUMN dueDate Date;"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbPayable_clientLandlordID:
        On Error GoTo MOD_tlbPayable_clientLandlordID
        Rst1.Open "Select clientLandlordID from tlbPayable;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbAgreement_ManagingAgentID
        Exit Function
MOD_tlbPayable_clientLandlordID:
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN clientLandlordID Text(10);"
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN AmountOrPercentage  Text(2);"
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN Percentage  Double;"
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN StopDate  Date;"
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN ONDD  BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbAgreement_ManagingAgentID:
        On Error GoTo MOD_tlbAgreement_ManagingAgentID
        Rst1.Open "Select ManagingAgentID from tlbAgreement;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_Secondarycode_Code
        'tlbAgreement
MOD_tlbAgreement_ManagingAgentID:
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN ManagingAgentID Text(10);"
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN TotalAmount  Double;"
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN EachPeriod  Double;"
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN StopDate  Date;"
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN CapAmount  Double;"
        UpdateDatabase2 = 1
        Exit Function

ADD_Secondarycode_Code:
        'There we are not creating goto by  error generation
     Rst1.Open "SELECT Code FROM SecondaryCode;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close

     If Rst1.Fields.Item("Code").DefinedSize = 10 Then  '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo MOD_Secondarycode_Code
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If

         GoTo ADD_FUND_FUNDCode
     End If
     Exit Function
MOD_Secondarycode_Code:
    'added by anol 2021-01-08
     Conn1.Execute "ALTER TABLE SecondaryCode ALTER COLUMN Code  TEXT(15);"
     Conn1.Execute "ALTER TABLE NominalLedger ALTER COLUMN CAType  TEXT(2);"
     UpdateDatabase2 = 1
'write code to populate TenantDEPOSIT to update records procedure
'Following code added on 2021-01-22
ADD_FUND_FUNDCode:
        'There we are not creating goto by  error generation
     Rst1.Open "SELECT FUNDCode FROM Fund;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close

     If Rst1.Fields.Item("FUNDCode").DefinedSize = 12 Then  '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo MOD_FUND_FUNDCode
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If

         GoTo ADD_table_RentSummaryStatement
     End If
     Exit Function
MOD_FUND_FUNDCode:
        Conn1.Execute "ALTER TABLE Fund ALTER COLUMN FUNDCode  TEXT(15);"
        UpdateDatabase2 = 1
        Exit Function
ADD_table_RentSummaryStatement:
     On Error GoTo MOD_table_RentSummaryStatement
     Rst1.Open "Select * from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_tlbPayable_ClientID
'     Exit Function
MOD_table_RentSummaryStatement:
    Conn1.Execute "Create TABLE RentSummaryStatement " & _
        "(" & _
               "StatementID Number, statementNo number,ClientIDLandlordID Text(10)," & _
               "BankCode  text(15),PreviousStatementDate Date, " & _
               "StatementDate date,StatementOpBal  Double, Retentions Double," & _
               "RetenstionDescription Text(250),ClearRetentions BIT,AccrualsAcBalance  Double, " & _
               "SupplierAcBalance  Double,ManagingAgentAcBalance  Double,LandlordACBalance  Double, " & _
               "ClientACBalance     Double,ListOfFundId   Text(250),ListOfPayableTypeID Text(250),ListOfinputProperties Text(250), " & _
               "TenantDepositsReceived  Double, " & _
               "AvailableFunds  Double,PaymentsonAccount  Double, PayableAmount Double,StatementClosingBal  Double," & _
               "ClientPayments  Double,LandlordPayments  Double, ManagingAgentPayments Double, " & _
               "TenantReceipts  Double,SupplierPayments Double,BankPaymentReceipts  Double,ClientLandlordBalance Double, " & _
               "PINumber Text(250), PITransactionID Double,Generated_Date Date," & _
               "Printed BIT,Emailed BIT,Invoiced BIT,PostToHistory BIT " & _
       ");"
       Conn1.Execute "Create TABLE RentSummaryStatementPreview " & _
        "(" & _
               "StatementID Number, statementNo number,ClientIDLandlordID Text(10)," & _
               "BankCode  text(15),PreviousStatementDate Date, " & _
               "StatementDate date,StatementOpBal  Double, Retentions Double," & _
               "RetenstionDescription Text(250),ClearRetentions BIT,AccrualsAcBalance  Double, " & _
               "SupplierAcBalance  Double,ManagingAgentAcBalance  Double,LandlordACBalance  Double, " & _
               "ClientACBalance     Double,ListOfFundId   Text(250),ListOfPayableTypeID Text(250),ListOfinputProperties Text(250), " & _
               "TenantDepositsReceived  Double, " & _
               "AvailableFunds  Double,PaymentsonAccount  Double, PayableAmount Double, StatementClosingBal  Double, " & _
               "ClientPayments  Double,LandlordPayments  Double, ManagingAgentPayments Double," & _
               "TenantReceipts  Double,SupplierPayments Double, BankPaymentReceipts  Double,ClientLandlordBalance Double, " & _
               "PINumber Text(250), PITransactionID Double,Generated_Date Date," & _
               "Printed BIT,Emailed BIT,Invoiced BIT,PostToHistory BIT " & _
       ");"
         Conn1.Execute "Create TABLE RetentionDetails " & _
        "(" & _
               "StatementID Number, SlNumber number,Description Text(250),amount Double,isCleared BIT" & _
        ");"

        Conn1.Execute "ALTER TABLE Fund ALTER COLUMN FUNDCode  TEXT(15);"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbPayable_ClientID:
         On Error GoTo MOD_tlbPayable_ClientID
         Rst1.Open "Select ClientID from  tlbPayable"", Conn1, adOpenKeyset, adLockReadOnly"
         Rst1.Close
         GoTo ADD_Client_RentSummaryTemplate
         'Do not use goto until you have a new section, just exit
         Exit Function
MOD_tlbPayable_ClientID:
        Conn1.Execute "ALTER Table tlbPayable add column ClientID text(10);"
        Conn1.Execute "ALTER Table tlbPayable add column Printed BIT;"
        Conn1.Execute "ALTER Table tlbPayable add column Emailed BIT;"
        Conn1.Execute "ALTER Table tlbPayable add column Invoiced BIT;"
        UpdateDatabase2 = 1
        Exit Function
                
ADD_Client_RentSummaryTemplate:
        On Error GoTo MOD_Client_RentSummaryTemplate
        Rst1.Open "Select RentSummaryTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Supplier_RentSummaryTemplate
'     Exit Function
MOD_Client_RentSummaryTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN RentSummaryTemplate  TEXT(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_Supplier_RentSummaryTemplate:
        On Error GoTo MOD_Supplier_RentSummaryTemplate
        Rst1.Open "Select RentSummaryTemplate from Supplier;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Client_PaymentType
        Exit Function
MOD_Supplier_RentSummaryTemplate:
        Conn1.Execute "ALTER TABLE Supplier ADD COLUMN RentSummaryTemplate  TEXT(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_Client_PaymentType:
        On Error GoTo MOD_Client_PaymentType
        Rst1.Open "Select PaymentType from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Client_RemittanceTemplate
        Exit Function
MOD_Client_PaymentType:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN PaymentType  TEXT(10);"
        Conn1.Execute "ALTER TABLE Client ADD COLUMN PaymentTerms  INT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_Client_RemittanceTemplate:
        On Error GoTo MOD_Client_RemittanceTemplate
        Rst1.Open "Select RemittanceTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Supplier_RemittanceTemplate
        Exit Function
MOD_Client_RemittanceTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN RemittanceTemplate  TEXT(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_Supplier_RemittanceTemplate:
        On Error GoTo MOD_Supplier_RemittanceTemplate
        Rst1.Open "Select RemittanceTemplate from Supplier;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_table_tblPurInvPreview
        Exit Function
MOD_Supplier_RemittanceTemplate:
        Conn1.Execute "ALTER TABLE Supplier ADD COLUMN RemittanceTemplate  TEXT(250);"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_table_tblPurInvPreview:
     On Error GoTo MOD_table_tblPurInvPreview
     Rst1.Open "Select * from tblPurInvPreview;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_tlbAgreement_FDD
     Exit Function
MOD_table_tblPurInvPreview:
    Conn1.Execute "Create TABLE tblPurInvPreview " & _
        "(" & _
               "MY_ID Text(25),SlNumber Long,Supp_AC Text(10),TRAN_DATE Date," & _
               "TransactionType int,Inv_no Text(20),TOTAL_AMOUNT Currency,UPDATE_SAGE BIT, " & _
               "DR_CR BIT,RECHARGED BIT,TTP  int,CATEGORY_CODE   int," & _
               "History  BIT,CL_ID  Text(10),TrfPayment BIT,ScheduleID  Long, " & _
               "Prn  Text(1),PropertyID Text(10),DueDate  Date,NLPost  BIT, " & _
               "PostingDate  Date,PO Text(25)" & _
       ");"
       
     Conn1.Execute "Create TABLE tblPurInvSRecPreview " & _
        "(" & _
               "MY_ID TEXT(25),ParentID  TEXT(50),TRAN_ID TEXT(20),TRANS TEXT(4),UNIT_ID TEXT(25),NOMINAL_CODE text(10)," & _
               "DEPT_ID Long,PROJ_REF  TEXT(8),COST_CODE TEXT(10),JOB_ID TEXT(10),DESCRIPTION TEXT(255),NET_AMOUNT Currency, " & _
               "TAX_CODE TEXT(3),VAT Currency,TOTAL_AMOUNT Currency,UPDATE_SAGE BIT,RECHARGED INT,CATEGORY_CODE INT," & _
               "ScheduleID LONG,CL_ID TEXT(10),RecoverablePt INT,TrfPayment  BIT,Invoiced BIT,PoPiCross Text(25) " & _
       ");"
       
       UpdateDatabase2 = 1
       Exit Function
ADD_tlbAgreement_FDD:
        On Error GoTo MOD_tlbAgreement_FDD
        Rst1.Open "Select FDD from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_GlobalData_QDueDate1
        Exit Function
MOD_tlbAgreement_FDD:
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN FDD  Date;"
        UpdateDatabase2 = 1
        Exit Function
ADD_GlobalData_QDueDate1:
        On Error GoTo MOD_GlobalData_QDueDate1
        Rst1.Open "Select QDueDate1 from GlobalData;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_isManagementFee
        Exit Function
MOD_GlobalData_QDueDate1:
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN QDueDate1 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN QDueDate2 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN QDueDate3 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN QDueDate4 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN HYDueDate1 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN HYDueDate2 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate1 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate2 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate3 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate4 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate5 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate6 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate7 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate8 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate9 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate10 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate11 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN MDueDate12 Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN YDueDate  Text(12);"
        Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN NoOfDaysToSendMFB4Due INT;"
        
        UpdateDatabase2 = 1
        Exit Function
ADD_tblPurInv_isManagementFee:
        On Error GoTo MOD_tblPurInv_isManagementFee
        Rst1.Open "Select isManagementFee from tblPurInv;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo Modify_END_DATE_tlbAgreement
        Exit Function
        
MOD_tblPurInv_isManagementFee:
        Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN isManagementFee BIT;" 'by default they are false
        UpdateDatabase2 = 1
        Exit Function
Modify_END_DATE_tlbAgreement:
     Rst1.Open "SELECT END_DATE FROM tlbAgreement;", Conn1, adOpenStatic, adLockReadOnly
     If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
    'No need after this because this is the last section
    GoTo ADD_tlbAgreement_amount
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Mod_tlbAgreement_END_DATE
     End If
Exit Function
Mod_tlbAgreement_END_DATE:
    Conn1.Execute "ALTER TABLE tlbAgreement ALTER COLUMN StopDate TEXT(15);"
    Conn1.Execute "ALTER TABLE tlbAgreement ALTER COLUMN END_DATE TEXT(15);"
    UpdateDatabase2 = 1
    Exit Function
         'tlbAgreement amount
ADD_tlbAgreement_amount:
        On Error GoTo MOD_tlbAgreement_amount
        Rst1.Open "Select amount from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbBankPayment_RentSumStatement
        Exit Function
        
MOD_tlbAgreement_amount:
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN amount Currency;" 'by default they are false
        UpdateDatabase2 = 1
        Exit Function
        
'ADD_tlbReceipt_RentSumStatement:
'        On Error GoTo MOD_tlbReceipt_RentSumStatement
'        Rst1.Open "Select RentSumStatement from tlbReceipt;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        GoTo ADD_tlbReceipt_RentSumStatementPreview
'        Exit Function
'
'MOD_tlbReceipt_RentSumStatement:
'        Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN RentSumStatement text(100);" 'by default they are false
'        UpdateDatabase2 = 1
'        Exit Function

'ADD_tlbReceipt_RentSumStatementPreview:
'        On Error GoTo MOD_tlbReceipt_RentSumStatementPreview
'        Rst1.Open "Select RentSumStatementPreview from tlbReceipt;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        GoTo ADD_tlbPayment_RentSumStatement
'        Exit Function
'
'MOD_tlbReceipt_RentSumStatementPreview:
'        Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN RentSumStatementPreview text(100);" 'by default they are false
'        UpdateDatabase2 = 1
'        Exit Function
'ADD_tlbPayment_RentSumStatement:
'        On Error GoTo MOD_tlbPayment_RentSumStatement
'        Rst1.Open "Select RentSumStatement from tlbPayment;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        GoTo ADD_tlbPayment_RentSumStatementPreview
'        Exit Function
'
'MOD_tlbPayment_RentSumStatement:
'        Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN RentSumStatement text(100);"
'        UpdateDatabase2 = 1
'        Exit Function
        
'ADD_tlbPayment_RentSumStatementPreview:
'        On Error GoTo MOD_tlbPayment_RentSumStatementPreview
'        Rst1.Open "Select RentSumStatementPreview from tlbPayment;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        GoTo ADD_tlbBankPayment_RentSumStatement
'        Exit Function
'
'MOD_tlbPayment_RentSumStatementPreview:
'        Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN RentSumStatementPreview text(100);"
'        UpdateDatabase2 = 1
'        Exit Function
        '********************************************
        
ADD_tlbBankPayment_RentSumStatement:
        On Error GoTo MOD_tlbBankPayment_RentSumStatement
        Rst1.Open "Select RentSumStatement from tlbBankPayment;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbBankPayment_RentSumStatementPreview
        Exit Function
        
MOD_tlbBankPayment_RentSumStatement:
        Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN RentSumStatement text(100);"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbBankPayment_RentSumStatementPreview:
        On Error GoTo MOD_tlbBankPayment_RentSumStatementPreview
        Rst1.Open "Select RentSumStatementPreview from tlbBankPayment;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_ManagementFeeSL
        Exit Function

MOD_tlbBankPayment_RentSumStatementPreview:
        Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN RentSumStatementPreview text(100);"
        UpdateDatabase2 = 1
        Exit Function
        
'ADD_tlbReceipt_ISMGTFEE:
'        On Error GoTo MOD_tlbReceipt_ISMGTFEE
'        Rst1.Open "Select ISMGTFEE from tlbReceipt;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        GoTo ADD_tblPurInv_ManagementFeeSL
'        Exit Function
'
'MOD_tlbReceipt_ISMGTFEE:
'        Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN ISMGTFEE  BIT;" 'by default they are false
'        Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN PIREFMGTFEE  text(100);" '
'        Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN ChargeDate Date;" '
'        UpdateDatabase2 = 1
'        Exit Function
        
        
ADD_tblPurInv_ManagementFeeSL:
        On Error GoTo MOD_tblPurInv_ManagementFeeSL
        Rst1.Open "Select ManagementFeeSL from tblPurInv;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbAgreement_LastChargeDate
        Exit Function
        
MOD_tblPurInv_ManagementFeeSL:
        Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN ManagementFeeSL LONG;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbAgreement_LastChargeDate:
        On Error GoTo MOD_tlbAgreement_ManagementFeeSL
        Rst1.Open "Select LastChargeDate from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Client_ClientAddressLine5
        Exit Function
        
MOD_tlbAgreement_ManagementFeeSL:
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN LastChargeDate Date;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_Client_ClientAddressLine5:
        On Error GoTo MOD_Client_ClientAddressLine5
        Rst1.Open "Select ClientAddressLine5 from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo Change_FundMatrix_FundCode
        Exit Function
        
MOD_Client_ClientAddressLine5:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN ClientAddressLine5 TEXT(250);"
        Conn1.Execute "ALTER TABLE Client ADD COLUMN ClientOfficeAddressLine5 TEXT(250);"
        Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd5 TEXT(250);"
        UpdateDatabase2 = 1
        Exit Function
        
'ADD_FundMatrix_FundCode:
'        On Error GoTo MOD_FundMatrix_FundCode
'        Rst1.Open "Select FundCode from FundMatrix;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
''        GoTo Modify_END_DATE_tlbAgreement  tlbAgreement
'        Exit Function
'

     
Change_FundMatrix_FundCode:
   Rst1.Open "Select FundCode from FundMatrix", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("FundCode").DefinedSize = 12 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE FundMatrix ALTER COLUMN FundCode TEXT(15)"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
       GoTo Change_tlbBankReconcilation_AccountNum
   End If
   
Change_tlbBankReconcilation_AccountNum:
   Rst1.Open "Select AccountNum from tlbBankReconcilation", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("AccountNum").DefinedSize < 30 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE tlbBankReconcilation ALTER COLUMN AccountNum TEXT(30)"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
       GoTo Change_tblPrevGLU_SAN
   End If


Change_tblPrevGLU_SAN:
   Rst1.Open "Select SAN from tblPrevGLU", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("SAN").DefinedSize < 30 Then
       Rst1.Close
       Conn1.Execute "ALTER TABLE tblPrevGLU ALTER COLUMN SAN  TEXT(30)"
       UpdateDatabase2 = 1
       Exit Function
   Else
       Rst1.Close
       GoTo Change_PropertyID_length
   End If
   
Change_PropertyID_length:
   Rst1.Open "Select PropertyID from Property", Conn1, adOpenKeyset, adLockReadOnly
   If Rst1.Fields.Item("PropertyID").DefinedSize < 10 Then
        Rst1.Close
        Set Rst1 = Nothing
        Conn1.Execute "ALTER TABLE Property ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE ClientGlobalData ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE ClientProAgr ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE DemandTypes ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE GlobalData ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE GlobalInsurance ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE GlobalRC  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE GlobalSC ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE InterestRates  ALTER COLUMN PropertyID text(10)"
        'when you have a big amount of data the following commant resulting in a timeout, you need to do this on fresh data or there is less data
        ' Or you have to chage the field size manually by hand
'        Conn1.Execute "ALTER TABLE NLPosting ALTER COLUMN PROPERTY_ID text(10)"
        Conn1.Execute "ALTER TABLE PropertyAnalysis  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE PropertyInsurance  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE PropertyLandlord ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE PropertyMaintHistory  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE PropertySafety  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE PropertyUtilities  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tblBatchPayment  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tblBatchReceipt  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tblBatchTransaction  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tblPurInv  ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN PropertyID text(10)"
        Conn1.Execute "ALTER TABLE tlbPayment  ALTER COLUMN UnitID text(10)"
        Conn1.Execute "ALTER TABLE tlbPaymentSplit ALTER COLUMN TRANS text(10)"
        Conn1.Execute "ALTER TABLE FundMatrix ALTER COLUMN PropertyID text(10)"
        UpdateDatabase2 = 1
        Exit Function
   Else
        Rst1.Close
        Set Rst1 = Nothing
   End If
   

ADD_TENANTDEPOSIT_DptType:
        Rst1.Open "Select DptType from TENANTDEPOSIT;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("DptType").DefinedSize < 11 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE TENANTDEPOSIT ALTER COLUMN DptType  text(30);" 'by default they are false
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
ADD_tlbBankPayment_PROJ_REF:
        Rst1.Open "Select PROJ_REF from tlbBankPayment;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("PROJ_REF").DefinedSize < 21 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN PROJ_REF  text(100);" 'by default they are false
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
ADD_tblPurInvSRec_TRANS:
        Rst1.Open "Select TRANS from tblPurInvSRec;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("TRANS").DefinedSize < 10 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE tblPurInvSRec ALTER COLUMN TRANS  text(10);"
                Conn1.Execute "ALTER TABLE tblPurInvSRecPreview ALTER COLUMN TRANS  text(10);"
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
ADD_ChargeTypes_propertyID1:
        Rst1.Open "Select propertyID from ChargeTypes;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("propertyID").DefinedSize < 10 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE ChargeTypes ALTER COLUMN propertyID  text(10);"
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
ADD_PayableTypes_propertyID1:
        Rst1.Open "Select propertyID from PayableTypes;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("propertyID").DefinedSize < 10 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE PayableTypes ALTER COLUMN propertyID  text(10);"
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
'ADD_tlbAgreement_DEMAND_TYPE:
'        Rst1.Open "Select DEMAND_TYPE from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
'        If Rst1.Fields.Item("DEMAND_TYPE").DefinedSize = 1 Then
'                Rst1.Close
'                Conn1.Execute "ALTER TABLE tlbAgreement ALTER COLUMN DEMAND_TYPE  INT;"
'                UpdateDatabase2 = 1
'                Exit Function
'        End If
'        Rst1.Close

ADD_RentSummaryStatement_BankPayment:
        On Error GoTo MOD_RentSummaryStatement_BankPayment
        Rst1.Open "Select BankPayment from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbPayable_PayeeType
        Exit Function
MOD_RentSummaryStatement_BankPayment:
        Conn1.Execute "ALTER TABLE RentSummaryStatement add BankPayment Double;"
        Conn1.Execute "ALTER TABLE RentSummaryStatement add BankReceipts  Double;"
        Conn1.Execute "ALTER TABLE RentSummaryStatementPreview add BankPayment Double;"
        Conn1.Execute "ALTER TABLE RentSummaryStatementPreview add BankReceipts  Double;"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbPayable_PayeeType:
        On Error GoTo MOD_tlbPayable_PayeeType
        Rst1.Open "Select PayeeType from tlbPayable;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbPayable_AmountOrPercentage
        Exit Function
MOD_tlbPayable_PayeeType:
        Conn1.Execute "ALTER TABLE tlbPayable ADD COLUMN PayeeType Text(50);"
         UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbPayable_AmountOrPercentage:
        Rst1.Open "Select AmountOrPercentage from tlbPayable;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item("AmountOrPercentage").DefinedSize = 2 Then
                Rst1.Close
                Conn1.Execute "ALTER TABLE tlbPayable ALTER COLUMN AmountOrPercentage  text(10);"
                UpdateDatabase2 = 1
                Exit Function
        End If
        Rst1.Close
ADD_RentSummaryStatementPreview_BankACBalancePreview:
        On Error GoTo MOD_RentSummaryStatementPreview_BankACBalancePreview
        Rst1.Open "Select  BankACBalancePreview from RentSummaryStatementPreview;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RentSummaryStatement_BankACBalance
        Exit Function
MOD_RentSummaryStatementPreview_BankACBalancePreview:
        Conn1.Execute "ALTER TABLE RentSummaryStatementPreview add BankACBalancePreview Double;"
        UpdateDatabase2 = 1
        Exit Function
ADD_RentSummaryStatement_BankACBalance:
        On Error GoTo MOD_RentSummaryStatement_BankACBalance
        Rst1.Open "Select  BankACBalance from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Clients_ConsolidatedStatement
        Exit Function
MOD_RentSummaryStatement_BankACBalance:
        Conn1.Execute "ALTER TABLE RentSummaryStatement add BankACBalance Double;"
        UpdateDatabase2 = 1
        Exit Function
ADD_Clients_ConsolidatedStatement:
        On Error GoTo MOD_Clients_ConsolidatedStatement
        Rst1.Open "Select  ConsolidatedStatement from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Clients_StClientHomeTel
        Exit Function
MOD_Clients_ConsolidatedStatement:
        Conn1.Execute "ALTER TABLE Client add ConsolidatedStatement int;"
        Conn1.Execute "Update Client set ConsolidatedStatement =1;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_Clients_StClientHomeTel:
        On Error GoTo MOD_Clients_StClientHomeTel
        Rst1.Open "Select  StClientHomeTel from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Clients_StToClientAddress
        Exit Function
MOD_Clients_StClientHomeTel:
        Conn1.Execute "ALTER TABLE Client add StClientHomeTel text(250);"
        Conn1.Execute "ALTER TABLE Client add StClientOfficeTel text(250);"
        Conn1.Execute "ALTER TABLE Client add StClientMobile text(250);"
        Conn1.Execute "ALTER TABLE Client add StClientPersonalEmail text(250);"
        Conn1.Execute "ALTER TABLE Client add StClientOfficeEmail text(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_Clients_StToClientAddress:
        On Error GoTo MOD_Clients_StToClientAddress
        Rst1.Open "Select  StToClientAddress from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Clients_AccountName
        Exit Function
MOD_Clients_StToClientAddress:
        Conn1.Execute "ALTER TABLE Client add StToClientAddress INT;"
        Conn1.Execute "Update Client set StToClientAddress =1;"
        Conn1.Execute "ALTER TABLE Client add StToStatementAddress INT;"
        Conn1.Execute "Update Client set StToStatementAddress =1;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_Clients_AccountName:
        On Error GoTo MOD_Clients_AccountName
        Rst1.Open "Select  AccountName from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_PayableTypes_PrintTemplate
        Exit Function
MOD_Clients_AccountName:
        Conn1.Execute "ALTER TABLE Client add AccountName text(200);"
        Conn1.Execute "ALTER TABLE Client add SortCode text(20);"
        Conn1.Execute "ALTER TABLE Client add AccountNumber text(200);"
        Conn1.Execute "ALTER TABLE Client ADD COLUMN UsePayableTemplate INT;"
        Conn1.Execute "ALTER TABLE Client ADD COLUMN BankPaymentRef Text(255);"
        Conn1.Execute "Update Client set UsePayableTemplate =0;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_PayableTypes_PrintTemplate:
        On Error GoTo MOD_PayableTypes_PrintTemplate
        Rst1.Open "Select PrintTemplate from PayableTypes;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo MODIFY_ClientPROAGR_REVIEW_DATE
        Exit Function
MOD_PayableTypes_PrintTemplate:
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN PrintTemplate Text(255);"
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN EmailTemplate Text(255);"
        Conn1.Execute "ALTER TABLE PayableTypes ADD COLUMN SortCode Text(255);"
         UpdateDatabase2 = 1
         Exit Function
MODIFY_ClientPROAGR_REVIEW_DATE:
      Rst1.Open "SELECT REVIEW_DATE FROM ClientPROAGR;", Conn1, adOpenStatic, adLockReadOnly
     If Rst1.Fields(0).Type = 135 Then '202 means type text and 135 means datetime

         If Rst1.State = 1 Then
            Rst1.Close
         End If
            'Beacuse this fild user optional and can be null values
                Conn1.Execute "ALTER TABLE ClientPROAGR ALTER COLUMN REVIEW_DATE Text(255);"
                Conn1.Execute "ALTER TABLE ClientPROAGR ALTER COLUMN agreementStartDate Text(255);"
                Conn1.Execute "ALTER TABLE ClientPROAGR ALTER COLUMN agreementEndDate Text(255);"
                 UpdateDatabase2 = 1
       Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
       End If
    'No need after this because this is the last section
    
    
ADD_RptTransactions_VATAMOUNT:
        On Error GoTo MOD_RptTransactions_VATAMOUNT
        Rst1.Open "Select VATAMOUNT from RptTransactions;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
'        Exit Function
        GoTo ADD_GlobalData_TaxBasis
MOD_RptTransactions_VATAMOUNT:
    'RptTransactions
    Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN VatAmount Currency;"
    Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN NetAmount Currency;"
    Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN VAT_PERIOD_END_DATE DATE"
    Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN DeleteFlag BIT;"
    
    'PayTransactions
    Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN VatAmount Text(255);"
    Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN NetAmount Currency;"
    Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN VAT_PERIOD_END_DATE DATE"
    Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN DeleteFlag BIT;"
    
    'Global DATA
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN chkProduceVatReturn BIT"
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN TaxBasis DATE"
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN LastCompletedTaxReturnDate Date"
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN TaxInterval INT"
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN CurrentTaxPeriod Date"
    Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN isAgentToSubmit BIT"
    UpdateDatabase2 = 1
    Exit Function
        
ADD_GlobalData_TaxBasis:
        'On Error GoTo MOD_RptTransactions_VATAMOUNT
        Rst1.Open "Select TaxBasis from GlobalData;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
       If Rst1.Fields(0).Type = 135 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE GlobalData ALTER COLUMN TaxBasis Text(255)"
                Conn1.Execute "Update PayTransactions set DeleteFlag =false;"
                Conn1.Execute "Update RptTransactions set DeleteFlag =false;"
                'Because this fild user optional and can be null values
                 UpdateDatabase2 = 1
                 Exit Function
       Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
       End If
ADD_RptTransactions_SplitIDofSI:
        On Error GoTo MOD_RptTransactions_SplitIDofSI
        Rst1.Open "Select SplitIDofSI from RptTransactions;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_PayTransactions_SplitIDofPI
        Exit Function
MOD_RptTransactions_SplitIDofSI:
        Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN SplitIDofSI INT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_PayTransactions_SplitIDofPI:
        On Error GoTo MOD_PayTransactions_SplitIDofPI
        Rst1.Open "Select SplitIDofPI from PayTransactions;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
        GoTo ADD_PayTransactions_VATAMOUNT
        Exit Function
MOD_PayTransactions_SplitIDofPI:
        Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN SplitIDofPI INT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_PayTransactions_VATAMOUNT:
        'On Error GoTo MOD_RptTransactions_VATAMOUNT
        Rst1.Open "Select VATAMOUNT from PayTransactions;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
       If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE PayTransactions ALTER COLUMN VATAMOUNT Currency"
                Conn1.Execute "Update GlobalData set isAgentToSubmit =true;"
                  
                UpdateDatabase2 = 1
                Exit Function
       Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
       End If
MODIFTY_RentAnalysis_RRDemandType:
        'On Error GoTo MOD_RptTransactions_VATAMOUNT
        Rst1.Open "Select RRDemandType from RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly
        Rst1.Close
       If Rst1.Fields(0).DefinedSize = 1 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE RentAnalysis ALTER COLUMN RRDemandType INT"
                UpdateDatabase2 = 1
                Exit Function
       Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
       End If
ADD_RentSummaryStatement_BankACBalance1:
        On Error GoTo MOD_RentSummaryStatement_BankACBalance1
        Rst1.Open "Select  BankACBalance from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbPaymentsplit_PayTransactionIDSplit
        Exit Function
MOD_RentSummaryStatement_BankACBalance1:
        Conn1.Execute "ALTER TABLE RentSummaryStatement add BankACBalance Double;"
        UpdateDatabase2 = 1
        Exit Function
        'PayTransactionID
        'added this field by anol 2021-10-20 to support client statement
ADD_tlbPaymentsplit_PayTransactionIDSplit:
        On Error GoTo MOD_tlbPaymentsplit_PayTransactionIDSplit
        Rst1.Open "Select  PayTransactionIDSplit from tlbPaymentsplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbPaymentsplit_PropertyID
        Exit Function
MOD_tlbPaymentsplit_PayTransactionIDSplit:
        Conn1.Execute "ALTER TABLE tlbPaymentsplit add PayTransactionIDSplit long;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbPaymentsplit_PropertyID:
        On Error GoTo MOD_tlbPaymentsplit_PropertyID
        Rst1.Open "Select  PropertyID from tlbPaymentsplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo CreateTable1_PayTransactionsSplit
        'Exit Function
        'added this field by anol 2021-10-20 to support client statement
MOD_tlbPaymentsplit_PropertyID:
        Conn1.Execute "ALTER TABLE tlbPaymentsplit add PropertyID Text(10);"
        UpdateDatabase2 = 1
        Exit Function
        
        
CreateTable1_PayTransactionsSplit:
   On Error GoTo CreateTable_PayTransactionsSplit

   Rst1.Open "SELECT * FROM PayTransactionsSplit;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
   GoTo CreateTable1_RptTransactionsSplit
        
CreateTable_PayTransactionsSplit:
   Conn1.Execute _
      "CREATE TABLE PayTransactionsSplit" & _
         "(" & _
                "TransactionID LONG, " & _
                "TranType Text(5), " & _
                "Alloc_Unalloc Number, " & _
                "FromTran LONG, " & _
                "ToTran Long, " & _
                "Allocdate Date, " & _
                "PaymentAmount currency, " & _
                "BankCode Text(4), " & _
                "NominalCode Text(15), " & _
                "FundID LONG, " & _
                "NetAmount currency, " & _
                "VATAMOUNT currency, " & _
                "VAT_PERIOD_END_DATE Date, " & _
                "SplitIDofPI INT,  " & _
                "DeleteFlag BIT " & _
 ");"
        UpdateDatabase2 = 1
        Exit Function
 ' iam fully using DeleteFlag in PayTransaction and PayTransactionsSplit table, they are active
CreateTable1_RptTransactionsSplit:
   On Error GoTo CreateTable_RptTransactionsSplit

   Rst1.Open "SELECT * FROM RptTransactionsSplit;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close
    'Exit Function
   GoTo ADD_tlbReceiptsplit_RptTransactionsSplit
CreateTable_RptTransactionsSplit:
   Conn1.Execute _
      "CREATE TABLE RptTransactionsSplit" & _
         "(" & _
                "TransactionID LONG, " & _
                "FromTran LONG, " & _
                "ToTran Long, " & _
                "Allocdate Date, " & _
                "ReceiptAmount currency, " & _
                "BankCode Text(4), " & _
                "NominalCode Text(15), " & _
                "FundID LONG, " & _
                "NetAmount currency, " & _
                "VATAMOUNT currency, " & _
                "VAT_PERIOD_END_DATE Date, " & _
                "SplitIDofSI INT,  " & _
                "DeleteFlag BIT " & _
 ");"
 UpdateDatabase2 = 1
 Exit Function
 'added this field by anol 2021-10-20 to support client statement
ADD_tlbReceiptsplit_RptTransactionsSplit:
        On Error GoTo MOD_tlbReceiptsplit_RptTransactionsSplit
        Rst1.Open "Select RptTransactionsIDSplit from tlbReceiptSplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbReceiptSplit_PropertyID
        Exit Function
MOD_tlbReceiptsplit_RptTransactionsSplit:
        Conn1.Execute "ALTER TABLE tlbReceiptSplit add RptTransactionsIDSplit long;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbReceiptSplit_PropertyID:
        On Error GoTo MOD_tlbReceiptSplit_PropertyID
        Rst1.Open "Select  PropertyID from tlbReceiptSplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandRecords_ReportNetAmount
        'Exit Function
        'added this field by anol 2021-10-20 to support client statement
MOD_tlbReceiptSplit_PropertyID:
        Conn1.Execute "ALTER TABLE tlbReceiptSplit add PropertyID Text(10);"
        UpdateDatabase2 = 1
        Exit Function
ADD_DemandRecords_ReportNetAmount:
        On Error GoTo MOD_DemandRecords_ReportNetAmount
        Rst1.Open "Select  ReportNetAmount from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_ReportPaymentAmount
        'Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_DemandRecords_ReportNetAmount:
        Conn1.Execute "ALTER TABLE DemandRecords add ReportNetAmount Currency;"
        Conn1.Execute "ALTER TABLE DemandRecords add ReportVATAmount Currency;"
        Conn1.Execute "ALTER TABLE DemandRecords add ReportReceivedAmount Currency;"
        Conn1.Execute "ALTER TABLE DemandRecords add ReportDateFrom Date;"
        Conn1.Execute "ALTER TABLE DemandRecords add ReportDateTo Date;"
        UpdateDatabase2 = 1
        Exit Function
ADD_tblPurInv_ReportPaymentAmount:
        On Error GoTo MOD_tblPurInv_ReportPaymentAmount
        Rst1.Open "Select  ReportPaymentAmount from tblPurInv;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RentSummaryStatement_ConsolidatedPropID
        'Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_tblPurInv_ReportPaymentAmount:
        Conn1.Execute "ALTER TABLE tblPurInv add ReportInvNetAmount Currency;"
        Conn1.Execute "ALTER TABLE tblPurInv add ReportINVVATAmount Currency;"
        
        Conn1.Execute "ALTER TABLE tblPurInv add ReportPaymentAmount Currency;"
        Conn1.Execute "ALTER TABLE tblPurInv add ReportPayDescription Text(255);"
        
        Conn1.Execute "ALTER TABLE tblPurInv add ReportNominalCode Text(255);"
        Conn1.Execute "ALTER TABLE tblPurInv add ReportNominalName Text(255);"
        
        UpdateDatabase2 = 1
        Exit Function
        
ADD_RentSummaryStatement_ConsolidatedPropID:
        On Error GoTo MOD_RentSummaryStatement_ConsolidatedPropID
        Rst1.Open "Select  ConsolidatedPropID from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandRecords_ReportDemandTypeDesc
        Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_RentSummaryStatement_ConsolidatedPropID:
        Conn1.Execute "ALTER TABLE RentSummaryStatement add ConsolidatedPropID Text(10);"
        Conn1.Execute "ALTER TABLE RentSummaryStatement add StatementNobyProperty Long;"
        Conn1.Execute "ALTER TABLE RentSummaryStatement add ReportGenID Long;"
        Conn1.Execute "ALTER TABLE RentSummaryStatement add isConsolidated BIT;"
        
        UpdateDatabase2 = 1
        Exit Function
ADD_DemandRecords_ReportDemandTypeDesc:
        On Error GoTo MOD_DemandRecords_ReportDemandTypeDesc
        Rst1.Open "Select  ReportDemandTypeDesc from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RentSummaryStatement_isfinalized
        Exit Function
        'added this field by anol 2021-11-02 to support client statement
MOD_DemandRecords_ReportDemandTypeDesc:
        Conn1.Execute "ALTER TABLE DemandRecords add ReportDemandTypeDesc Text(255);"
        UpdateDatabase2 = 1
        Exit Function
ADD_RentSummaryStatement_isfinalized:
        On Error GoTo MOD_RentSummaryStatement_isfinalized
        Rst1.Open "Select  isfinalized from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_ReportFromDate
        Exit Function
        'added this field by anol 2021-11-02 to support client statement
MOD_RentSummaryStatement_isfinalized:
        Conn1.Execute "ALTER TABLE RentSummaryStatement add isfinalized int;"
        UpdateDatabase2 = 1
        Exit Function
        
        
ADD_tblPurInv_ReportFromDate:
        On Error GoTo MOD_tblPurInv_ReportFromDate
        Rst1.Open "Select  ReportFromDate from tblPurInv;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_ReportFromDatePreview
        Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_tblPurInv_ReportFromDate:
        Conn1.Execute "ALTER TABLE tblPurInv add ReportFromDate Date;"
        Conn1.Execute "ALTER TABLE tblPurInv add ReportToDate Date;"
        
        UpdateDatabase2 = 1
        Exit Function
ADD_tblPurInv_ReportFromDatePreview:
        On Error GoTo MOD_tblPurInv_ReportFromDatePreview
        Rst1.Open "Select  ReportFromDatePreview from tblPurInvPreview;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tblPurInv_isRentPayable
        Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_tblPurInv_ReportFromDatePreview:
        Conn1.Execute "ALTER TABLE tblPurInvPreview add ReportFromDatePreview Date;"
        Conn1.Execute "ALTER TABLE tblPurInvPreview add ReportToDatePreview Date;"
        
        UpdateDatabase2 = 1
        Exit Function
ADD_tblPurInv_isRentPayable:
        On Error GoTo MOD_tblPurInv_isRentPayable
        Rst1.Open "Select  isRentPayable from tblPurInv;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_TENANTDEPOSIT_SageAccountNumber
        'Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_tblPurInv_isRentPayable:
        Conn1.Execute "ALTER TABLE tblPurInv add isRentPayable BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_TENANTDEPOSIT_SageAccountNumber:
        'On Error GoTo MOD_TENANTDEPOSIT_SageAccountNumber
        Rst1.Open "Select  TenantID from TENANTDEPOSIT;", Conn1, adOpenKeyset, adLockReadOnly
        'Rst1.Close
        'GoTo ADD_RentSummaryStatement_ConsolidatedPropID
        If Rst1.Fields.Item("TenantID").DefinedSize < 30 Then
             Rst1.Close
             Set Rst1 = Nothing
             Conn1.Execute "ALTER TABLE TENANTDEPOSIT ALTER COLUMN  TenantID Text(30)"
             UpdateDatabase2 = 1
             Exit Function
          Else
             Rst1.Close
             Set Rst1 = Nothing
          End If
ADD_tlbClientBanks_LASTBACSFDATE:
        On Error GoTo MOD_tlbClientBanks_LASTBACSFDATE
        Rst1.Open "Select  LASTBACSFDATE from tlbClientBanks;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandRecords_ReportCreditAmount
        'Exit Function
        'added this field by anol 2021-12-06to support client statement
MOD_tlbClientBanks_LASTBACSFDATE:
        Conn1.Execute "ALTER TABLE tlbClientBanks add LASTBACSFDATE Date;"
        Conn1.Execute "ALTER TABLE tlbClientBanks add LASTBACSFNO INT;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_DemandRecords_ReportCreditAmount:
        On Error GoTo MOD_DemandRecords_ReportCreditAmount
        Rst1.Open "Select  ReportCreditAmount from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandRecords_ReportOSAmount
        'Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_DemandRecords_ReportCreditAmount:
        Conn1.Execute "ALTER TABLE DemandRecords add ReportCreditAmount Currency;"
        UpdateDatabase2 = 1
         Exit Function
        'ReportOSAmount
ADD_DemandRecords_ReportOSAmount:
        On Error GoTo MOD_DemandRecords_ReportOSAmount
        Rst1.Open "Select  ReportOSAmount from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandSplitRecords_ReportNetAmountS
        Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_DemandRecords_ReportOSAmount:
        Conn1.Execute "ALTER TABLE DemandRecords add ReportOSAmount Currency;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_DemandSplitRecords_ReportNetAmountS:
        On Error GoTo MOD_DemandSplitRecords_ReportNetAmountS
        Rst1.Open "Select  ReportNetAmountS from DemandSplitRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandSplitRecords_ReportCsShowFlag
        'Exit Function
        'added this field by anol 2021-10-28 to support client statement
MOD_DemandSplitRecords_ReportNetAmountS:
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportNetAmountS Currency;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportVATAmountS Currency;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportReceivedAmountS Currency;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportCreditAmountS Currency;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add reportOSAmountS Currency;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportDateFromS Date;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportDateToS Date;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportDemandTypeDescS Text(255);"
        UpdateDatabase2 = 1
        Exit Function
ADD_DemandSplitRecords_ReportCsShowFlag:
        On Error GoTo MOD_DemandSplitRecords_ReportCsShowFlag
        Rst1.Open "Select  ReportCsShowFlag from DemandSplitRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_TenantDeposit_ParentRFID
        Exit Function
        
MOD_DemandSplitRecords_ReportCsShowFlag:
        Conn1.Execute "ALTER TABLE DemandSplitRecords add ReportCsShowFlag Text(1);"
        UpdateDatabase2 = 1
         Exit Function
ADD_TenantDeposit_ParentRFID:
        On Error GoTo MOD_TenantDeposit_ParentRFID
        Rst1.Open "Select  ParentRFID from TenantDeposit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Paytransaction_ClientID
        Exit Function
        
MOD_TenantDeposit_ParentRFID:
        Conn1.Execute "ALTER TABLE TenantDeposit add ParentRFID Text(50);"
        UpdateDatabase2 = 1
         Exit Function
         
ADD_Paytransaction_ClientID:
        On Error GoTo MOD_Paytransaction_ClientID
        Rst1.Open "Select  ClientID from Paytransactions;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbReceiptsplit_ISMGTFeeS
        Exit Function
        
MOD_Paytransaction_ClientID:
        Conn1.Execute "ALTER TABLE Paytransactions add ClientID Text(10);"
        Conn1.Execute "ALTER TABLE Paytransactions add PIOrPPR Text(100);"
        Conn1.Execute "ALTER TABLE Paytransactions add PPOrPAOrPC Text(100);"
        Conn1.Execute "ALTER TABLE Paytransactions add PaymentSAGEAC Text(100);"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbReceiptsplit_ISMGTFeeS:
        On Error GoTo MOD_tlbReceiptsplit_ISMGTFeeS
        Rst1.Open "Select ISMGTFeeS from tlbReceiptsplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbAgreement_ReportClientID
        Exit Function
        
MOD_tlbReceiptsplit_ISMGTFeeS:
        Conn1.Execute "ALTER TABLE tlbReceiptsplit ADD COLUMN ISMGTFeeS  BIT;" 'by default they are false
        Conn1.Execute "Update tlbReceiptsplit RC,tlbReceipt R Set  ISMGTFeeS=ISMGTFee where R.TransactionID=RC.rptHeader ;"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbAgreement_ReportClientID:
        On Error GoTo MOD_tlbAgreement_ReportClientID
        Rst1.Open "Select ReportClientID from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_DemandSplitRecords_ReportClientID
        Exit Function
        
MOD_tlbAgreement_ReportClientID:
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN ReportClientID  text(10);"
        Conn1.Execute "ALTER TABLE tlbAgreement ADD COLUMN ReportPropertyID  text(10);"
        UpdateDatabase2 = 1
        Exit Function
ADD_DemandSplitRecords_ReportClientID:
        On Error GoTo MOD_DemandSplitRecords_ReportClientID
        Rst1.Open "Select ReportClientID from DemandSplitRecords;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbReceiptSplit_ClientStatementID
        Exit Function
        
MOD_DemandSplitRecords_ReportClientID:
        Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ReportClientID  text(10);"
        Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ReportPropertyID  text(10);"
        Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ReportSAGEAC  text(30);"
        Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ReportTransactionID  LONG;"
        Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ReportTransactionNo  text(30);"
        UpdateDatabase2 = 1
        Exit Function
ADD_tlbReceiptSplit_ClientStatementID:
        On Error GoTo MOD_tlbReceiptSplit_ClientStatementID
        Rst1.Open "Select ClientStatementID from tlbReceiptSplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Table_RentSummaryStatementdetails
        Exit Function
        
MOD_tlbReceiptSplit_ClientStatementID:
        'ClientStatementPrevID ClientStatementID
        Conn1.Execute "ALTER TABLE tlbReceiptSplit ADD COLUMN ClientStatementID  LONG;"
        Conn1.Execute "ALTER TABLE tlbReceiptSplit ADD COLUMN ClientStatementPrevID  LONG;"
        Conn1.Execute "ALTER TABLE tlbPaymentSplit ADD COLUMN ClientStatementID  LONG;"
        Conn1.Execute "ALTER TABLE tlbPaymentSplit ADD COLUMN ClientStatementPrevID  LONG;"
        Conn1.Execute "Update tlbReceiptSplit S,tlbReceipt R set ClientStatementID=RentSumStatement where R.TransactionID=S.rptHeader ;"
        Conn1.Execute "Update tlbPaymentSplit S,tlbPayment P set ClientStatementID=RentSumStatement where P.TransactionID=S.payHeader ;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_Table_RentSummaryStatementdetails:
        On Error GoTo MOD_Table_RentSummaryStatementdetails
        Rst1.Open "Select * from RentSummaryStatementdetails;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RetentionDetails_ClientID
        Exit Function
        
MOD_Table_RentSummaryStatementdetails:

      Conn1.Execute "CREATE TABLE RentSummaryStatementdetails " & _
         "(" & _
            "StatementID      LONG NOT NULL , " & _
            "PINumber      TEXT(100) NOT NULL, " & _
            "Amount    Double, " & _
            "OSamount        Double " & _
         ");"

        UpdateDatabase2 = 1
        Exit Function
ADD_RetentionDetails_ClientID:
        On Error GoTo MOD_RetentionDetails_ClientID
        Rst1.Open "Select ClientID from RetentionDetails;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RentSummaryStatementdetails_ClientID
        Exit Function
        
MOD_RetentionDetails_ClientID:
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN ClientID  Text(10);"
        UpdateDatabase2 = 1
        Exit Function
ADD_RentSummaryStatementdetails_ClientID:
        On Error GoTo MOD_RentSummaryStatementdetails_ClientID
        Rst1.Open "Select ClientID from RentSummaryStatementdetails;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RetentionDetailsPrev
        Exit Function
        
MOD_RentSummaryStatementdetails_ClientID:
        Conn1.Execute "ALTER TABLE RentSummaryStatementdetails ADD COLUMN ClientID  Text(10);"
        Conn1.Execute "ALTER TABLE RentSummaryStatementdetails ADD COLUMN SageAccountNumber  Text(30);"
        Conn1.Execute "ALTER TABLE RentSummaryStatementdetails ADD COLUMN ID  LONG;"
        Conn1.Execute "ALTER TABLE RentSummaryStatementdetails ADD COLUMN SplitID INT;"
        Conn1.Execute "ALTER TABLE RentSummaryStatementdetails ADD COLUMN PercentageLL Double;"
        
        UpdateDatabase2 = 1
        Exit Function
         
ADD_RetentionDetailsPrev:
        On Error GoTo MOD_RetentionDetailsPrev
        Rst1.Open "Select * from RetentionDetailsPrev;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_ReportManagingAgentFeesPaid
        Exit Function

MOD_RetentionDetailsPrev:
         Conn1.Execute "Create TABLE RetentionDetailsPrev " & _
        "(" & _
               "StatementID Number, SlNumber number,Description Text(250),amount Double,isCleared BIT,ClientID TEXT(10)" & _
        ");"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_ReportManagingAgentFeesPaid:
        On Error GoTo MOD_ReportManagingAgentFeesPaid
        Rst1.Open "Select * from ReportManagingAgentFeesPaid;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_ClientID_LesseeTemplate
        Exit Function

MOD_ReportManagingAgentFeesPaid:
         Conn1.Execute "Create TABLE ReportManagingAgentFeesPaid " & _
        "(" & _
               "ClientID Text(10),PropertyID Text(10),PI Text(25),TransactionID LONG, InvNo Text(10), TransDate Date ,Reference Text(250),Description Text(250),PaidAmount Double,OsAmount Double,NET Double,VAT Double,Total Double" & _
        ");"
        UpdateDatabase2 = 1
        Exit Function
ADD_ClientID_LesseeTemplate:
        On Error GoTo MOD_ClientID_LesseeTemplate
        Rst1.Open "Select LesseeTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo Modify_tlbPayment_Details
        Exit Function

MOD_ClientID_LesseeTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN LesseeTemplate text(250);"
        UpdateDatabase2 = 1
        Exit Function
        
Modify_tlbPayment_Details:             'this field is needed for Print payment list report in PI form
     'On Error GoTo Mod_Landlord_LandlordID ' this is no fixed by occuring error
     Rst1.Open "SELECT Details FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly
     Rst1.Close
     If Rst1.Fields.Item("Details").DefinedSize = 80 Then  '202 means type text and 135 means datetime
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo Mod_tlbPayment_Details
     Else
         If Rst1.State = 1 Then
            Rst1.Close
         End If
         GoTo ADD_RetentionDetails_ReferenceNo 'go t next label
     End If
     Exit Function

Mod_tlbPayment_Details:
    Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN Details Text(255);"
    UpdateDatabase2 = 1
    Exit Function
    
    
ADD_RetentionDetails_ReferenceNo:
        On Error GoTo MOD_RetentionDetails_ReferenceNo
        Rst1.Open "Select Reference from RetentionDetails;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_TenantDeposit_DepositSL
        Exit Function

MOD_RetentionDetails_ReferenceNo:
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN Reference Text(255);"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN RDate Date;"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN PropertyID text(10);"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN isHist BIT;"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN FUNDID LONG;"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN ID LONG;"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN isDeleted BIT;"
        Conn1.Execute "ALTER TABLE RetentionDetails ADD COLUMN isPrint BIT;"
        'We are not using RetentionDetailsPrev table

        UpdateDatabase2 = 1
        Exit Function
ADD_TenantDeposit_DepositSL:
        On Error GoTo MOD_TenantDeposit_DepositSL
        Rst1.Open "Select DepositSL from TenantDeposit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_GlobalSC_SCYearEndGenerated
        Exit Function

MOD_TenantDeposit_DepositSL:
        Conn1.Execute "alter table TenantDeposit ADD COLUMN DepositTypePrefix Text(10);"
        Conn1.Execute "alter table TenantDeposit ADD COLUMN DepositSL LONG;"
        Rst1.Open "Select * from TenantDeposit where left(transactionID,1)='D';", Conn1, adOpenKeyset, adLockOptimistic
        iCount1 = 1 'For type D
        While Not Rst1.EOF
            Rst1!DepositTypePrefix = "DP"
            Rst1!DepositSL = iCount1
            iCount1 = iCount1 + 1
            Rst1.Update
            Rst1.MoveNext
        Wend
        Rst1.Close
        
        Rst1.Open "Select * from TenantDeposit where left(transactionID,1)='R';", Conn1, adOpenKeyset, adLockOptimistic
        iCount1 = 1 'For type R
        While Not Rst1.EOF
            Rst1!DepositTypePrefix = "DR"
            Rst1!DepositSL = iCount1
            iCount1 = iCount1 + 1
            Rst1.Update
            Rst1.MoveNext
        Wend
        Rst1.Close
        
        Rst1.Open "Select * from TenantDeposit where left(transactionID,1)='RF';", Conn1, adOpenKeyset, adLockOptimistic
        iCount1 = 1 'For type RF
        While Not Rst1.EOF
            Rst1!DepositTypePrefix = "FR"
            Rst1!DepositSL = iCount1
            iCount1 = iCount1 + 1
            Rst1.Update
            Rst1.MoveNext
        Wend
        Rst1.Close
        
        Rst1.Open "Select * from TenantDeposit where left(transactionID,1)='E';", Conn1, adOpenKeyset, adLockOptimistic
        iCount1 = 1 'For type E
        While Not Rst1.EOF
            Rst1!DepositTypePrefix = "EX"
            Rst1!DepositSL = iCount1
            iCount1 = iCount1 + 1
            Rst1.Update
            Rst1.MoveNext
        Wend
        Rst1.Close
        
        UpdateDatabase2 = 1
        Exit Function
ADD_GlobalSC_SCYearEndGenerated:
        On Error GoTo MOD_GlobalSC_SCYearEndGenerated
        Rst1.Open "Select SCYearEndGenerated from GlobalSC;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_RentSummaryStatement_DateFinalized
        Exit Function

MOD_GlobalSC_SCYearEndGenerated:
        Conn1.Execute "ALTER TABLE GlobalSC ADD COLUMN SCYearEndGenerated BIT;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_RentSummaryStatement_DateFinalized:
        On Error GoTo MOD_RentSummaryStatement_DateFinalized
        Rst1.Open "Select DateFinalized from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_Client_ShowBankAccountFunds
        Exit Function

MOD_RentSummaryStatement_DateFinalized:
        Conn1.Execute "ALTER TABLE RentSummaryStatement ADD COLUMN DateFinalized DATE;"
        UpdateDatabase2 = 1
        Exit Function
'ADD_tlbClientBanks_FundId: FundId won't be used in tlbClientBanktable. I am removing it.Bank fund is one to many relationship.  instead new table will be use. BankFund
'        On Error GoTo MOD_tlbClientBanks_FundId
'        Rst1.Open "Select FundId from tlbClientBanks;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        'GoTo ADD_TenantDeposit_DepositSL
'        Exit Function
'
'MOD_tlbClientBanks_FundId:
'        Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN FundId INT;"
'        UpdateDatabase2 = 1
'        Exit Function

ADD_Client_ShowBankAccountFunds:
        On Error GoTo MOD_Client_ShowBankAccountFunds
        Rst1.Open "Select ShowBankAccountFunds from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_BankFund_Table
        Exit Function

MOD_Client_ShowBankAccountFunds:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN ShowBankAccountFunds BIT;"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_BankFund_Table:
        On Error GoTo MOD_BankFund_Table
        Rst1.Open "Select * from BankFund;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_table_ReportClientHistory
        Exit Function

MOD_BankFund_Table:
         Conn1.Execute "Create TABLE BankFund " & _
        "(" & _
               "ClientID Text(10),BankCode Text(15),FundID INT" & _
        ");"
        UpdateDatabase2 = 1
        Exit Function
        
'ADD_ReportBudVsAE_PropertyID:
'        On Error GoTo MOD_ReportBudVsAE_PropertyID
'        Rst1.Open "Select PropertyID from ReportBudVsAE;", Conn1, adOpenKeyset, adLockReadOnly
'        Rst1.Close
'        'GoTo ADD_BankFund_Table
'        Exit Function
'
'MOD_ReportBudVsAE_PropertyID:
'        Conn1.Execute "ALTER TABLE ReportBudVsAE ADD COLUMN PropertyID TEXT(10);"
'        Conn1.Execute "ALTER TABLE ReportBudVsAE ADD COLUMN FundID Long;"
'        Conn1.Execute "ALTER TABLE ReportBudVsAE ADD COLUMN FYrID TEXT(20);"
'        UpdateDatabase2 = 1
ADD_table_ReportClientHistory:
     On Error GoTo MOD_table_ReportClientHistory
     Rst1.Open "Select * from ReportClientHistory;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_Landlord_StClientHomeTel
     Exit Function
MOD_table_ReportClientHistory:
    Conn1.Execute "Create TABLE ReportClientHistory " & _
         "(" & _
            "ReportingDate DateTime  NOT NULL, " & _
            "SessionID     TEXT(100) NOT NULL, " & _
            "SIGN text(1), " & _
            "transactionID LONG, " & _
            "Type Number, " & _
            "Type_desc text(255), " & _
            "PF text(255), " & _
            "Pdate datetime, " & _
            "Details text(255), " & _
            "extref text(255), " & _
            "amount currency, " & _
            "Osamount  currency, " & _
            "Balance  currency, " & _
            "flag number, " & _
            "isMaster number, " & _
            "ActualINV TEXT(255), " & _
            "ClientID Text(10) " & _
            ");"
      UpdateDatabase2 = 1
      Exit Function
ADD_Landlord_StClientHomeTel:
        On Error GoTo MOD_Landlord_StClientHomeTel
        Rst1.Open "Select  StToLandlordAddress from Landlord;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_ReportClientHistory_SageAccountNumber
        Exit Function
MOD_Landlord_StClientHomeTel:
        
        Conn1.Execute "ALTER TABLE Supplier add StToLandlordAddress INT;"
        Conn1.Execute "Update Supplier set StToLandlordAddress =1;"
        Conn1.Execute "ALTER TABLE Supplier add StToStatementAddress INT;"
        Conn1.Execute "Update Supplier set StToStatementAddress =1;"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordHomeTel text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordHomeEmail text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordOfficeTel text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordStatementEmail text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordMobile text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordStatementHometel text(250);"
        Conn1.Execute "ALTER TABLE Supplier add StLandlordOfficeEmail text(250);"
        
        '*****update the same fields for landlord
        Conn1.Execute "ALTER TABLE Landlord add StToLandlordAddress INT;"
        Conn1.Execute "Update Landlord set StToLandlordAddress =1;"
        Conn1.Execute "ALTER TABLE Landlord add StToStatementAddress INT;"
        Conn1.Execute "Update Landlord set StToStatementAddress =1;"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordHomeTel text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordHomeEmail text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordOfficeTel text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordStatementEmail text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordMobile text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordStatementHometel text(250);"
        Conn1.Execute "ALTER TABLE Landlord add StLandlordOfficeEmail text(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_ReportClientHistory_SageAccountNumber:
        On Error GoTo MOD_ReportClientHistory_SageAccountNumber
        Rst1.Open "Select  SageAccountNumber from ReportClientHistory;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_tlbReceiptSplit_PIREFMGTFEES
        Exit Function
MOD_ReportClientHistory_SageAccountNumber:
        Conn1.Execute "ALTER TABLE ReportClientHistory add SageAccountNumber text(30);"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_tlbReceiptSplit_PIREFMGTFEES:
        On Error GoTo MOD_tlbReceiptSplit_PIREFMGTFEES
        Rst1.Open "Select  PIREFMGTFEES from tlbReceiptSplit;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_table_ReportClientStatementDemands
        Exit Function
MOD_tlbReceiptSplit_PIREFMGTFEES:
        'Conn1.Execute "ALTER TABLE tlbAgreement DROP COLUMN DEMAND_TYPE;"
        Conn1.Execute "ALTER TABLE tlbReceipt DROP COLUMN ISMGTFEE;"
        Conn1.Execute "ALTER TABLE tlbReceiptSplit add PIREFMGTFEES text(100);"
        Conn1.Execute "ALTER TABLE tlbReceiptSplit add ChargeDateS DateTime;"
        Conn1.Execute "Update tlbReceiptSplit S,tlbReceipt R set PIREFMGTFEES=PIREFMGTFEE,ChargeDateS=ChargeDate where R.TransactionID=S.rptHeader ;"
        Conn1.Execute "ALTER TABLE tlbReceipt DROP COLUMN PIREFMGTFEE;"
        Conn1.Execute "Create TABLE ManagementFee " & _
         "(" & _
            "PI_ActualID Text(25), " & _
            "SRSlNumber  Long, " & _
            "ReceiptSLNumber  Long, " & _
            "ReceiptType Number, " & _
            "ChargingMethod text(30), " & _
            "SageAccountNumber text(30), " & _
            "ReceiptTypeDescription text(60), " & _
            "PropertyID Text(10), " & _
            "FundID Number, " & _
            "RptAmtType Text(100), " & _
            "ExtRef Text(100), " & _
            "ChargeDate DateTime, " & _
            "ReceiptDate DateTime, " & _
            "ReceiptTransactionID Long, " & _
            "ReceiptSplitID Long, " & _
            "ReceiptAmount Currency, " & _
            "AgrPercentage  Currency, " & _
            "MgtFeeAmt  Currency, " & _
            "VATPercentage  Currency, " & _
            "VAT  Currency, " & _
            "MgtFeeAmtTotal Currency );"
         Conn1.Execute "Create TABLE ManagementFeePreview " & _
         "(" & _
            "PI_ActualID Text(25), " & _
            "SRSlNumber  Long, " & _
            "ReceiptSLNumber  Long, " & _
            "ReceiptType Number, " & _
            "ChargingMethod text(30), " & _
            "SageAccountNumber text(30), " & _
            "ReceiptTypeDescription text(60), " & _
            "PropertyID Text(10), " & _
            "FundID Number, " & _
            "RptAmtType Text(100), " & _
            "ExtRef Text(100), " & _
            "ChargeDate DateTime, " & _
            "ReceiptDate DateTime, " & _
            "ReceiptTransactionID Long, " & _
            "ReceiptSplitID Long, " & _
            "ReceiptAmount Currency, " & _
            "AgrPercentage  Currency, " & _
            "MgtFeeAmt  Currency, " & _
            "VATPercentage  Currency, " & _
            "VAT  Currency, " & _
            "MgtFeeAmtTotal Currency );"
            
        UpdateDatabase2 = 1
        Exit Function
ADD_table_ReportClientStatementDemands:
     On Error GoTo MOD_table_ReportClientStatementDemands
     Rst1.Open "Select * from ReportClientStatementDemands;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Add_DemandRecords_exclCRNtoCS
     Exit Function
MOD_table_ReportClientStatementDemands:
    Conn1.Execute "Create TABLE ReportClientStatementDemands " & _
         "(" & _
            "StatementID Long, " & _
            "SITrxID  Long, " & _
            "ClientID  TEXT(100), " & _
            "PropertyID     TEXT(10), " & _
            "DemandID      Long, " & _
            "DemandSLNumber      Long, " & _
            "SplitID      INT, " & _
            "SageAccountNumber  TEXT(30), " & _
            "UnitNumber      TEXT(25), " & _
            "TransactionType INT, " & _
            "TypeOfDemand     Long, " & _
            "DueDate     Date, " & _
            "DateFrom      Date, " & _
            "DateTo      Date, " & _
            "NetAmount  Currency, " & _
            "VATAmount    Currency, " & _
            "ReceivedAmountS  Currency, " & _
            "CreditAmount    Currency, " & _
            "OSAmount  Currency " & _
            ");"
       Conn1.Execute "Create TABLE ReportClientStatementDemandsPreview " & _
         "(" & _
            "StatementID Long, " & _
            "SITrxID  Long, " & _
            "ClientID  TEXT(100), " & _
            "PropertyID     TEXT(10), " & _
            "DemandID      Long, " & _
            "DemandSLNumber      Long, " & _
            "SplitID      INT, " & _
            "SageAccountNumber  TEXT(30), " & _
            "UnitNumber      TEXT(25), " & _
            "TransactionType INT, " & _
            "TypeOfDemand     Long, " & _
            "DueDate     Date, " & _
            "DateFrom      Date, " & _
            "DateTo      Date, " & _
            "NetAmount  Currency, " & _
            "VATAmount    Currency, " & _
            "ReceivedAmountS  Currency, " & _
            "CreditAmount    Currency, " & _
            "OSAmount  Currency " & _
            ");"
            Conn1.Execute "Create TABLE ReportClientStatementPurchases " & _
         "(" & _
            "StatementID Long, " & _
            "Type     INT, " & _
            "MY_ID      TEXT(25), " & _
            "SplitID      Long, " & _
            "TransactionID Long, " & _
            "ClientID  TEXT(100), " & _
            "PropertyID     TEXT(10), " & _
            "SupplierID     TEXT(250), " & _
            "TranDate  Date, " & _
            "NOMINAL_CODE      TEXT(10), " & _
            "PaymentDescription  TEXT(255), " & _
            "PaymentRef  TEXT(100), " & _
            "NetAmount  Currency, " & _
            "VATAmount    Currency, " & _
            "PaymentAmount  Currency, " & _
            "CreditAmount    Currency, " & _
            "OSAmount  Currency " & _
            ");"
            Conn1.Execute "Create TABLE ReportClientStatementPurchasesPreview " & _
         "(" & _
            "StatementID Long, " & _
            "Type     INT, " & _
            "MY_ID      TEXT(25), " & _
            "SplitID      Long, " & _
            "TransactionID Long, " & _
            "ClientID  TEXT(100), " & _
            "PropertyID     TEXT(10), " & _
            "SupplierID     TEXT(250), " & _
            "TranDate  Date, " & _
            "NOMINAL_CODE      TEXT(10), " & _
            "PaymentDescription  TEXT(255), " & _
            "PaymentRef  TEXT(100), " & _
            "NetAmount  Currency, " & _
            "VATAmount    Currency, " & _
            "PaymentAmount  Currency, " & _
            "CreditAmount    Currency, " & _
            "OSAmount  Currency " & _
            ");"
            
     UpdateDatabase2 = 1
     Exit Function
Add_DemandRecords_exclCRNtoCS:
    '2023-01-26 anol exclCRNtoCS field means you dont want to see some credit note in the CS Report
     On Error GoTo MOD_DemandRecords_exclCRNtoCS
     Rst1.Open "Select exclCRNtoCS from DemandRecords;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Modify_NLPosting_PropertyID
     Exit Function
MOD_DemandRecords_exclCRNtoCS:
     Conn1.Execute "ALTER TABLE DemandRecords add column exclCRNtoCS BIT"
     Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     UpdateDatabase2 = 1
     Exit Function
Modify_NLPosting_PropertyID:
        'On Error GoTo MOD_ReportClientHistory_SageAccountNumber
        Rst1.Open "Select  Property_ID from NLPosting;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item(0).DefinedSize < 10 Then
              Rst1.Close
              Set Rst1 = Nothing
                Conn1.Execute "Create TABLE NLPOSTINGTemp " & _
                "(" & _
                   "THIS_RECORD TEXT(50),  PARENT_RECORD      TEXT(50)," & _
                   "TRANS_ID       TEXT(50), UNIQUE_REFERENCE_NO  AUTOINCREMENT, " & _
                   "POSTED_DATE DateTime, TRANSACTION_DATE DateTime,  " & _
                   "TRANSACTION_TYPE     INT, " & _
                   "ACCOUNT_NUMBER     TEXT(30), " & _
                   "PROPERTY_ID  TEXT(10), " & _
                   "UNIT_ID      TEXT(25), " & _
                   "JOB_NO  TEXT(10), " & _
                   "FUND_ID  Long, " & _
                   "SCHEDULE_ID  Long, " & _
                   "AMOUNT    Currency, " & _
                   "REFERENCE  TEXT(100), " & _
                   "NOMINAL_CODE    TEXT(15), " & _
                   "TRANSACTION_DESCRIPTION    TEXT(255), " & _
                   "VAT_PERIOD_END_DT  Currency, " & _
                   "YEAR_END_CLOSED    BIT, " & _
                   "USER_NUMBER  TEXT(50), " & _
                   "VAT_RECON    BIT, " & _
                   "AMOUNT_TYPE   TEXT(1), ClientID  TEXT(10)," & _
                   "DeleteFlag BIT, TRANSACTION_REF TEXT(10) " & _
                   ");"
                   'Transfer Data
                   Dim szSQL As String
                   Dim szSQL2 As String
                   Dim adoDst As New ADODB.Recordset
                   Dim adoSrc As New ADODB.Recordset
                   szSQL = "SELECT * FROM NLPosting;"
                   adoSrc.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
                
                  szSQL2 = "SELECT * FROM NLPostingTemp;"
                  adoDst.Open szSQL2, Conn1, adOpenDynamic, adLockOptimistic
                
                   While Not adoSrc.EOF
                      With adoDst
                        .AddNew
                            .Fields.Item("THIS_RECORD").Value = adoSrc.Fields.Item("THIS_RECORD").Value
                            .Fields.Item("PARENT_RECORD").Value = adoSrc.Fields.Item("PARENT_RECORD").Value
                            .Fields.Item("TRANS_ID").Value = adoSrc.Fields.Item("TRANS_ID").Value
                           ' .Fields.Item("UNIQUE_REFERENCE_NO").Value = adoSrc.Fields.Item("UNIQUE_REFERENCE_NO").Value
                            .Fields.Item("POSTED_DATE").Value = Format(adoSrc.Fields.Item("POSTED_DATE").Value, "DD MMMM YYYY")
                            .Fields.Item("TRANSACTION_DATE").Value = Format(adoSrc.Fields.Item("TRANSACTION_DATE").Value, "DD MMMM YYYY")
                            .Fields.Item("TRANSACTION_TYPE").Value = adoSrc.Fields.Item("TRANSACTION_TYPE").Value
                            .Fields.Item("ACCOUNT_NUMBER").Value = adoSrc.Fields.Item("ACCOUNT_NUMBER").Value
                            .Fields.Item("PROPERTY_ID").Value = adoSrc.Fields.Item("PROPERTY_ID").Value
                            .Fields.Item("UNIT_ID").Value = adoSrc.Fields.Item("UNIT_ID").Value
                            .Fields.Item("JOB_NO").Value = adoSrc.Fields.Item("JOB_NO").Value
                            .Fields.Item("FUND_ID").Value = adoSrc.Fields.Item("FUND_ID").Value
                            .Fields.Item("SCHEDULE_ID").Value = adoSrc.Fields.Item("SCHEDULE_ID").Value
                            .Fields.Item("Amount").Value = adoSrc.Fields.Item("Amount").Value
                            .Fields.Item("REFERENCE").Value = adoSrc.Fields.Item("REFERENCE").Value
                            .Fields.Item("NOMINAL_CODE").Value = adoSrc.Fields.Item("NOMINAL_CODE").Value
                            .Fields.Item("TRANSACTION_DESCRIPTION").Value = adoSrc.Fields.Item("TRANSACTION_DESCRIPTION").Value
                            .Fields.Item("VAT_PERIOD_END_DT").Value = adoSrc.Fields.Item("VAT_PERIOD_END_DT").Value
                            .Fields.Item("YEAR_END_CLOSED").Value = adoSrc.Fields.Item("YEAR_END_CLOSED").Value
                            .Fields.Item("USER_NUMBER").Value = adoSrc.Fields.Item("USER_NUMBER").Value
                            .Fields.Item("VAT_RECON").Value = adoSrc.Fields.Item("VAT_RECON").Value
                            .Fields.Item("AMOUNT_TYPE").Value = adoSrc.Fields.Item("AMOUNT_TYPE").Value
                            .Fields.Item("ClientID").Value = adoSrc.Fields.Item("ClientID").Value
                            .Fields.Item("DeleteFlag").Value = adoSrc.Fields.Item("DeleteFlag").Value
                            .Fields.Item("TRANSACTION_REF").Value = adoSrc.Fields.Item("TRANSACTION_REF").Value
                        .Update
                         adoSrc.MoveNext
                     End With
                Wend
                adoSrc.Close
                adoDst.Close
                    
                   Conn1.Execute "Drop TABLE NLPOSTING "
                   Conn1.Execute "Create TABLE NLPOSTING " & _
                "(" & _
                   "THIS_RECORD TEXT(50),  PARENT_RECORD      TEXT(50)," & _
                   "TRANS_ID       TEXT(50), UNIQUE_REFERENCE_NO  AUTOINCREMENT, " & _
                   "POSTED_DATE DateTime, TRANSACTION_DATE DateTime,  " & _
                   "TRANSACTION_TYPE     INT, " & _
                   "ACCOUNT_NUMBER     TEXT(30), " & _
                   "PROPERTY_ID  TEXT(10), " & _
                   "UNIT_ID      TEXT(25), " & _
                   "JOB_NO  TEXT(10), " & _
                   "FUND_ID  Long, " & _
                   "SCHEDULE_ID  Long, " & _
                   "AMOUNT    Currency, " & _
                   "REFERENCE  TEXT(100), " & _
                   "NOMINAL_CODE    TEXT(15), " & _
                   "TRANSACTION_DESCRIPTION    TEXT(255), " & _
                   "VAT_PERIOD_END_DT  Currency, " & _
                   "YEAR_END_CLOSED    BIT, " & _
                   "USER_NUMBER  TEXT(50), " & _
                   "VAT_RECON    BIT, " & _
                   "AMOUNT_TYPE   TEXT(1), ClientID  TEXT(10)," & _
                   "DeleteFlag BIT, TRANSACTION_REF TEXT(10) " & _
                   ");"
                   'Transfer Data
'                   Dim szSQL As String
'                   Dim szSQL2 As String
'                   Dim adoDst As New ADODB.Recordset
'                   Dim adoSrc As New ADODB.Recordset
                   szSQL = "SELECT * FROM NLPostingTemp;"
                   adoSrc.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
                
                  szSQL2 = "SELECT * FROM NLPosting;"
                  adoDst.Open szSQL2, Conn1, adOpenDynamic, adLockOptimistic
                
                   While Not adoSrc.EOF
                      With adoDst
                        .AddNew
                            .Fields.Item("THIS_RECORD").Value = adoSrc.Fields.Item("THIS_RECORD").Value
                            .Fields.Item("PARENT_RECORD").Value = adoSrc.Fields.Item("PARENT_RECORD").Value
                            .Fields.Item("TRANS_ID").Value = adoSrc.Fields.Item("TRANS_ID").Value
                           ' .Fields.Item("UNIQUE_REFERENCE_NO").Value = adoSrc.Fields.Item("UNIQUE_REFERENCE_NO").Value
                            .Fields.Item("POSTED_DATE").Value = Format(adoSrc.Fields.Item("POSTED_DATE").Value, "DD MMMM YYYY")
                            .Fields.Item("TRANSACTION_DATE").Value = Format(adoSrc.Fields.Item("TRANSACTION_DATE").Value, "DD MMMM YYYY")
                            .Fields.Item("TRANSACTION_TYPE").Value = adoSrc.Fields.Item("TRANSACTION_TYPE").Value
                            .Fields.Item("ACCOUNT_NUMBER").Value = adoSrc.Fields.Item("ACCOUNT_NUMBER").Value
                            .Fields.Item("PROPERTY_ID").Value = adoSrc.Fields.Item("PROPERTY_ID").Value
                            .Fields.Item("UNIT_ID").Value = adoSrc.Fields.Item("UNIT_ID").Value
                            .Fields.Item("JOB_NO").Value = adoSrc.Fields.Item("JOB_NO").Value
                            .Fields.Item("FUND_ID").Value = adoSrc.Fields.Item("FUND_ID").Value
                            .Fields.Item("SCHEDULE_ID").Value = adoSrc.Fields.Item("SCHEDULE_ID").Value
                            .Fields.Item("Amount").Value = adoSrc.Fields.Item("Amount").Value
                            .Fields.Item("REFERENCE").Value = adoSrc.Fields.Item("REFERENCE").Value
                            .Fields.Item("NOMINAL_CODE").Value = adoSrc.Fields.Item("NOMINAL_CODE").Value
                            .Fields.Item("TRANSACTION_DESCRIPTION").Value = adoSrc.Fields.Item("TRANSACTION_DESCRIPTION").Value
                            .Fields.Item("VAT_PERIOD_END_DT").Value = adoSrc.Fields.Item("VAT_PERIOD_END_DT").Value
                            .Fields.Item("YEAR_END_CLOSED").Value = adoSrc.Fields.Item("YEAR_END_CLOSED").Value
                            .Fields.Item("USER_NUMBER").Value = adoSrc.Fields.Item("USER_NUMBER").Value
                            .Fields.Item("VAT_RECON").Value = adoSrc.Fields.Item("VAT_RECON").Value
                            .Fields.Item("AMOUNT_TYPE").Value = adoSrc.Fields.Item("AMOUNT_TYPE").Value
                            .Fields.Item("ClientID").Value = adoSrc.Fields.Item("ClientID").Value
                            .Fields.Item("DeleteFlag").Value = adoSrc.Fields.Item("DeleteFlag").Value
                            .Fields.Item("TRANSACTION_REF").Value = adoSrc.Fields.Item("TRANSACTION_REF").Value
                        .Update
                         adoSrc.MoveNext
                     End With
                Wend
                adoSrc.Close
                adoDst.Close
                
                   Conn1.Execute "Drop TABLE NLPOSTINGTemp "
                   
                   UpdateDatabase2 = 1
              Exit Function
       Else
              Rst1.Close
              Set Rst1 = Nothing
       End If
Add_Client_LastModifiedBy:
     On Error GoTo Mod_Client_LastModifiedBy
     Rst1.Open "Select LastModifiedBy from Client;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Add_tlbAgreement_Fund
     Exit Function
Mod_Client_LastModifiedBy:
     Conn1.Execute "ALTER TABLE Client add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Client add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Property add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Property add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Units add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Units add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Tenants add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Tenants add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Supplier add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Supplier add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE LeaseDetails add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE LeaseDetails add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE DemandTypes add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE DemandTypes add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE PayableTypes add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE PayableTypes add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE ChargeTypes add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE ChargeTypes add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE NominalLedger add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE NominalLedger add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE DemandRecords add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE DemandRecords add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE tblPurInv add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tblPurInv add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE tlbReceipt add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbReceipt add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE tlbPayment add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbPayment add column LastModifiedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     
     Conn1.Execute "ALTER TABLE tlbBankPayment add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbBankPayment add column LastModifiedDate DateTime"
     
     Conn1.Execute "ALTER TABLE NJ_Header add column LastModifiedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE NJ_Header add column LastModifiedDate DateTime"
     
     
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     
     UpdateDatabase2 = 1
     Exit Function
Add_tlbAgreement_Fund:
'adde by anol  2023-04-22
     On Error GoTo Mod_tlbAgreement_Fund
     Rst1.Open "Select Fund from tlbAgreement;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE tlbAgreement ALTER COLUMN Fund LONG"
                Conn1.Execute "ALTER TABLE tlbAgreement ALTER COLUMN Frequency LONG"
                UpdateDatabase2 = 1
                Exit Function
     Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
     End If
     'Exit Function
Mod_tlbAgreement_Fund:
'adde by anol  2023-04-30
ADD_table_ReportViewBankBalance:
     On Error GoTo MOD_table_ReportViewBankBalance
     Rst1.Open "Select * from ReportViewBankBalance;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo ADD_table_ReportViewBankBalanceReconciled
     Exit Function
MOD_table_ReportViewBankBalance:
    Conn1.Execute "Create TABLE ReportViewBankBalance " & _
         "(" & _
            "ClientID  TEXT(100), " & _
            "BankCode     TEXT(100), " & _
            "BankName      TEXT(100), " & _
            "Amount  Currency " & _
            ");"
     UpdateDatabase2 = 1
     Exit Function
     
     'adde by anol  2023-05-11
ADD_table_ReportViewBankBalanceReconciled:
     On Error GoTo MOD_table_ReportViewBankBalanceReconciled
     Rst1.Open "Select * from ReportViewBankBalanceReconciled;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Add_RetentionDetails_BankCode
     Exit Function
MOD_table_ReportViewBankBalanceReconciled:
    Conn1.Execute "Create TABLE ReportViewBankBalanceReconciled " & _
        "(" & _
            "ClientID  TEXT(100), " & _
            "ClientName  TEXT(100), " & _
            "BankAccountName     TEXT(100), " & _
            "BankAccountNumber      TEXT(100), " & _
            "Receipt  Currency, " & _
            "Payment Currency, " & _
            "LastReconciledDate Date," & _
            "LastReconciledBankBalance  Currency, " & _
            "LastReconciledStatementBalance  Currency, " & _
            "UnreconciledCashBookBalance  Currency, " & _
            "CashbookcurrentBalance Currency " & _
        ");"
     UpdateDatabase2 = 1
     Exit Function


Add_RetentionDetails_BankCode:
    '2023-01-26 anol exclCRNtoCS field means you dont want to see some credit note in the CS Report
     On Error GoTo MOD_RetentionDetails_BankCode
     Rst1.Open "Select BankCode from RetentionDetails;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Add_tlbPayable_Fund
     Exit Function
MOD_RetentionDetails_BankCode:
     Conn1.Execute "ALTER TABLE RetentionDetails add column BankCode TEXT(15)"
     UpdateDatabase2 = 1
     Exit Function
     
Add_tlbPayable_Fund:
'adde by anol  2023-04-22
     On Error GoTo Mod_tlbPayable_Fund
     Rst1.Open "Select PAY_FUND from tlbPayable;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE tlbPayable ALTER COLUMN PAY_FUND LONG"
                UpdateDatabase2 = 1
                Exit Function
     Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
     End If
     'Exit Function
Mod_tlbPayable_Fund:
'adde by anol  2023-04-30
Add_LInsuranceCharges_InsuranceDept:
'adde by anol  2023-06-13
     On Error GoTo Mod_LInsuranceCharges_InsuranceDept
     Rst1.Open "Select InsuranceDept from LInsuranceCharges;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE LInsuranceCharges ALTER COLUMN InsuranceDept LONG"
                UpdateDatabase2 = 1
                Exit Function
     Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
     End If

Mod_LInsuranceCharges_InsuranceDept:

Add_tlbBankPayment_DEPT_ID:
'adde by anol  2023-06-13
     'On Error GoTo Mod_tlbBankPayment_DEPT_ID
     Rst1.Open "Select DEPT_ID from tlbBankPayment;", Conn1, adOpenKeyset, adLockReadOnly
     If Rst1.Fields(0).Type = 202 Then '202 means type text and 135 means datetime
                If Rst1.State = 1 Then
                   Rst1.Close
                End If
                Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN DEPT_ID LONG"
                UpdateDatabase2 = 1
                Exit Function
     Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
     End If
     
'Mod_tlbBankPayment_DEPT_ID:
'added by anol  2023-06-14

ADD_ClientID_LesseeAccTemplate:
        On Error GoTo MOD_ClientID_LesseeAccTemplate
        Rst1.Open "Select LesseeAccTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_ClientID_CSPreviewTemplate
        Exit Function

MOD_ClientID_LesseeAccTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN LesseeAccTemplate text(250);"
        UpdateDatabase2 = 1
        Exit Function
        
ADD_ClientID_CSPreviewTemplate:
        On Error GoTo MOD_ClientID_CSPreviewTemplate
        Rst1.Open "Select CSPreviewTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_ClientID_CSTemplate
        Exit Function
MOD_ClientID_CSPreviewTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN CSPreviewTemplate text(250);"
        UpdateDatabase2 = 1
        Exit Function
ADD_ClientID_CSTemplate:
        On Error GoTo MOD_ClientID_CSTemplate
        Rst1.Open "Select CSTemplate from Client;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo Add_Client_CreatedBy
        Exit Function
MOD_ClientID_CSTemplate:
        Conn1.Execute "ALTER TABLE Client ADD COLUMN CSTemplate text(250);"
        UpdateDatabase2 = 1
        Exit Function
 
Add_Client_CreatedBy:
     On Error GoTo Mod_Client_CreatedBy
     Rst1.Open "Select CreatedBy from Client;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo Modify_NJ_Header_propertyID
     Exit Function
Mod_Client_CreatedBy:
     Conn1.Execute "ALTER TABLE Client add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Client add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Property add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Property add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Units add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Units add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Tenants add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Tenants add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE Supplier add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE Supplier add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE LeaseDetails add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE LeaseDetails add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE DemandTypes add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE DemandTypes add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE PayableTypes add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE PayableTypes add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE ChargeTypes add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE ChargeTypes add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE NominalLedger add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE NominalLedger add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE DemandRecords add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE DemandRecords add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     Conn1.Execute "ALTER TABLE tblPurInv add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tblPurInv add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE tlbReceipt add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbReceipt add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     Conn1.Execute "ALTER TABLE tlbPayment add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbPayment add column CreatedDate DateTime"
     'Conn1.Execute "Update DemandRecords set  exclCRNtoCS =False"
     
     
     Conn1.Execute "ALTER TABLE tlbBankPayment add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE tlbBankPayment add column CreatedDate DateTime"
     
     Conn1.Execute "ALTER TABLE NJ_Header add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE NJ_Header add column CreatedDate DateTime"
     
     Conn1.Execute "ALTER TABLE RentSummaryStatement add column CreatedBy TEXT(255)"
     Conn1.Execute "ALTER TABLE RentSummaryStatement add column CreatedDate DateTime"
     UpdateDatabase2 = 1
Exit Function
Modify_NJ_Header_propertyID:
        Rst1.Open "Select  PropertyID from NJ_Header;", Conn1, adOpenKeyset, adLockReadOnly
        If Rst1.Fields.Item(0).DefinedSize < 10 Then
              Rst1.Close
              Set Rst1 = Nothing
              Conn1.Execute "ALTER TABLE NJ_Header ALTER column PropertyID TEXT(10)"
              UpdateDatabase2 = 1
               Exit Function
        Else
             If Rst1.State = 1 Then
                    Rst1.Close
            End If
        End If
Add_tblPurInvPreview:
        On Error GoTo mod_tblPurInvPreview
        Rst1.Open "Select DescriptionANDDates from tblPurInvPreview;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo Add_RentSummaryStatement_InclSupplierOS
        Exit Function
mod_tblPurInvPreview:
        Conn1.Execute "ALTER TABLE tblPurInvPreview ADD COLUMN DescriptionANDDates text(250);"
        UpdateDatabase2 = 1
        Exit Function
        
Add_RentSummaryStatement_InclSupplierOS:
        On Error GoTo Mod_RentSummaryStatement_InclSupplierOS
        Rst1.Open "Select InclSupplierOS from RentSummaryStatement;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        GoTo ADD_table_ClientStatementPurchasesSnapshot
        Exit Function
Mod_RentSummaryStatement_InclSupplierOS:
        Conn1.Execute "ALTER TABLE RentSummaryStatement ADD COLUMN InclSupplierOS BIT;"
        Conn1.Execute "ALTER TABLE RentSummaryStatement ADD COLUMN InclMngtFeesDue BIT;"
        UpdateDatabase2 = 1
        Exit Function
ADD_table_ClientStatementPurchasesSnapshot:
     On Error GoTo MOD_table_ClientStatementPurchasesSnapshot
     Rst1.Open "Select * from ClientStatementPurchasesSnapshot;", Conn1, adOpenKeyset, adLockReadOnly
     Rst1.Close
     GoTo MOD_ClientStatementPurchasesSnapshot_isManagementFee
     Exit Function
MOD_table_ClientStatementPurchasesSnapshot:
     Conn1.Execute "Create TABLE ClientStatementPurchasesSnapshot " & _
         "(" & _
            "StatementID Long, " & _
            "Type     INT, " & _
            "MY_ID      TEXT(25), " & _
            "SplitID      Long, " & _
            "TransactionID Long, " & _
            "ClientID  TEXT(100), " & _
            "PropertyID     TEXT(10), " & _
            "SupplierID     TEXT(250), " & _
            "TranDate  Date, " & _
            "NOMINAL_CODE      TEXT(10), " & _
            "PaymentDescription  TEXT(255), " & _
            "PaymentRef  TEXT(100), " & _
            "NetAmount  Currency, " & _
            "VATAmount    Currency, " & _
            "PaymentAmount  Currency, " & _
            "CreditAmount    Currency, " & _
            "OSAmount  Currency " & _
            ");"
    UpdateDatabase2 = 1
     Exit Function
MOD_ClientStatementPurchasesSnapshot_isManagementFee:
        On Error GoTo MODD_ClientStatementPurchasesSnapshot_isManagementFee
        Rst1.Open "Select isManagementFee from ClientStatementPurchasesSnapshot;", Conn1, adOpenKeyset, adLockReadOnly
        Rst1.Close
        'GoTo ADD_table_ClientStatementPurchasesSnapshot
        Exit Function
MODD_ClientStatementPurchasesSnapshot_isManagementFee:
        Conn1.Execute "ALTER TABLE ClientStatementPurchasesSnapshot ADD COLUMN isManagementFee BIT;"
        UpdateDatabase2 = 1
        Exit Function
        
End Function
Private Function UpdateDatabase3(Conn1 As ADODB.Connection) As Boolean

End Function
Private Sub Update2DecimalPlace(szTable As String, szTableID As String, szField As String)
   On Error GoTo Err_Catch

   Dim adoS       As New ADODB.Recordset
   Dim adod       As New ADODB.Recordset
   Dim szSQL      As String

   szSQL = "SELECT " & szTableID & ", " & szField & " " & _
           "FROM " & szTable & ";"

   adoS.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
   adod.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   While Not adod.EOF
'If adoD.Fields.Item(szTableID).Value = "1103181609414935206" Then
'MsgBox ""
'End If
'http://www.w3schools.com/ado/met_comm_createparameter.asp
      If adod.Fields.Item(szTableID).Type = 3 Then
         adoS.Find szTableID & " = " & adod.Fields.Item(szTableID).Value & "", , , 1
      Else
         adoS.Find szTableID & " = '" & adod.Fields.Item(szTableID).Value & "'", , , 1
      End If
'Debug.Print CCur(adoS.Fields.Item(szField).Value)
      adod.Fields.Item(szField).Value = RoundingNumber(adoS.Fields.Item(szField).Value, 2)
      adoS.MoveFirst
      adod.Update
      adod.MoveNext
   Wend

   adoS.Close
   Set adoS = Nothing
   adod.Close
   Set adod = Nothing

   Exit Sub
Err_Catch:
   Debug.Print Err.description
End Sub

Private Sub UpdateBR_BP_SameAcc()
   Dim szSQL      As String
   Dim Rst2       As New ADODB.Recordset

   szSQL = "SELECT * " & _
           "FROM   tlbBankPayment " & _
           "WHERE  BANK_AC = NOMINAL_CODE AND " & _
               "ISNULL(CT);"
'Debug.Print szSQL
   Rst1.Open szSQL, Conn1, adOpenStatic, adLockReadOnly
'Debug.Print Rst1.RecordCount

   szSQL = "SELECT * " & _
           "FROM   tlbBankPayment;"

   With Rst2
      .Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

      While Not Rst1.EOF
         .AddNew
         .Fields.Item("MY_ID").Value = UniqueID()
         .Fields.Item("ClientID").Value = Rst1.Fields.Item("ClientID").Value
         .Fields.Item("BANK_AC").Value = Rst1.Fields.Item("BANK_AC").Value
         .Fields.Item("PropertyID").Value = Rst1.Fields.Item("PropertyID").Value
         .Fields.Item("UNIT_ID").Value = Rst1.Fields.Item("UNIT_ID").Value
         .Fields.Item("DESCRIPTION").Value = Rst1.Fields.Item("DESCRIPTION").Value
         .Fields.Item("PROJ_REF").Value = Rst1.Fields.Item("PROJ_REF").Value
         .Fields.Item("NOMINAL_CODE").Value = Rst1.Fields.Item("NOMINAL_CODE").Value
         .Fields.Item("DEPT_ID").Value = Rst1.Fields.Item("DEPT_ID").Value
         .Fields.Item("TRAN_DATE").Value = Rst1.Fields.Item("TRAN_DATE").Value
         .Fields.Item("NET_AMOUNT").Value = Rst1.Fields.Item("NET_AMOUNT").Value
         .Fields.Item("TAX_CODE").Value = Rst1.Fields.Item("TAX_CODE").Value
         .Fields.Item("VAT").Value = Rst1.Fields.Item("VAT").Value

         If Rst1.Fields.Item("TransactionType").Value = 11 Then
            .Fields.Item("TransactionType").Value = 12
            .Fields.Item("TRANS").Value = "BR"
         End If
         If Rst1.Fields.Item("TransactionType").Value = 12 Then
            .Fields.Item("TransactionType").Value = 11
            .Fields.Item("TRANS").Value = "BP"
         End If
         .Fields.Item("TRAN_ID").Value = SlNumber(.Fields.Item("TRANS").Value, "tlbBankPayment", Conn1)
         .Fields.Item("CT").Value = "F"

         .Update
         Conn1.Execute "UPDATE tlbBankPayment SET CT = 'F' WHERE MY_ID = '" & Rst1.Fields.Item("MY_ID").Value & "';"
         Rst1.MoveNext
      Wend
      .Close
   End With

   Rst1.Close
End Sub

Private Function UpdateDatabase(Conn1 As ADODB.Connection)
   Dim i As Long, szaPath() As String, szPath As String, szSQL As String
'   DoEvents
   UpdateDatabase = 0
'
'  SKIP some updates.
   GoTo ADDNEW_REC_DPTYP


'   Check the update of the database.
'   New table on 02/03/07 tblPoA
'###############################################################################################################
'   On Error GoTo MissingTable_tblPoA
'
'   Rst1.Open "SELECT * FROM tblPoA;", conn1,adOpenStatic,adLockReadOnly
'
'   Rst1.Close
'   GoTo GLOBAL_RC
'
'MissingTable_tblPoA:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblPoA"
'   exit function

'   New table on 18/04/07 GlobalRC
'###############################################################################################################
GLOBAL_RC:
   On Error GoTo MissingTable_GlobalRC

   Rst1.Open "SELECT * FROM GlobalRC;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo GLOBAL_LEASE_UPDATE

MissingTable_GlobalRC:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - GlobalRC"
   UpdateDatabase = -1
   Exit Function

'   New table on 24/04/07 tblPrevGLU
'###############################################################################################################
GLOBAL_LEASE_UPDATE:
   On Error GoTo MissingTable_GLOBAL_LEASE_UPDATE

   Rst1.Open "SELECT * FROM tblPrevGLU;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo GLOBAL_INTEREST

MissingTable_GLOBAL_LEASE_UPDATE:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblPrevGLU"
   UpdateDatabase = -1
   Exit Function

'   New table on 07/06/07 InterestRates
'###############################################################################################################
GLOBAL_INTEREST:
   On Error GoTo MissingTable_GLOBAL_INTEREST

   Rst1.Open "SELECT * FROM InterestRates;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo LEASEID_RENTANALYSIS_

MissingTable_GLOBAL_INTEREST:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - InterestRates"
   UpdateDatabase = -1
   Exit Function

'   Altering field length of LeaseID on 14/06/07 RentAnalysis
'###############################################################################################################
LEASEID_RENTANALYSIS_:
   On Error GoTo FieldSize_LEASEID_RENTANALYSIS

   If Leaseid_RentAnalysis Then GoTo Text1

FieldSize_LEASEID_RENTANALYSIS:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - Leaseid_RentAnalysis"
   UpdateDatabase = 1
   Exit Function

'   Altering field length of Text1 on 14/06/07 LeaseDetails
'###############################################################################################################
Text1:
   On Error GoTo MissingTable_TEXT1

   Rst1.Open "SELECT Text1 FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Text1").DefinedSize < 100 Then
      Rst1.Close
      Set Rst1 = Nothing
      Conn1.Execute "ALTER TABLE LeaseDetails ALTER COLUMN Text1 text(100)"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo GLOBAL_INTEREST_RATE

MissingTable_TEXT1:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - Text1"
   UpdateDatabase = 1
   Exit Function

'  In the Global table if BaseInterestRate is not using as foreign key of InterestRates then set the value of
'  BaseInterestRate = 0; on 17/06/2007    GlobaData
'###############################################################################################################
GLOBAL_INTEREST_RATE:
   On Error GoTo MissingTable_GLOBAL_INTEREST_RATE
   Dim Rst2 As New ADODB.Recordset

   Rst1.Open "SELECT BaseInterestRate FROM GlobalData;", Conn1, adOpenDynamic, adLockOptimistic

   If Not Rst1.EOF Then
      If Not IsNull(Rst1!BaseInterestRate) Or Rst1!BaseInterestRate <> "" Then
         If Val(Rst1!BaseInterestRate) > 0 Then
            Rst2.Open "SELECT * FROM GlobalData, InterestRates " & _
                        "WHERE InterestRates.RateID = " & Val(Rst1!BaseInterestRate) & ";", Conn1, adOpenStatic, adLockReadOnly
            If Rst2.EOF Then
               If Val(Rst1!BaseInterestRate) - CInt(Rst1!BaseInterestRate) <> 0 Then
                  Rst1!BaseInterestRate = 0
                  Rst1.Update
                  MsgBox "Base interest rate has been changed of this company." & Chr(10) & _
                         "Please reset it through the Global form.", vbExclamation + vbOKOnly, "Global Interest Rate"
               End If
            End If
            Rst2.Close
         End If
      End If
   End If
   Rst1.Close
   GoTo ADD_NEW_COLUMNS

MissingTable_GLOBAL_INTEREST_RATE:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - GlobaData"
   UpdateDatabase = 1
   Exit Function

'   Add new columns on 19/06/07 tlbReceipt
'###############################################################################################################
ADD_NEW_COLUMNS:
   On Error GoTo MissingTable_ADD_NEW_COLUMNS

   Rst1.Open "SELECT * FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   If Val(Rst1.Fields.Count) = 24 Then
      Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN IntCalDate TEXT(25)"
   End If

   Rst1.Close
   GoTo ADDNEW_REC_RAT

MissingTable_ADD_NEW_COLUMNS:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new record on 31/07/07 PrimaryCode
'###############################################################################################################
ADDNEW_REC_RAT:
   On Error GoTo MissingTable_ADDNEW_REC_RAT

   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'RAT';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "RAT"
      Rst1!Value = "RECEIPT AMOUNT TYPE"
      Rst1.Update
   End If
   Rst1.Close

   GoTo ADDNEW_COL_RptAmtType

MissingTable_ADDNEW_REC_RAT:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 31/07/07 tlbReceipt
'###############################################################################################################
ADDNEW_COL_RptAmtType:
   On Error GoTo MissingTable_ADDNEW_COL_RptAmtType

   Rst1.Open "SELECT * FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   If Val(Rst1.Fields.Count) = 25 Then
      Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN RptAmtType TEXT(10)"
   End If
   Rst1.Close

   GoTo TENANT_DEPOSIT

MissingTable_ADDNEW_COL_RptAmtType:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col - RptAmtType) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   New table on 10/08/07 TenantDeposit
'###############################################################################################################
TENANT_DEPOSIT:
   On Error GoTo MissingTable_TENANT_DEPOSIT

   Rst1.Open "SELECT * FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADDNEW_REC_DPTYP

MissingTable_TENANT_DEPOSIT:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - Import TenantDeposit"
   UpdateDatabase = -1
   Exit Function

'   Add new record on 14/08/07 PrimaryCode
'###############################################################################################################
ADDNEW_REC_DPTYP:
   On Error GoTo MissingTable_ADDNEW_REC_DPTYP

   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'EXPTYP';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "EXPTYP"
      Rst1!Value = "EXPENSES TYPE"
      Rst1.Update
   End If
   Rst1.Close
   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'DPTYP';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "DPTYP"
      Rst1!Value = "DEPOSIT TYPE"
      Rst1.Update
   End If
   Rst1.Close
   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'RTYP';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "RTYP"
      Rst1!Value = "REFUND TYPE"
      Rst1.Update
   End If
   Rst1.Close
   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'RPT';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "RPT"
      Rst1!Value = "DEFAULT REPORT NAME"
      Rst1!Flexible = False
      Rst1.Update
   End If
   Rst1.Close
   'Primary Code ATLD
   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'ATLD';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then                       '22/01/2008
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "ATLD"
      Rst1!Value = "ALARM TERMINATE LEASE DAYS"
      Rst1!Flexible = True
      Rst1.Update
   End If
   Rst1.Close

   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'ATLD';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "ATLD"
         !Code = "TL"
         !Value = "0"
         .Update
         'Secondary code to keep track of Alarm Y/N
         .AddNew
         !PrimaryCode = "ATLD"
         !Code = "TA"
         !Value = "N"
         .Update
      End With
   End If
   Rst1.Close

   'Primary Code IPT             '20/03/2009
   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'IPT';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "IPT"
      Rst1!Value = "INSPECTOR"
      Rst1!Flexible = True
      Rst1.Update
   End If
   Rst1.Close

   GoTo RESIZE_DESCRIPTION_DEMAND_SPLIT_RECORD

MissingTable_ADDNEW_REC_DPTYP:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - DPTYP) - PrimaryCode"
   UpdateDatabase = 1
   Exit Function
''  I have marged it in the previous checking
''   Add new record on 15/08/07 PrimaryCode
''###############################################################################################################
'ADDNEW_REC_RTYP:
'   On Error GoTo MissingTable_ADDNEW_REC_RTYP
'
'   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'RTYP';", conn1,adOpenStatic,adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
'      Rst1.AddNew
'      Rst1!Code = "RTYP"
'      Rst1!Value = "REFUND TYPE"
'      Rst1.Update
'      Rst1.Close
'   Else
'      Rst1.Close
'   End If
'
'   GoTo RESIZE_DESCRIPTION_DEMAND_SPLIT_RECORD
'
'MissingTable_ADDNEW_REC_RTYP:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - PrimaryCode"
'   exit function

'   Extend the field size on 11/09/07 DemandSplitRecords
'###############################################################################################################
RESIZE_DESCRIPTION_DEMAND_SPLIT_RECORD:

   On Error GoTo MissingTable_RESIZE_DESCRIPTION_DEMAND_SPLIT_RECORD

   Rst1.Open "SELECT Description FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Description").DefinedSize = 50 Or Rst1.Fields.Item("Description").DefinedSize = 70 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN Description TEXT(255)"
   End If
   Rst1.Close
   Set Rst1 = Nothing

   GoTo RESIZE_DESCRIPTION_DR_CURRENT_PRINT

MissingTable_RESIZE_DESCRIPTION_DEMAND_SPLIT_RECORD:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - Description of DSR"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size on 28/01/08 tlbDRCurrentPrint
'###############################################################################################################
RESIZE_DESCRIPTION_DR_CURRENT_PRINT:

   On Error GoTo MissingTable_RESIZE_DESCRIPTION_DR_CURRENT_PRINT

   Rst1.Open "SELECT Description FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Description").DefinedSize = 50 Or Rst1.Fields.Item("Description").DefinedSize = 70 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN Description TEXT(255)"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo ADDNEW_COL_DR

MissingTable_RESIZE_DESCRIPTION_DR_CURRENT_PRINT:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - Description of DSR"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 27/09/07 DemandRecords
'###############################################################################################################
ADDNEW_COL_DR:
   On Error GoTo MISSING_ADDNEW_COL_DR

   Rst1.Open "SELECT AdjTag FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_LeaseRef_DR

MISSING_ADDNEW_COL_DR:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN AdjTag TEXT(1)"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new column "LeaseRef" on 27/09/07 DemandRecords
'###############################################################################################################
ADDNEW_COL_LeaseRef_DR:
   On Error GoTo MISSING_ADDNEW_COL_LeaseRef_DR

   Rst1.Open "SELECT LeaseRef FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_AdjTag

MISSING_ADDNEW_COL_LeaseRef_DR:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN LeaseRef TEXT(20)"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 11/09/07 tlbReceipt
'###############################################################################################################
ADDNEW_COL_AdjTag:
   On Error GoTo MISSING_ADDNEW_COL_AdjTag

   Rst1.Open "SELECT AdjTag FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_DETAILS_DR

MISSING_ADDNEW_COL_AdjTag:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN AdjTag TEXT(1)"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_AdjTag"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 28/09/07 DemandRecords
'###############################################################################################################
ADDNEW_COL_DETAILS_DR:
   On Error GoTo MISSING_ADDNEW_COL_DETAILS_DR

   Rst1.Open "SELECT Details FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo TableMissing_RptTransactions

MISSING_ADDNEW_COL_DETAILS_DR:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN Details TEXT(70);"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   New table on 28/09/07 RptTransactions
'###############################################################################################################
TableMissing_RptTransactions:
   On Error GoTo EH_TableMissing_RptTransactions

   Rst1.Open "SELECT * FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo MODIFY_COL_DemandRef_RPT

EH_TableMissing_RptTransactions:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: RptTransactions"
   UpdateDatabase = -1
   Exit Function

'   Modify column data type DemandRef on 01/10/07 tlbReceipt
'###############################################################################################################
MODIFY_COL_DemandRef_RPT:
   On Error GoTo CHANGE_MODIFY_COL_DemandRef_RPT

   Rst1.Open "SELECT DemandRef FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type <> adInteger Then
      If Rst1.RecordCount > 0 Then
         szSQL = "UPDATE tlbReceipt, DemandRecords, DemandSplitRecords " & _
                 "SET tlbReceipt.DemandRef = DemandRecords.DemandID " & _
                 "WHERE tlbReceipt.DemandRef = DemandSplitRecords.DSR AND " & _
                     "DemandSplitRecords.DemandID = DemandRecords.DemandID;"
         Conn1.Execute szSQL
      End If
      Rst1.Close
      Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN DemandRef Long;"
      GoTo CHANGE_MODIFY_COL_DemandRef_RPT
   Else
      Rst1.Close
   End If

   GoTo ADDNEW_COL_DemandTypes_DTGroup

CHANGE_MODIFY_COL_DemandRef_RPT:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "MODIFY_COL_DemandRef_RPT"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 19/10/07 DemandTypes
'###############################################################################################################
ADDNEW_COL_DemandTypes_DTGroup:
   On Error GoTo MISSING_ADDNEW_COL_DemandTypes_DTGroup

   Rst1.Open "SELECT DTGroup FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo DATABASE_MissingTable_Schedule

MISSING_ADDNEW_COL_DemandTypes_DTGroup:
'  Add new column
   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN DTGroup SHORT;"
'  Assign group number for existing Demand Types
   Rst1.Open "SELECT DTGroup FROM DemandTypes;", Conn1, adOpenDynamic, adLockOptimistic

   i = 0
   While Not Rst1.EOF
      i = i + 1
      Rst1!DTGroup = i
      Rst1.Update
      Rst1.MoveNext
   Wend
   Rst1.Close
   
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DemandTypes_DTGroup"
   UpdateDatabase = 1
   Exit Function

'   New table on 23/10/07 Schedule
'###############################################################################################################
DATABASE_MissingTable_Schedule:
   On Error GoTo Missing_Table_Schedule

   Rst1.Open "SELECT * FROM Schedule;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo DATABASE_MissingTable_Template

Missing_Table_Schedule:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table - Schedule"
   UpdateDatabase = -1
   Exit Function

'   New table on 23/10/07 Template
'###############################################################################################################
DATABASE_MissingTable_Template:
   On Error GoTo Missing_Table_Template

   Rst1.Open "SELECT * FROM Template;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADDNEW_COL_TemplatePrint_Tenants

Missing_Table_Template:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table - Template"
   UpdateDatabase = -1
   Exit Function

'   Add new column TemplatePrint on 23/10/07 Tenants
'###############################################################################################################
ADDNEW_COL_TemplatePrint_Tenants:
   On Error GoTo MISSING_ADDNEW_COL_TemplatePrint_Tenants

   Rst1.Open "SELECT TemplatePrint FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_DemandReportName_DemandTypes

MISSING_ADDNEW_COL_TemplatePrint_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN TemplatePrint TEXT(1);"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new column DemandReportName on 23/10/07 DemandTypes
'###############################################################################################################
ADDNEW_COL_DemandReportName_DemandTypes:
   On Error GoTo MISSING_ADDNEW_COL_DemandReportName_DemandTypes

   Rst1.Open "SELECT DemandReportName FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_EmailInvoiceTemplate_DemandTypes

MISSING_ADDNEW_COL_DemandReportName_DemandTypes:
   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN DemandReportName TEXT(100);"
   Conn1.Execute "UPDATE DemandTypes SET DemandReportName = 'InvoiceDemand.rpt';"
   UpdateDatabase = 1
   Exit Function

'   Add new column EmailInvoiceTemplate on 29/03/2011 DemandTypes
'###############################################################################################################
ADDNEW_COL_EmailInvoiceTemplate_DemandTypes:
   On Error GoTo MISSING_ADDNEW_COL_EmailInvoiceTemplate_DemandTypes

   Rst1.Open "SELECT EmailInvoiceTemplate FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo MODIFY_TypeOfDemand_tlbDRCurrentPrint

MISSING_ADDNEW_COL_EmailInvoiceTemplate_DemandTypes:
   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN EmailInvoiceTemplate TEXT(100);"
   Conn1.Execute "UPDATE DemandTypes " & _
                 "SET EmailInvoiceTemplate = Mid(DemandReportName,1, Len(DemandReportName) -4) & '_Email.rpt';"
   UpdateDatabase = 1
   Exit Function

'   Modify DATA TYPE TypeOfDemand on 24/10/07 tlbDRCurrentPrint
'###############################################################################################################
MODIFY_TypeOfDemand_tlbDRCurrentPrint:
   On Error GoTo CHANGE_MODIFY_TypeOfDemand_tlbDRCurrentPrint

   Rst1.Open "SELECT TypeOfDemand FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = adUnsignedTinyInt Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN TypeOfDemand BYTE;"
      GoTo CHANGE_MODIFY_TypeOfDemand_tlbDRCurrentPrint
   End If

   GoTo ADDNEW_COL_FirstPay_Tenants

CHANGE_MODIFY_TypeOfDemand_tlbDRCurrentPrint:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase = 1
   Exit Function

'   Add a new columns FirstPay on 24/10/07 Tenants
'###############################################################################################################
ADDNEW_COL_FirstPay_Tenants:
   On Error GoTo MISSING_ADDNEW_COL_FirstPay_Tenants

   Rst1.Open "SELECT FirstPay FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADDNEW_COL_4Cols_Tenants

MISSING_ADDNEW_COL_FirstPay_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN FirstPay CURRENCY;"

'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new 4 columns on 24/10/07 Tenants
'###############################################################################################################
ADDNEW_COL_4Cols_Tenants:
   On Error GoTo MISSING_ADDNEW_COL_4Cols_Tenants

   Rst1.Open "SELECT FurtherPay FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo MODIFY_REPORT_PATH_SecondaryCode

MISSING_ADDNEW_COL_4Cols_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN FurtherPay CURRENCY;"
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN Freq TEXT(50);"
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN Ref TEXT(50);"
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN StDate DATE;"

'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Modify value of FPATH2 on 25/10/07 SecondaryCode
'###############################################################################################################
MODIFY_REPORT_PATH_SecondaryCode:
   Rst1.Open "SELECT VALUE FROM SECONDARYCODE WHERE Code = 'FPATH2';", Conn1, adOpenDynamic, adLockOptimistic

   szaPath = Split(Rst1!Value, "\")

   If szaPath(1) = "CompanyReports" Then
      Rst1.Close
      GoTo MISSING_TABLE_TemplateUnitSelection
   End If

   szPath = "\CompanyReports"

   For i = 2 To UBound(szaPath) - 1
      szPath = szPath & "\" & szaPath(i)
   Next i

   Rst1!Value = szPath
   Rst1.Update
   Rst1.Close
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   New table on 25/10/07 TemplateUnitSelection
'###############################################################################################################
MISSING_TABLE_TemplateUnitSelection:
   On Error GoTo ERR_MISSING_TABLE_TemplateUnitSelection

   Rst1.Open "SELECT * FROM TemplateUnitSelection;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADDNEW_REC_INV

ERR_MISSING_TABLE_TemplateUnitSelection:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - TemplateUnitSelection"
   UpdateDatabase = -1
   Exit Function

'   Add new record on 26/10/07 SecondaryCode
'###############################################################################################################
ADDNEW_REC_INV:
   On Error GoTo MissingRec_ADDNEW_REC_INV

   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'RPT' AND Code = 'INV';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "RPT"
         !Code = "INV"
         !Value = "Demand Invoice"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADDNEW_COL_TemplatePrint_tlbDRCurrentPrint

MissingRec_ADDNEW_REC_INV:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new column on 27/10/07 tlbDRCurrentPrint
'###############################################################################################################
ADDNEW_COL_TemplatePrint_tlbDRCurrentPrint:
   On Error GoTo MISSING_ADDNEW_COL_TemplatePrint_tlbDRCurrentPrint

   Rst1.Open "SELECT DemandReportName FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_REC_GR

MISSING_ADDNEW_COL_TemplatePrint_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN DemandReportName TEXT(100);"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new record on 31/10/07 PrimaryCode
'###############################################################################################################
ADDNEW_REC_GR:
   On Error GoTo MissingRec_REC_GR

   Rst1.Open "SELECT Code FROM PrimaryCode WHERE Code = 'GR';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "GR"
         !Value = "GROUP RANGE"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADDNEW_REC_STRNG

MissingRec_REC_GR:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new record on 31/10/07 SecondaryCode
'###############################################################################################################
ADDNEW_REC_STRNG:
   On Error GoTo MissingRec_ADDNEW_REC_STRNG

   Rst1.Open "SELECT * FROM SecondaryCode WHERE PrimaryCode = 'GR' AND Code = 'STRNG';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "GR"
         !Code = "STRNG"
         !Value = "1"
         !description = "START RANGE"
         .Update
         .AddNew
         !PrimaryCode = "GR"
         !Code = "ENDRNG"
         !Value = "9"
         !description = "END RANGE"
         .Update
      End With
   Else
      If Rst1!Value = "START RANGE" Then
         Rst1.Close
         Conn1.Execute "UPDATE SecondaryCode SET SecondaryCode.Value = '1' WHERE PrimaryCode = 'GR' AND Code = 'STRNG';"
         Conn1.Execute "UPDATE SecondaryCode SET SecondaryCode.Value = '9' WHERE PrimaryCode = 'GR' AND Code = 'ENDRNG';"
      End If
   End If
   Rst1.Close

   GoTo ADD_ScheduleID_LServiceCharges

MissingRec_ADDNEW_REC_STRNG:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add a column ScheduleID on 01/11/07 LServiceCharges
'###############################################################################################################
ADD_ScheduleID_LServiceCharges:
   On Error GoTo CHANGE_ADD_ScheduleID_LServiceCharges

   Rst1.Open "SELECT ScheduleID FROM LServiceCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_DTCategory_tlbDRCurrentPrint

CHANGE_ADD_ScheduleID_LServiceCharges:
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN ScheduleID LONG;"
   Conn1.Execute "UPDATE LServiceCharges SET ScheduleID = 0;"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase = 1
   Exit Function

'   Add new column DTCategory on 02/11/07 tlbDRCurrentPrint
'###############################################################################################################
ADDNEW_COL_DTCategory_tlbDRCurrentPrint:
   On Error GoTo MISSING_ADDNEW_COL_DTCategory_tlbDRCurrentPrint

   Rst1.Open "SELECT DTCategory FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_DeptNumber_Property

MISSING_ADDNEW_COL_DTCategory_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN DTCategory BYTE;"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new column DeptNumber on 05/11/07 Property
'###############################################################################################################
ADDNEW_COL_DeptNumber_Property:
   On Error GoTo MISSING_ADDNEW_COL_DeptNumber_Property

   Rst1.Open "SELECT DeptNumber FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_REC_FREQ

MISSING_ADDNEW_COL_DeptNumber_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DeptNumber INTEGER;"
   Conn1.Execute "UPDATE Property SET DeptNumber = 999;"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new record on 06/11/07 Frequencies
'###############################################################################################################
ADDNEW_REC_FREQ:
   On Error GoTo MissingRec_ADDNEW_REC_FREQ

   Rst1.Open "SELECT * FROM Frequencies WHERE ID = 16;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM Frequencies;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Id = 16
         !Frequency = "4 Monthly in Advance"
         !CalDays = "4m"
         !PartOfYear = 3
         .Update
         .AddNew
         !Id = 17
         !Frequency = "4 Monthly in Arrears"
         !CalDays = "-4m"
         !PartOfYear = 3
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADD_StopSC_LServiceCharges

MissingRec_ADDNEW_REC_FREQ:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function
'
''   Add new column ScheduleID on 08/11/07 tblPurInv
''###############################################################################################################
'ADDNEW_COL_ScheduleID_tblPurInv:
'   On Error GoTo MISSING_ADDNEW_COL_ScheduleID_tblPurInv
'
'   Rst1.Open "SELECT ScheduleID FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_StopSC_LServiceCharges
'
'MISSING_ADDNEW_COL_ScheduleID_tblPurInv:
'   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN ScheduleID LONG;"
''   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
'   UpdateDatabase = 1
'   exit function

'   Add a column StopSC on 16/11/07 LServiceCharges
'###############################################################################################################
ADD_StopSC_LServiceCharges:
   On Error GoTo CHANGE_ADD_StopSC_LServiceCharges

   Rst1.Open "SELECT StopSC FROM LServiceCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_StopRC_LRentCharges

CHANGE_ADD_StopSC_LServiceCharges:
   Conn1.Execute "ALTER TABLE LServiceCharges ADD COLUMN StopSC TEXT(20);"
   UpdateDatabase = 1
   Exit Function

'   Add a column StopRC on 16/11/07 LRentCharges
'###############################################################################################################
ADD_StopRC_LRentCharges:
   On Error GoTo CHANGE_ADD_StopRC_LRentCharges

   Rst1.Open "SELECT StopRC FROM LRentCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_StopIC_LInsuranceCharges

CHANGE_ADD_StopRC_LRentCharges:
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN StopRC TEXT(20);"
   UpdateDatabase = 1
   Exit Function

'   Add a column StopIC on 16/11/07 LInsuranceCharges
'###############################################################################################################
ADD_StopIC_LInsuranceCharges:
   On Error GoTo CHANGE_ADD_StopIC_LInsuranceCharges

   Rst1.Open "SELECT StopIC FROM LInsuranceCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_REMINDER_ID_PropertyMaintHistory

CHANGE_ADD_StopIC_LInsuranceCharges:
   Conn1.Execute "ALTER TABLE LInsuranceCharges ADD COLUMN StopIC TEXT(20);"
   UpdateDatabase = 1
   Exit Function

'   Add a column REMINDER_ID on 16/11/07 PropertyMaintHistory
'###############################################################################################################
ADD_REMINDER_ID_PropertyMaintHistory:
   On Error GoTo CHANGE_ADD_REMINDER_ID_PropertyMaintHistory

   Rst1.Open "SELECT REMINDER_ID FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_REMINDER_ID_UnitMaintHistory

CHANGE_ADD_REMINDER_ID_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN REMINDER_ID TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column REMINDER_ID on 21/11/07 UnitMaintHistory
'###############################################################################################################
ADD_REMINDER_ID_UnitMaintHistory:
   On Error GoTo CHANGE_ADD_REMINDER_ID_UnitMaintHistory

   Rst1.Open "SELECT REMINDER_ID FROM UnitMaintHistory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_REMINDER_ID_TenantEventHistory

CHANGE_ADD_REMINDER_ID_UnitMaintHistory:
   Conn1.Execute "ALTER TABLE UnitMaintHistory ADD COLUMN REMINDER_ID TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column REMINDER_ID on 22/11/07 TenantEventHistory
'###############################################################################################################
ADD_REMINDER_ID_TenantEventHistory:
   On Error GoTo CHANGE_ADD_REMINDER_ID_TenantEventHistory

   Rst1.Open "SELECT REMINDER_ID FROM TenantEventHistory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_REMINDER_ID_RentAnalysis

CHANGE_ADD_REMINDER_ID_TenantEventHistory:
   Conn1.Execute "ALTER TABLE TenantEventHistory ADD COLUMN REMINDER_ID TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column REMINDER_ID on 25/01/08 RentAnalysis
'###############################################################################################################
ADD_REMINDER_ID_RentAnalysis:
   On Error GoTo CHANGE_ADD_REMINDER_ID_RentAnalysis

   Rst1.Open "SELECT REMINDER_ID FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_RRDemandType_RentAnalysis

CHANGE_ADD_REMINDER_ID_RentAnalysis:
   Conn1.Execute "ALTER TABLE RentAnalysis ADD COLUMN REMINDER_ID TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column RRDemandType on 28/03/08 RentAnalysis
'###############################################################################################################
ADD_RRDemandType_RentAnalysis:
   On Error GoTo CHANGE_ADD_RRDemandType_RentAnalysis

   Rst1.Open "SELECT RRDemandType FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_COMMENTS_RentAnalysis

CHANGE_ADD_RRDemandType_RentAnalysis:
   Conn1.Execute "ALTER TABLE RentAnalysis ADD COLUMN RRDemandType Byte;"
   UpdateDatabase = 1
   Exit Function

'   Add a column COMMENTS on 28/03/08 RentAnalysis
'###############################################################################################################
ADD_COMMENTS_RentAnalysis:
   On Error GoTo CHANGE_ADD_COMMENTS_RentAnalysis

   Rst1.Open "SELECT Comments FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Status_RentAnalysis

CHANGE_ADD_COMMENTS_RentAnalysis:
   Conn1.Execute "ALTER TABLE RentAnalysis ADD COLUMN Comments TEXT(255);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Status on 28/03/08 RentAnalysis
'###############################################################################################################
ADD_Status_RentAnalysis:
   On Error GoTo CHANGE_ADD_Status_RentAnalysis

   Rst1.Open "SELECT Status FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Landlord_Property

CHANGE_ADD_Status_RentAnalysis:
   Conn1.Execute "ALTER TABLE RentAnalysis ADD COLUMN Status TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Landlord on 19/11/07 Property
'###############################################################################################################
ADD_Landlord_Property:
   On Error GoTo CHANGE_ADD_Landlord_Property

   Rst1.Open "SELECT Landlord FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo Modify_table_PropertyMaintHistory

CHANGE_ADD_Landlord_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN Landlord TEXT(100);"
   UpdateDatabase = 1
   Exit Function

'   Modify the TABLE on 06/12/07 PropertyMaintHistory
'###############################################################################################################
Modify_table_PropertyMaintHistory:
   On Error GoTo Error_MODIFY_TABLE_PropertyMaintHistory

   Rst1.Open "SELECT Job_DiaryName FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   GoTo ADD_RemindTime_PropertyMaintHistory

Error_MODIFY_TABLE_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ALTER ID TEXT(10);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Job_DiaryName TEXT(40);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory DROP COLUMN Description;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN AssignedTo TEXT(50);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory DROP COLUMN Contact;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ExpectedStartDate TEXT(25);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ExpectedCompletionDate TEXT(25);"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Detail MEMO;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN BudgetCost CURRENCY;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory DROP COLUMN EstimateCost;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN ActualCost CURRENCY;"
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN RecordType TEXT(1);"

'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_COL_SPARE4"
   UpdateDatabase = 1
   Exit Function

'   Add a column RemindTime on 10/12/07 PropertyMaintHistory
'###############################################################################################################
ADD_RemindTime_PropertyMaintHistory:
   On Error GoTo CHANGE_ADD_RemindTime_PropertyMaintHistory

   Rst1.Open "SELECT RemindTime FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_spare1_tlbReceipt

CHANGE_ADD_RemindTime_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN RemindTime TEXT(10);"
   UpdateDatabase = 1
   Exit Function

'   Add a column spare1 on 11/12/07 tlbReceipt
'###############################################################################################################
ADD_spare1_tlbReceipt:
   On Error GoTo CHANGE_ADD_spare1_tlbReceipt

   Rst1.Open "SELECT spare1 FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Flexible_PrimaryCode

CHANGE_ADD_spare1_tlbReceipt:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN spare1 TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Modify Value of Flexible on 13/12/07 PrimaryCode
'###############################################################################################################
ADD_Flexible_PrimaryCode:
   Conn1.Execute "UPDATE PrimaryCode SET Flexible = TRUE WHERE Flexible = FALSE AND Code = 'MTYP';"

   GoTo TableMissing_Supplier

'   New table on 13/12/07 Supplier
'###############################################################################################################
TableMissing_Supplier:
   On Error GoTo EH_TableMissing_Supplier

   Rst1.Open "SELECT * FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADDNEW_REC_SCODE

EH_TableMissing_Supplier:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: Supplier"
   UpdateDatabase = -1
   Exit Function

'   Add new record on 09/01/08 PrimaryCode
'###############################################################################################################
ADDNEW_REC_SCODE:
   On Error GoTo MissingRec_REC_SCODE

   Rst1.Open "SELECT CODE FROM PrimaryCode WHERE Code = 'SCODE';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "SCODE"
         !Value = "SUPPLIER CODE"
         .Update
      End With
   End If
   Rst1.Close

   GoTo MODIFY_DT_GlobalSC

MissingRec_REC_SCODE:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Change datatype of a column on 10/01/08 GlobalSC
'###############################################################################################################
MODIFY_DT_GlobalSC:
   On Error GoTo CHANGE_MODIFY_DT_GlobalSC

   Rst1.Open "SELECT Fund FROM GlobalSC;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE GlobalSC ALTER COLUMN Fund LONG;"
      GoTo CHANGE_MODIFY_DT_GlobalSC
   End If

   GoTo ADD_GenerateDemand_DemandRecords

CHANGE_MODIFY_DT_GlobalSC:
   UpdateDatabase = 1
   Exit Function

'   Add a column GenerateDemand on 16/01/08 DemandRecords
'###############################################################################################################
ADD_GenerateDemand_DemandRecords:
   On Error GoTo CHANGE_ADD_GenerateDemand_DemandRecords

   Rst1.Open "SELECT GenerateDemand FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo TableMissing_PROPERTYLANDLORD

CHANGE_ADD_GenerateDemand_DemandRecords:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN GenerateDemand TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   New table on 16/01/08 PROPERTYLANDLORD
'###############################################################################################################
TableMissing_PROPERTYLANDLORD:
   On Error GoTo EH_TableMissing_PROPERTYLANDLORD

   Rst1.Open "SELECT * FROM PROPERTYLANDLORD;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo TableMissing_Landlord

EH_TableMissing_PROPERTYLANDLORD:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: PROPERTYLANDLORD"
   UpdateDatabase = -1
   Exit Function
'
''   Add a column CL_ID on 17/01/08 tblPurInv
''###############################################################################################################
'ADD_CL_ID_tblPurInv:
'   On Error GoTo CHANGE_ADD_CL_ID_tblPurInv
'
'   Rst1.Open "SELECT CL_ID FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo TableMissing_Landlord
'
'CHANGE_ADD_CL_ID_tblPurInv:
'   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN CL_ID TEXT(10);"
'   Conn1.Execute "UPDATE tblPurInv, Client SET CL_ID = Client.ClientID;"
'   UpdateDatabase = 1
'   exit function

'   New table on 16/01/08 Landlord
'###############################################################################################################
TableMissing_Landlord:
   On Error GoTo EH_TableMissing_Landlord

   Rst1.Open "SELECT * FROM Landlord;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADD_SupplierType_Supplier

EH_TableMissing_Landlord:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: Landlord"
   UpdateDatabase = -1
   Exit Function
'
''   Add 2 columns Usage & LeaseID on 30/01/08 UnitInsurance
''###############################################################################################################
'ADD_Usage_LeaseID_UnitInsurance:
'   On Error GoTo CHANGE_ADD_Usage_LeaseID_UnitInsurance
'
'   Rst1.Open "SELECT Usage FROM UnitInsurance;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
'
'   GoTo ADD_SupplierType_Supplier
'
'CHANGE_ADD_Usage_LeaseID_UnitInsurance:
'   Conn1.Execute "ALTER TABLE UnitInsurance ADD COLUMN Usage TEXT(100);"
'   Conn1.Execute "ALTER TABLE UnitInsurance ADD COLUMN LeaseID TEXT(20);"
'   UpdateDatabase = 1
'   Exit Function

'   Add a columns SupplierType on 30/01/08 Supplier
'###############################################################################################################
ADD_SupplierType_Supplier:
   On Error GoTo CHANGE_ADD_SupplierType_Supplier

   Rst1.Open "SELECT SupplierType FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_RentalPrice_Units

CHANGE_ADD_SupplierType_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN SupplierType TEXT(10);"
   UpdateDatabase = 1
   Exit Function
   
'   Add a columns RentalPrice on 31/01/08 Units
'###############################################################################################################
ADD_RentalPrice_Units:
   On Error GoTo CHANGE_ADD_RentalPrice_Units

   Rst1.Open "SELECT RentalPrice FROM Units;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_GroupNo_TenantDeposit

CHANGE_ADD_RentalPrice_Units:
   Conn1.Execute "ALTER TABLE Units ADD COLUMN RentalPrice TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns GroupNo on 19/02/08 TenantDeposit
'###############################################################################################################
ADD_GroupNo_TenantDeposit:
   On Error GoTo CHANGE_ADD_GroupNo_TenantDeposit

   Rst1.Open "SELECT GroupNo FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_UtilityUsage_LeaseDetails

CHANGE_ADD_GroupNo_TenantDeposit:
   Conn1.Execute "ALTER TABLE TenantDeposit ADD COLUMN GroupNo Number;"
   UpdateDatabase = 1
   Exit Function

'   Add a columns UtilityUsage on 19/02/08 LeaseDetails
'###############################################################################################################
ADD_UtilityUsage_LeaseDetails:
   On Error GoTo CHANGE_ADD_UtilityUsage_LeaseDetails

   Rst1.Open "SELECT UtilityUsage FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo TableMissing_LUtilityUsage

CHANGE_ADD_UtilityUsage_LeaseDetails:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN UtilityUsage TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   New table on 19/01/08 LUtilityUsage
'###############################################################################################################
TableMissing_LUtilityUsage:
   On Error GoTo EH_TableMissing_LUtilityUsage

   Rst1.Open "SELECT * FROM LUtilityUsage;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADDNEW_REC_UUSE

EH_TableMissing_LUtilityUsage:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: LUtilityUsage"
   UpdateDatabase = -1
   Exit Function

'   Add new record on 19/03/08 PrimaryCode
'###############################################################################################################
ADDNEW_REC_UUSE:
   On Error GoTo MissingRec_REC_UUSE

   Rst1.Open "SELECT Code FROM PrimaryCode WHERE Code = 'UUSE';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "UUSE"
         !Value = "UNIT USAGES"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADD_RentalPrice_UnitUsages

MissingRec_REC_UUSE:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function
   
'   Add a columns UnitUsages on 19/03/08 Units
'###############################################################################################################
ADD_RentalPrice_UnitUsages:
   On Error GoTo CHANGE_ADD_UnitUsages_Units

   Rst1.Open "SELECT UnitUsages FROM Units;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Unit_Usage

CHANGE_ADD_UnitUsages_Units:
   Conn1.Execute "ALTER TABLE Units ADD COLUMN UnitUsages TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Usage on 27/03/08 LeaseDetails
'###############################################################################################################
ADD_Unit_Usage:
   On Error GoTo CHANGE_ADD_Unit_Usage

   Rst1.Open "SELECT Usage FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_REMINDER_ID_ALARM_LEASEDETAILS

CHANGE_ADD_Unit_Usage:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN Usage TEXT(50);"
   UpdateDatabase = 1
   Exit Function
   
'   Add a column REMINDER_ID, ALARM on 28/03/08 LeaseDetails
'###############################################################################################################
ADD_REMINDER_ID_ALARM_LEASEDETAILS:
   On Error GoTo CHANGE_ADD_REMINDER_ID_ALARM_LEASEDETAILS

   Rst1.Open "SELECT REMINDER_ID, ALARM FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo tlbLetterReports

CHANGE_ADD_REMINDER_ID_ALARM_LEASEDETAILS:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN ALARM TEXT(1);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN REMINDER_ID TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   New table on 03/04/08 tlbLetterReports
'###############################################################################################################
tlbLetterReports:
   On Error GoTo MissingTable_tlbLetterReports

   Rst1.Open "SELECT * FROM tlbLetterReports;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADD_Records_PaymentDates

MissingTable_tlbLetterReports:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbLetterReports"
   UpdateDatabase = -1
   Exit Function

'   Add Records columns on 03/04/2008 PaymentDates
'###############################################################################################################
ADD_Records_PaymentDates:
   On Error GoTo MissingTable_ADD_NEW_COLUMNS

   Rst1.Open "SELECT * FROM GlobalData;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount > 0 Then
      Rst2.Open "SELECT * FROM PaymentDates;", Conn1, adOpenDynamic, adLockOptimistic
      If Rst2.RecordCount = 0 Then
         Rst2.AddNew
         Rst2!DateSetID = 0
         Rst2!NameOfSet = "Default"
         Rst2!MonthlyDueDate1 = Rst1!MonthlyDueDate1
         Rst2!MonthlyDueDate2 = Rst1!MonthlyDueDate2
         Rst2!MonthlyDueDate3 = Rst1!MonthlyDueDate3
         Rst2!MonthlyDueDate4 = Rst1!MonthlyDueDate4
         Rst2!MonthlyDueDate5 = Rst1!MonthlyDueDate5
         Rst2!MonthlyDueDate6 = Rst1!MonthlyDueDate6
         Rst2!MonthlyDueDate7 = Rst1!MonthlyDueDate7
         Rst2!MonthlyDueDate8 = Rst1!MonthlyDueDate8
         Rst2!MonthlyDueDate9 = Rst1!MonthlyDueDate9
         Rst2!MonthlyDueDate10 = Rst1!MonthlyDueDate10
         Rst2!MonthlyDueDate11 = Rst1!MonthlyDueDate11
         Rst2!MonthlyDueDate12 = Rst1!MonthlyDueDate12
         Rst2!QuarterlyDueDate1 = Rst1!QuarterlyDueDate1
         Rst2!QuarterlyDueDate2 = Rst1!QuarterlyDueDate2
         Rst2!QuarterlyDueDate3 = Rst1!QuarterlyDueDate3
         Rst2!QuarterlyDueDate4 = Rst1!QuarterlyDueDate4
         Rst2!HalfYearlyDueDate1 = Rst1!HalfYearlyDueDate1
         Rst2!HalfYearlyDueDate2 = Rst1!HalfYearlyDueDate2
         Rst2!YearlyDueDate = Rst1!YearlyDueDate
         Rst2.Update
      End If
      Rst2.Close
      Rst1.Close
   Else
      Rst1.Close
   End If

   GoTo ADD_COLUMNS_TO_SUPPLIER

MissingTable_ADD_Records_PaymentDates:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add columnS on 28/03/08 Supplier
'###############################################################################################################
ADD_COLUMNS_TO_SUPPLIER:
   On Error GoTo CHANGE_ADD_COLUMNS_TO_SUPPLIER

   Rst1.Open "SELECT CreditLimit,NominalCode, AccountType, PaymentType, PaymentTerms, SortCode, AcNo, AcName,BPR FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_REC_PAYT
   
CHANGE_ADD_COLUMNS_TO_SUPPLIER:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN CreditLimit Number;"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN nominalcode number;"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN Accounttype TEXT(10);"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN PaymentType TEXT(10);"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN PaymentTerms Number;"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN SortCode TEXT(8);"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN AcNo TEXT(8);"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN AcName TEXT(100);"
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN BPR TEXT(100);"
   
   UpdateDatabase = 1
   Exit Function

'   Add new record on 14/04/08 PrimaryCode
'###############################################################################################################
ADDNEW_REC_PAYT:
   On Error GoTo MissingRec_REC_PAYT

   Rst1.Open "SELECT Code FROM PrimaryCode WHERE Code = 'PAYT';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "PAYT"
         !Value = "PAYMENT TYPE"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADDNEW_REC_ACCT:

MissingRec_REC_PAYT:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new record on 14/04/08 PrimaryCode
'###############################################################################################################
ADDNEW_REC_ACCT:
   On Error GoTo MissingRec_REC_ACCT

   Rst1.Open "SELECT Code FROM PrimaryCode WHERE Code = 'ACCT';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "ACCT"
         !Value = "ACCOUNT TYPE"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADD_COLUMNS_TO_LRentCharges
MissingRec_REC_ACCT:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function
   
'   Add columns on 17/04/08 LRentCharges
'###############################################################################################################
ADD_COLUMNS_TO_LRentCharges:
   On Error GoTo CHANGE_ADD_COLUMNS_TO_LRentCharges

   Rst1.Open "SELECT StopByRR, RR_ID, PRR FROM LRentCharges;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo ADD_VatCode_Supplier

CHANGE_ADD_COLUMNS_TO_LRentCharges:
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN StopByRR TEXT(1);"
   Conn1.Execute "UPDATE LRentCharges SET StopByRR = 'N'"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN RR_ID LONG;"
   Conn1.Execute "ALTER TABLE LRentCharges ADD COLUMN PRR BYTE;"
   Conn1.Execute "UPDATE LRentCharges SET PRR = 0;"

   UpdateDatabase = 1
   Exit Function

'   Add a columns VatCode on 07/05/08 Supplier
'###############################################################################################################
ADD_VatCode_Supplier:
   On Error GoTo CHANGE_ADD_VatCode_Supplier

   Rst1.Open "SELECT VatCode FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_LRSD_tlbClientBanks

CHANGE_ADD_VatCode_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN VatCode TEXT(3);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns LRSD on 19/05/08 tlbClientBanks
'###############################################################################################################
ADD_LRSD_tlbClientBanks:
   On Error GoTo CHANGE_ADD_LRSD_tlbClientBanks

   Rst1.Open "SELECT LRSD FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_6Cols_tlbBank

CHANGE_ADD_LRSD_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN LRSD TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add 6 columns on 20/05/08 tlbBank
'###############################################################################################################
ADD_6Cols_tlbBank:
   On Error GoTo CHANGE_ADD_6Cols_tlbBank

   Rst1.Open "SELECT Contact FROM tlbBank;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Reconciled_RptTransactions

CHANGE_ADD_6Cols_tlbBank:
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN Contact TEXT(80);"
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN Tel TEXT(20);"
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN Fax TEXT(20);"
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN Mobile TEXT(20);"
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN eMail TEXT(50);"
   Conn1.Execute "ALTER TABLE tlbBank ADD COLUMN Website TEXT(80);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns Reconciled on 22/05/08 RptTransactions
'###############################################################################################################
ADD_Reconciled_RptTransactions:
   On Error GoTo CHANGE_ADD_Reconciled_RptTransactions

   Rst1.Open "SELECT Reconciled FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo RESIZE_TypeOfStore_LeaseDetails
   GoTo ADD_ReconNow_RptTransactions

CHANGE_ADD_Reconciled_RptTransactions:
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN Reconciled Currency;"
   UpdateDatabase = 1
   Exit Function

'   Add a columns ReconNow on 23/05/08 RptTransactions
'###############################################################################################################
ADD_ReconNow_RptTransactions:
   On Error GoTo CHANGE_ADD_ReconNow_RptTransactions

   Rst1.Open "SELECT ReconNow FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo RESIZE_TypeOfStore_LeaseDetails

CHANGE_ADD_ReconNow_RptTransactions:
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN ReconNow TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of TypeOfStore on 10/06/2008 LeaseDetails
'###############################################################################################################
RESIZE_TypeOfStore_LeaseDetails:

   On Error GoTo MissingTable_RESIZE_TypeOfStore_LeaseDetails

   Rst1.Open "SELECT TypeOfStore FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE LeaseDetails ALTER COLUMN TypeOfStore TEXT(50);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo UPADTE_VAL_InvoiceTo_Tenants

MissingTable_RESIZE_TypeOfStore_LeaseDetails:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "Col Size - Details of DSR"
   UpdateDatabase = 1
   Exit Function

'   Update Value InvoiceTo on 03/07/08 Tenants
'###############################################################################################################
UPADTE_VAL_InvoiceTo_Tenants:

   Conn1.Execute "UPDATE Tenants SET InvoiceTo = 'H' WHERE InvoiceTo = '' OR Isnull(InvoiceTo);"

   GoTo ADD_PLControl_Supplier

'   Add a columns PLControl on 14/07/08 Supplier
'###############################################################################################################
ADD_PLControl_Supplier:
   On Error GoTo CHANGE_ADD_PLControl_Supplier

   Rst1.Open "SELECT PLControl FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_REC_GID 'TableMissing_BankTransactions

CHANGE_ADD_PLControl_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN PLControl Long;"
   UpdateDatabase = 1
   Exit Function

''   Add a commands in the registry on 21/07/08
''###############################################################################################################
'ADD_REGISTRY_VALUE:
'   Dim szChoice As String, szaChoice() As String
'
''  Remember choice
'   szChoice = GetSetting("PropertyManagement", "ChoosedOption", "AutoID-c" & CStr(SCID))
'   If Len(szChoice) <= 0 Then
'      szChoice = "U#CL#L#S#P#MA"
'      SaveSetting "PropertyManagement", "ChoosedOption", "AutoID-c" & CStr(SCID), szChoice
'   End If
'
''   New table on 23/07/08 BankTransactions
''###############################################################################################################
'TableMissing_BankTransactions:
'   On Error GoTo EH_TableMissing_BankTransactions
'
'   Rst1.Open "SELECT * FROM BankTransactions;", Conn1, adOpenStatic, adLockReadOnly
'
'   Rst1.Close
''   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
'   GoTo ADDNEW_REC_GID
'
'EH_TableMissing_BankTransactions:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: BankTransactions"
'   UpdateDatabase = -1
'   Exit Function

'   Add new record GID on 31/07/08 PrimaryCode
'###############################################################################################################
ADDNEW_REC_GID:
   On Error GoTo MissingRec_REC_GID

   Rst1.Open "SELECT Code FROM PrimaryCode WHERE Code = 'GID';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PrimaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !Code = "GID"
         !Value = "GENERATE ID"
         !Flexible = False
         .Update
      End With
   End If
   Rst1.Close

   GoTo SECONDARYCODE_ADDNEW_REC_GID

MissingRec_REC_GID:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new value GID on 31/07/08 SecondaryCode
'###############################################################################################################
SECONDARYCODE_ADDNEW_REC_GID:
   Rst1.Open "SELECT * FROM SecondaryCode WHERE Code = 'GID' AND PrimaryCode = 'GID';", Conn1, adOpenDynamic, adLockOptimistic

   If Rst1.EOF Then
      Rst1.AddNew
      Rst1!PrimaryCode = "GID"
      Rst1!Code = "GID"
      Rst1!Value = "U#CL#L#S#P#MA"
      Rst1.Update
   End If
   
   Rst1.Close

'   Add a columns isPrint on 06/08/08 Units
'###############################################################################################################
ADD_isPrint_Units:
   On Error GoTo CHANGE_ADD_isPrint_Units

   Rst1.Open "SELECT isPrint FROM Units;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo TableMissing_tlbPayment

CHANGE_ADD_isPrint_Units:
   Conn1.Execute "ALTER TABLE Units ADD COLUMN isPrint TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   New table on 14/08/08 tlbPayment
'###############################################################################################################
TableMissing_tlbPayment:
   On Error GoTo EH_TableMissing_tlbPayment

   Rst1.Open "SELECT * FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo TableMissing_PayTransactions

EH_TableMissing_tlbPayment:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: tlbPayment, tlbPaymentSplit"
   UpdateDatabase = -1
   Exit Function

'   New table on 14/08/08 PayTransactions
'###############################################################################################################
TableMissing_PayTransactions:
   On Error GoTo EH_TableMissing_PayTransactions

   Rst1.Open "SELECT * FROM PayTransactions ;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo MODIFY_DT_ChargingMethod_tblPrevGLU

EH_TableMissing_PayTransactions:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Import Table: PayTransactions "
   UpdateDatabase = -1
   Exit Function

'   Change datatype of column 'ChargingMethod' on 23/09/08 tblPrevGLU
'###############################################################################################################
MODIFY_DT_ChargingMethod_tblPrevGLU:
   On Error GoTo CHANGE_MODIFY_DT_ChargingMethod_tblPrevGLU

   Rst1.Open "SELECT ChargingMethod FROM tblPrevGLU;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).Type = adVarWChar Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE tblPrevGLU ALTER COLUMN ChargingMethod TEXT(30);"
      GoTo DELETE_RS_DemandRecords_tlbReceipt
   End If

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo DELETE_RS_DemandRecords_tlbReceipt

CHANGE_MODIFY_DT_ChargingMethod_tblPrevGLU:
   UpdateDatabase = 1
   Exit Function

'   Delete relationship between DemandRecords and tlbReceipt on 23/09/08
'###############################################################################################################
DELETE_RS_DemandRecords_tlbReceipt:
   On Error Resume Next

   Conn1.Execute "ALTER TABLE tlbReceipt DROP CONSTRAINT DemandRecordstlbReceipt;"

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo MODIFY_CS_LeaseID_tblPrevGLU

'   Change ColumnSize of column 'LeaseID' on 23/09/08 tblPrevGLU
'###############################################################################################################
MODIFY_CS_LeaseID_tblPrevGLU:
   On Error GoTo CHANGE_MODIFY_CS_LeaseID_tblPrevGLU

   Rst1.Open "SELECT LeaseID FROM tblPrevGLU;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).DefinedSize = 25 Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "DELETE * FROM tblPrevGLU;"
      Conn1.Execute "DROP INDEX LeaseID ON tblPrevGLU;"
      Conn1.Execute "ALTER TABLE tblPrevGLU DROP COLUMN LeaseID;"
      Conn1.Execute "ALTER TABLE tblPrevGLU ADD COLUMN LeaseID TEXT(25);"
   End If

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo UPDATE_LRSI_STOP_DATE

CHANGE_MODIFY_CS_LeaseID_tblPrevGLU:
   UpdateDatabase = 1
   Exit Function

'   Set the emply data to NULL on 08/10/08
'###############################################################################################################
UPDATE_LRSI_STOP_DATE:
   On Error Resume Next

   Conn1.Execute "UPDATE LRentCharges SET StopRC = NULL WHERE StopRC='';"
   Conn1.Execute "UPDATE LServiceCharges SET StopSC = NULL WHERE StopSC='';"
   Conn1.Execute "UPDATE LInsuranceCharges SET StopIC = NULL WHERE StopIC='';"

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
'   GoTo ADD_PLControlName_Supplier

'   Add a columns PLControlName on 14/10/08 Supplier
'###############################################################################################################
ADD_PLControlName_Supplier:
   On Error GoTo CHANGE_ADD_PLControlName_Supplier

   Rst1.Open "SELECT PLControlName FROM Supplier;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo ADD_TransactionID_TenantDeposit

CHANGE_ADD_PLControlName_Supplier:
   Conn1.Execute "ALTER TABLE Supplier ADD COLUMN PLControlName TEXT(100);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns TransactionID on 22/10/08 TenantDeposit
'###############################################################################################################
ADD_TransactionID_TenantDeposit:
   On Error GoTo CHANGE_ADD_TransactionID_TenantDeposit

   Rst1.Open "SELECT TransactionID FROM TenantDeposit;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo ADD_Manager_Property

CHANGE_ADD_TransactionID_TenantDeposit:
   Conn1.Execute "ALTER TABLE TenantDeposit ADD COLUMN TransactionID TEXT(5);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns Manager on 10/12/08 Property
'###############################################################################################################
ADD_Manager_Property:
   On Error GoTo CHANGE_ADD_Manager_Property

   Rst1.Open "SELECT Manager FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo RENAME_SPARE12_CHARGINGFIGURE_tlbDRCurrentPrint

CHANGE_ADD_Manager_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN Manager TEXT(80);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN ContactDetails TEXT(200);"
   
   UpdateDatabase = 1
   Exit Function

'   Rename a column spare12-ChargingFigure on 23/12/08 tlbDRCurrentPrint
'###############################################################################################################
RENAME_SPARE12_CHARGINGFIGURE_tlbDRCurrentPrint:
   On Error GoTo CHANGE_RENAME_SPARE12_CHARGINGFIGURE_tlbDRCurrentPrint

   Rst1.Open "SELECT ChargingFigure FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo ADD_SlNumber_tlbReceipt

CHANGE_RENAME_SPARE12_CHARGINGFIGURE_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN ChargingFigure LONG;"
   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ChargingFigure LONG;"
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint DROP COLUMN spare12;"

   UpdateDatabase = 1
   Exit Function

'   Add a columns SlNumber on 15/01/09 tlbReceipt & RptTransactions
'###############################################################################################################
ADD_SlNumber_tlbReceipt:
   On Error GoTo CHANGE_ADD_SlNumber_tlbReceipt

   Rst1.Open "SELECT SlNumber FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo NEW_tblPurInv

CHANGE_ADD_SlNumber_tlbReceipt:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN SlNumber Long;"
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN SlNumber Long;"

   UpdateDatabase = 1
   Exit Function

'   New table on 20/01/09 tblPurInv
'###############################################################################################################
NEW_tblPurInv:
   On Error GoTo MissingTable_NEW_tblPurInv

   Rst1.Open "SELECT * FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo NEW_tblPurInvSRec

MissingTable_NEW_tblPurInv:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Add New Table - tblPurInv"
   UpdateDatabase = -1
   Exit Function

'   New table on 20/01/09 tblPurInvSRec
'###############################################################################################################
NEW_tblPurInvSRec:
   On Error GoTo MissingTable_NEW_tblPurInvSRec

   Rst1.Open "SELECT * FROM tblPurInvSRec;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo DELETE_RS_GlobalSC_GlobalSCDtls

MissingTable_NEW_tblPurInvSRec:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Add New Table - tblPurInvSRec"
   UpdateDatabase = -1
   Exit Function

'   Delete relationship between GlobalSC and GlobalSCDtls on 21/01/2009
'###############################################################################################################
DELETE_RS_GlobalSC_GlobalSCDtls:
   On Error Resume Next

   Conn1.Execute "ALTER TABLE GlobalSCDtls DROP CONSTRAINT GlobalSCGlobalSCDtls;"

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo RESIZE_UniqueID_3Tables

'   Extend the field size of TableID on 21/01/2009 GlobalRC, GlobalSC, GlobalSCDtls
'###############################################################################################################
RESIZE_UniqueID_3Tables:

   On Error GoTo EXTEND_FIELD_SIZE_3TABLES

   Rst1.Open "SELECT BudgetID FROM GlobalRC;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 25 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE GlobalRC ALTER COLUMN BudgetID TEXT(25);"
      Conn1.Execute "ALTER TABLE GlobalSCDtls ALTER COLUMN BudgetID TEXT(25);"
      Conn1.Execute "ALTER TABLE GlobalSC ALTER COLUMN BudgetID TEXT(25);"
      Conn1.Execute "ALTER TABLE GlobalSCDtls ALTER COLUMN BudgetDtlID TEXT(25);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo RESIZE_Reference_tlbDRCurrentPrint

EXTEND_FIELD_SIZE_3TABLES:
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of Reference on 21/01/2009 tlbDRCurrentPrint
'###############################################################################################################
RESIZE_Reference_tlbDRCurrentPrint:

   On Error GoTo EXTEND_FIELD_SIZE_tlbDRCurrentPrint

   Rst1.Open "SELECT Reference FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN Reference TEXT(50);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo TENANT_ACCOUNT_HISTORY

EXTEND_FIELD_SIZE_tlbDRCurrentPrint:
   UpdateDatabase = 1
   Exit Function

'   Tenant Account History Details fixed on 27/01/2009 tlbReceipt
'###############################################################################################################
TENANT_ACCOUNT_HISTORY:
   On Error GoTo ERROR_TENANT_ACCOUNT_HISTORY

   Rst1.Open "select * from demandrecords where isnull(details);", Conn1, adOpenStatic, adLockReadOnly

   If Not Rst1.EOF Then
      Conn1.Execute "UPDATE demandrecords AS dr, demandsplitrecords AS ds SET dr.details = ds.Description " & _
                    "WHERE dr.demandid=ds.demandid And isnull(dr.details) And ds.splitid=1;"

      Conn1.Execute "UPDATE demandrecords AS dr, tlbReceipt AS r SET r.details = dr.Details " & _
                    "WHERE r.demandref=dr.demandid And isnull(r.details);"
   End If

ERROR_TENANT_ACCOUNT_HISTORY:
   Rst1.Close
   Set Rst1 = Nothing

   GoTo ADDNEW_TRAN_TYPE_REFUND

'   Add new record on 27/01/2009 tlbTransactionTypes
'###############################################################################################################
ADDNEW_TRAN_TYPE_REFUND:
   Rst1.Open "SELECT * FROM tlbTransactionTypes WHERE CONSTANT = 'sdoSRR';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM tlbTransactionTypes;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !TYPE_ID = 23
         !description = "Sales Receipt Refund"
         !CONSTANT = "sdoSRR"
         !ACCOUNTS_LINE = "Prestige"
         .Update
         .AddNew
         !TYPE_ID = 24
         !description = "Purchase Payment Refund"
         !CONSTANT = "sdoPPR"
         !ACCOUNTS_LINE = "Prestige"
         .Update
      End With
   End If
   Rst1.Close

   GoTo ADD_PropertyID_tlbBankPayment

'   Add a columns PropertyID on 17/02/09 tlbBankPayment
'###############################################################################################################
ADD_PropertyID_tlbBankPayment:
   On Error GoTo CHANGE_ADD_PropertyID_tlbBankPayment

   Rst1.Open "SELECT PropertyID FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo MODIFY_ChargingFigure_tlbDRCurrentPrint

CHANGE_ADD_PropertyID_tlbBankPayment:
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN PropertyID TEXT(8);"
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN ClientID TEXT(10);"

   UpdateDatabase = 1
   Exit Function

MissingRec_ADDNEW_REFUND:
   UpdateDatabase = 1
   Exit Function

'   Modify DATA TYPE ChargingFigure on 18/02/09 tlbDRCurrentPrint
'###############################################################################################################
MODIFY_ChargingFigure_tlbDRCurrentPrint:
   On Error GoTo CHANGE_MODIFY_ChargingFigure_tlbDRCurrentPrint

   Rst1.Open "SELECT ChargingFigure FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN ChargingFigure CURRENCY;"
      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN ChargingFigure CURRENCY;"
   Else
      Rst1.Close
   End If

   GoTo ADD_RentAnalysis_RentAnalysis

CHANGE_MODIFY_ChargingFigure_tlbDRCurrentPrint:
   UpdateDatabase = 1
   Exit Function

'   Add a columns RRStatus on 27/02/09 RentAnalysis
'###############################################################################################################
ADD_RentAnalysis_RentAnalysis:
   On Error GoTo CHANGE_ADD_RentAnalysis_RentAnalysis

   Rst1.Open "SELECT RRStatus FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADD_Location_PropertyMaintHistory

CHANGE_ADD_RentAnalysis_RentAnalysis:
   Conn1.Execute "ALTER TABLE RentAnalysis ADD COLUMN RRStatus TEXT(1);"

   UpdateDatabase = 1
   Exit Function

'   Add a column Location on 06/03/09 PropertyMaintHistory
'###############################################################################################################
ADD_Location_PropertyMaintHistory:
   On Error GoTo CHANGE_ADD_Location_PropertyMaintHistory

   Rst1.Open "SELECT Location FROM PropertyMaintHistory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_LastDate_UnitSafety

CHANGE_ADD_Location_PropertyMaintHistory:
   Conn1.Execute "ALTER TABLE PropertyMaintHistory ADD COLUMN Location TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column LastDate on 17/03/09 UnitSafety
'###############################################################################################################
ADD_LastDate_UnitSafety:
   On Error GoTo CHANGE_ADD_LastDate_UnitSafety

   Rst1.Open "SELECT LastDate FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Comment_UnitSafety

CHANGE_ADD_LastDate_UnitSafety:
   Conn1.Execute "ALTER TABLE UnitSafety ADD COLUMN LastDate TEXT(20);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Comment on 17/03/09 UnitSafety
'###############################################################################################################
ADD_Comment_UnitSafety:
   On Error GoTo CHANGE_ADD_Comment_UnitSafety

   Rst1.Open "SELECT Comment FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Alarm_UnitSafety

CHANGE_ADD_Comment_UnitSafety:
   Conn1.Execute "ALTER TABLE UnitSafety ADD COLUMN Comment TEXT(254);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Alarm on 17/03/09 UnitSafety
'###############################################################################################################
ADD_Alarm_UnitSafety:
   On Error GoTo CHANGE_ADD_Alarm_UnitSafety

   Rst1.Open "SELECT Alarm FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo UPADTE_VAL_EXPnVAL_SecondaryCode

CHANGE_ADD_Alarm_UnitSafety:
   Conn1.Execute "ALTER TABLE UnitSafety ADD COLUMN Alarm TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Update Value EXP n VAL on 18/03/09 SecondaryCode
'###############################################################################################################
UPADTE_VAL_EXPnVAL_SecondaryCode:

   Rst1.Open "SELECT Code FROM SecondaryCode " & _
             "WHERE Code = 'PLN' AND SecondaryCode.Value = 'Planned';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Conn1.Execute "UPDATE SecondaryCode " & _
                    "SET Code = 'PLN', " & _
                    "SecondaryCode.Value = 'Planned' " & _
                    "WHERE PrimaryCode = 'SSTA' AND Code = 'EXP';"

      Conn1.Execute "UPDATE SecondaryCode " & _
                    "SET Code = 'UPN', " & _
                    "SecondaryCode.Value = 'Unplanned' " & _
                    "WHERE PrimaryCode = 'SSTA' AND Code = 'VAL';"

      Conn1.Execute "UPDATE PrimaryCode " & _
                    "SET Flexible = 0 " & _
                    "WHERE Code = 'SSTA';"
   Else
      Rst1.Close
   End If

   GoTo ADD_Attachment_UnitSafety

'   Add a column Attachment on 19/03/09 UnitSafety
'###############################################################################################################
ADD_Attachment_UnitSafety:
   On Error GoTo CHANGE_ADD_Attachment_UnitSafety

   Rst1.Open "SELECT Attachment FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo MODIFY_UnitSafetyID_UnitSafety

CHANGE_ADD_Attachment_UnitSafety:
   Conn1.Execute "ALTER TABLE UnitSafety ADD COLUMN Attachment TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Modify DATA TYPE UnitSafetyID on 19/03/09 UnitSafety
'###############################################################################################################
MODIFY_UnitSafetyID_UnitSafety:
   On Error GoTo CHANGE_MODIFY_UnitSafetyID_UnitSafety

   Rst1.Open "SELECT UnitSafetyID FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
      Conn1.Execute "ALTER TABLE UnitSafety ALTER COLUMN UnitSafetyID TEXT(30);"
   Else
      Rst1.Close
   End If

   GoTo RESIZE_OwnerID_AttachedFile

CHANGE_MODIFY_UnitSafetyID_UnitSafety:
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "CHANGE_MODIFY_COL_DSR"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of OwnerID on 19/03/2009 AttachedFile
'###############################################################################################################
RESIZE_OwnerID_AttachedFile:

   Rst1.Open "SELECT OwnerID FROM AttachedFile;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 30 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE AttachedFile ALTER COLUMN OwnerID TEXT(30);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   Add a column Module on 19/03/09 UnitSafety
'###############################################################################################################
ADD_Module_UnitSafety:
   On Error GoTo CHANGE_ADD_Module_UnitSafety

   Rst1.Open "SELECT Module FROM UnitSafety;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Occupier_UnitUtilities

CHANGE_ADD_Module_UnitSafety:
   Conn1.Execute "ALTER TABLE UnitSafety ADD COLUMN Module TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Occupier on 26/03/09 UnitUtilities
'###############################################################################################################
ADD_Occupier_UnitUtilities:
   On Error GoTo CHANGE_ADD_Occupier_UnitUtilities

   Rst1.Open "SELECT Occupier FROM UnitUtilities;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo MODIFY_UnitInsuranceID_UnitInsurance 'ADDNEW_REC_IRER

CHANGE_ADD_Occupier_UnitUtilities:
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN Occupier TEXT(10);"
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN Status TEXT(10);"
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN StartDate TEXT(20);"
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN InitialReading TEXT(20);"
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN Comments TEXT(200);"
   Conn1.Execute "ALTER TABLE UnitUtilities ADD COLUMN Module TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Modify DATA TYPE UnitInsuranceID on 27/03/09 UnitInsurance
'###############################################################################################################
MODIFY_UnitInsuranceID_UnitInsurance:
   On Error GoTo CHANGE_MODIFY_UnitInsuranceID_UnitInsurance

   Rst1.Open "SELECT UnitInsuranceID FROM UnitInsurance;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
      Conn1.Execute "ALTER TABLE UnitInsurance ALTER COLUMN UnitInsuranceID TEXT(30);"
   Else
      Rst1.Close
   End If

   GoTo ADDNEW_REC_IRER

CHANGE_MODIFY_UnitInsuranceID_UnitInsurance:
   UpdateDatabase = 1
   Exit Function

'   Add new record on 27/03/09 PrimaryCode
'###############################################################################################################
ADDNEW_REC_IRER:
   On Error GoTo MissingTable_ADDNEW_REC_IRER

   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'IRER';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "IRER"
      Rst1!Value = "INSURER"
      Rst1!Flexible = True
      Rst1.Update
   End If
   Rst1.Close

   GoTo ADD_CompReg_Client

MissingTable_ADDNEW_REC_IRER:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add a columns CompReg on 27/03/09 Client
'###############################################################################################################
ADD_CompReg_Client:
   On Error GoTo CHANGE_ADD_CompReg_Client

   Rst1.Open "SELECT CompReg FROM Client;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Occupier_PropertyUtilities

CHANGE_ADD_CompReg_Client:
   Conn1.Execute "ALTER TABLE Client ADD COLUMN CompReg TEXT(30);"
   Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd1 TEXT(50);"
   Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd2 TEXT(50);"
   Conn1.Execute "ALTER TABLE Client ADD COLUMN RegAdd3 TEXT(50);"
   Conn1.Execute "ALTER TABLE Client ADD COLUMN RegPostCode TEXT(10);"
   Conn1.Execute "ALTER TABLE Landlord ADD COLUMN CompReg TEXT(30);"
   Conn1.Execute "ALTER TABLE Landlord ADD COLUMN RegAdd1 TEXT(50);"
   Conn1.Execute "ALTER TABLE Landlord ADD COLUMN RegAdd2 TEXT(50);"
   Conn1.Execute "ALTER TABLE Landlord ADD COLUMN RegAdd3 TEXT(50);"
   Conn1.Execute "ALTER TABLE Landlord ADD COLUMN RegPostCode TEXT(10);"
   UpdateDatabase = 1
   Exit Function

'   Add a column Occupier on 26/03/09 PropertyUtilities
'###############################################################################################################
ADD_Occupier_PropertyUtilities:
   On Error GoTo CHANGE_ADD_Occupier_PropertyUtilities

   Rst1.Open "SELECT Occupier FROM PropertyUtilities;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo MODIFY_PropertyInsuranceID_PropertyInsurance

CHANGE_ADD_Occupier_PropertyUtilities:
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN Occupier TEXT(10);"
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN Status TEXT(10);"
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN StartDate TEXT(20);"
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN InitialReading TEXT(20);"
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN Comments TEXT(200);"
   Conn1.Execute "ALTER TABLE PropertyUtilities ADD COLUMN Module TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Modify DATA TYPE PropertyInsuranceID on 01/04/09 PropertyInsurance
'###############################################################################################################
MODIFY_PropertyInsuranceID_PropertyInsurance:
   On Error GoTo CHANGE_MODIFY_PropertyInsuranceID_PropertyInsurance

   Rst1.Open "SELECT PropertyInsuranceID FROM PropertyInsurance;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
      Conn1.Execute "ALTER TABLE PropertyInsurance ALTER COLUMN PropertyInsuranceID TEXT(30);"
      Conn1.Execute "ALTER TABLE PropertyInsurance ADD COLUMN StartDate TEXT(20);"
      Conn1.Execute "ALTER TABLE PropertyInsurance ADD COLUMN Comments TEXT(200);"
      Conn1.Execute "ALTER TABLE PropertyInsurance ADD COLUMN Attachment TEXT(1);"
      Conn1.Execute "ALTER TABLE PropertyInsurance ADD COLUMN Module TEXT(1);"
      Conn1.Execute "ALTER TABLE PropertyInsurance ADD COLUMN Usage TEXT(20);"
   Else
      Rst1.Close
   End If

   GoTo ADDNEW_REC_USTA

CHANGE_MODIFY_PropertyInsuranceID_PropertyInsurance:
   UpdateDatabase = 1
   Exit Function

'   Add new record on 07/04/09 PrimaryCode
'###############################################################################################################
ADDNEW_REC_USTA:
   On Error GoTo MissingRec_ADDNEW_REC_USTA

   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'USTA';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "USTA"
      Rst1!Value = "UTILITY STATUS"
      Rst1!Flexible = True
      Rst1.Update
   End If
   Rst1.Close

   GoTo ADD_SlNumber_PayTransactions

MissingRec_ADDNEW_REC_USTA:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - USTA) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add a column SlNumber on 15/04/09 PayTransactions
'###############################################################################################################
ADD_SlNumber_PayTransactions:
   On Error GoTo CHANGE_ADD_SlNumber_PayTransactions

   Rst1.Open "SELECT SlNumber FROM PayTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_SlNumber_tlbPayment

CHANGE_ADD_SlNumber_PayTransactions:
   Conn1.Execute "ALTER TABLE PayTransactions ADD COLUMN SlNumber Long;"
   UpdateDatabase = 1
   Exit Function

'   Add a column SlNumber on 15/04/09 tlbPayment
'###############################################################################################################
ADD_SlNumber_tlbPayment:
   On Error GoTo CHANGE_ADD_SlNumber_tlbPayment

   Rst1.Open "SELECT SlNumber FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Prn_tblPurInv

CHANGE_ADD_SlNumber_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN SlNumber Long;"
   UpdateDatabase = 1
   Exit Function

'   Add a column Prn on 06/05/09 tblPurInv
'###############################################################################################################
ADD_Prn_tblPurInv:
   On Error GoTo CHANGE_ADD_Prn_tblPurInv

   Rst1.Open "SELECT Prn FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

'SKIP SOME UPDATE
'   GoTo ADD_ClosingBal_tlbClientBanks
   GoTo RESIZE_ReconNow_RptTransactions

CHANGE_ADD_Prn_tblPurInv:
   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN Prn TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of ReconNow on 12/05/2009 RptTransactions
'###############################################################################################################
RESIZE_ReconNow_RptTransactions:

   On Error GoTo EXTEND_RESIZE_ReconNow_RptTransactions

   Rst1.Open "SELECT ReconNow FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE RptTransactions ALTER COLUMN ReconNow TEXT(50);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo RESIZE_ReconNow_PayTransactions

EXTEND_RESIZE_ReconNow_RptTransactions:
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN ReconNow TEXT(1);"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of ReconNow on 12/05/2009 PayTransactions
'###############################################################################################################
RESIZE_ReconNow_PayTransactions:

   On Error GoTo EXTEND_RESIZE_ReconNow_PayTransactions

   Rst1.Open "SELECT ReconNow FROM PayTransactions;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE PayTransactions ALTER COLUMN ReconNow TEXT(50);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   szDataBaseUpdateStatus = "Update has been done successfully." & Chr(13) & "     DATABASE IS UPTO DATE"
   GoTo ADD_ClosingBal_tlbClientBanks

EXTEND_RESIZE_ReconNow_PayTransactions:
   UpdateDatabase = 1
   Exit Function

'   Add a columns ClosingBal on 12/05/09 tlbClientBanks
'###############################################################################################################
ADD_ClosingBal_tlbClientBanks:
   On Error GoTo CHANGE_ADD_ClosingBal_tlbClientBanks

   Rst1.Open "SELECT ClosingBal FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Indentifier_tlbClientBanks

CHANGE_ADD_ClosingBal_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN ClosingBal CURRENCY;"
   UpdateDatabase = 1
   Exit Function

'   Add a columns Indentifier on 08/06/09 tlbClientBanks
'###############################################################################################################
ADD_Indentifier_tlbClientBanks:
   On Error GoTo CHANGE_ADD_Indentifier_tlbClientBanks

   Rst1.Open "SELECT Indentifier FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_xUNIT_ID_PropertyID_tblPurInv

CHANGE_ADD_Indentifier_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN Indentifier TEXT(50);"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN FileExten TEXT(8);"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN FileLoc TEXT(255);"
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN EB TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a column xUNIT_ID-PropertyID on 16/06/09 tblPurInv
'###############################################################################################################
ADD_xUNIT_ID_PropertyID_tblPurInv:
   On Error GoTo CHANGE_ADD_xUNIT_ID_PropertyID_tblPurInv

   Rst1.Open "SELECT PropertyID FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_PropertyID_DemandTypes

CHANGE_ADD_xUNIT_ID_PropertyID_tblPurInv:
   Conn1.Execute "ALTER TABLE tblPurInv ADD COLUMN PropertyID TEXT(4);"

   UpdateDatabase = 1
   Exit Function

'   Add a columns PropertyID on 06/08/09 DemandTypes
'###############################################################################################################
ADD_PropertyID_DemandTypes:
   On Error GoTo CHANGE_ADD_PropertyID_DemandTypes

   Rst1.Open "SELECT PropertyID FROM DemandTypes;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_NominalCode_tlbClientBanks

CHANGE_ADD_PropertyID_DemandTypes:
   Conn1.Execute "ALTER TABLE DemandTypes ADD COLUMN PropertyID TEXT(4);"
   Conn1.Execute "UPDATE DemandTypes SET PropertyID = 'ALL';"
   UpdateDatabase = 1
   Exit Function

'   Add a columns NominalCode on 10/08/09 tlbClientBanks
'###############################################################################################################
ADD_NominalCode_tlbClientBanks:
   On Error GoTo CHANGE_ADD_NominalCode_tlbClientBanks

   Rst1.Open "SELECT NominalCode FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_ChqNo_tlbPayment

CHANGE_ADD_NominalCode_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN NominalCode TEXT(18);"
   UpdateDatabase = 1
   Exit Function

'   Add a column ChqNo on 18/09/09 tlbPayment
'###############################################################################################################
ADD_ChqNo_tlbPayment:
   On Error GoTo CHANGE_ADD_ChqNo_tlbPayment

   Rst1.Open "SELECT ChqNo FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo NEW_NominalLedger

CHANGE_ADD_ChqNo_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN ChqNo TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   New table on 20/01/XX NominalLedger
'###############################################################################################################
NEW_NominalLedger:
   On Error GoTo MissingTable_NEW_NominalLedger

   Rst1.Open "SELECT * FROM NominalLedger;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo NEW_NLCategory

MissingTable_NEW_NominalLedger:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Add New Table - NominalLedger"
   UpdateDatabase = -1
   Exit Function

'   New table on 20/01/XX NLCategory
'###############################################################################################################
NEW_NLCategory:
   On Error GoTo MissingTable_NEW_NLCategory

   Rst1.Open "SELECT * FROM NLCategory;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo MODIFY_DT_NL_EW

MissingTable_NEW_NLCategory:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Add New Table - NLCategory"
   UpdateDatabase = -1
   Exit Function

'   Modify DATA TYPE Nominal Ledger on 01/10/09 Every Where
'###############################################################################################################
MODIFY_DT_NL_EW:
   On Error GoTo CHANGE_MODIFY_DT_NL_EW

'  THESE CHANGES HAVE NOT BEEN DONE IN BLANK DATABASE.

   Rst1.Open "SELECT FeeNCAmt FROM ChargeTypes;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields(0).DefinedSize = 8 Then
      Rst1.Close
'ChargeTypes
      Conn1.Execute "ALTER TABLE ChargeTypes ALTER COLUMN FeeNCAmt   TEXT(15);"
      Conn1.Execute "ALTER TABLE ChargeTypes ALTER COLUMN FeeNCVat   TEXT(15);"
      Conn1.Execute "ALTER TABLE ChargeTypes ALTER COLUMN FeeNCTotal TEXT(15);"
'DemandSplitRecords
      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN NominalCodeforAmount TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN NominalCodeforVAT TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN NominalCodeforTotal TEXT(15);"
'DemandSplPreview
      Conn1.Execute "ALTER TABLE DemandSplPreview ALTER COLUMN NominalCodeforAmount TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandSplPreview ALTER COLUMN NominalCodeforVAT TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandSplPreview ALTER COLUMN NominalCodeforTotal TEXT(15);"
'DemandTypes
      Conn1.Execute "ALTER TABLE DemandTypes ALTER COLUMN NominalCodeforAmount TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandTypes ALTER COLUMN NominalCodeforVAT TEXT(15);"
      Conn1.Execute "ALTER TABLE DemandTypes ALTER COLUMN NominalCodeforTotal TEXT(15);"
'PayableTypes
      Conn1.Execute "ALTER TABLE PayableTypes ALTER COLUMN PayNCAmt TEXT(15);"
      Conn1.Execute "ALTER TABLE PayableTypes ALTER COLUMN PayNCVat TEXT(15);"
      Conn1.Execute "ALTER TABLE PayableTypes ALTER COLUMN PayNCTotal TEXT(15);"
'PayTransactions
      Conn1.Execute "ALTER TABLE PayTransactions ALTER COLUMN NominalCode TEXT(15);"
'RptTransactions
      Conn1.Execute "ALTER TABLE RptTransactions ALTER COLUMN NominalCode TEXT(15);"
'Supplier
      Conn1.Execute "ALTER TABLE Supplier ALTER COLUMN NominalCode TEXT(15);"
'tblPoA
      Conn1.Execute "ALTER TABLE tblPoA ALTER COLUMN BankCode TEXT(15);"
      Conn1.Execute "ALTER TABLE tblPoA ALTER COLUMN NominalCode TEXT(15);"
'tlbChildDemandRecord
      Conn1.Execute "ALTER TABLE tlbChildDemandRecord ALTER COLUMN NominalCodeforAmount TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbChildDemandRecord ALTER COLUMN NominalCodeforVAT TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbChildDemandRecord ALTER COLUMN NominalCodeforTotal TEXT(15);"
'tlbClientBanks
'      Conn1.Execute "ALTER TABLE tlbClientBanks DROP CONSTRAINT NominalLedgertlbClientBanks;"
      Conn1.Execute "ALTER TABLE tlbClientBanks ALTER COLUMN NominalCode TEXT(15);"
'tlbDRCurrentPrint
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN NominalCodeforAmount TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN NominalCodeforVAT TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ALTER COLUMN NominalCodeforTotal TEXT(15);"
'tlbPayment
      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN BankCode TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbPayment ALTER COLUMN NominalCode TEXT(15);"
'tlbReceipt
      Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN BankCode TEXT(15);"
      Conn1.Execute "ALTER TABLE tlbReceipt ALTER COLUMN NominalCode TEXT(15);"
'NLCategory
'      Conn1.Execute "ALTER TABLE NominalLedger DROP CONSTRAINT NLCategoryNominalLedger;"
      Conn1.Execute "ALTER TABLE NLCategory ALTER COLUMN CategoryCode TEXT(8);"
'NominalLedger
      Conn1.Execute "ALTER TABLE NominalLedger ALTER COLUMN CategoryCode TEXT(8);"
      Conn1.Execute "ALTER TABLE NominalLedger ALTER COLUMN Code TEXT(15);"
   Else
      Rst1.Close
   End If

   GoTo ADD_SptSlNo_tlbDRCurrentPrint

CHANGE_MODIFY_DT_NL_EW:
MsgBox Err.Number & " : " & Err.description
   UpdateDatabase = 1
   Exit Function

'   Add a column SptSlNo on 04/11/09 tlbDRCurrentPrint
'###############################################################################################################
ADD_SptSlNo_tlbDRCurrentPrint:
   On Error GoTo CHANGE_ADD_SptSlNo_tlbDRCurrentPrint

   Rst1.Open "SELECT SptSlNo FROM tlbDRCurrentPrint;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_AdiComment_DemandSplitRecords

CHANGE_ADD_SptSlNo_tlbDRCurrentPrint:
   Conn1.Execute "ALTER TABLE tlbDRCurrentPrint ADD COLUMN SptSlNo SHORT;"
   UpdateDatabase = 1
   Exit Function

'   Add a column AdiComment on 06/11/09 DemandSplitRecords
'###############################################################################################################
ADD_AdiComment_DemandSplitRecords:
   On Error GoTo CHANGE_ADD_AdiComment_DemandSplitRecords

   Rst1.Open "SELECT AdiComment FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_DmdSlNo_DemandRecords

CHANGE_ADD_AdiComment_DemandSplitRecords:
   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN AdiComment MEMO;"
   UpdateDatabase = 1
   Exit Function

'   Add a column DmdSlNo on 06/11/09 DemandRecords
'###############################################################################################################
ADD_DmdSlNo_DemandRecords:
   On Error GoTo CHANGE_ADD_DmdSlNo_DemandRecords

   Rst1.Open "SELECT DmdSlNo FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_PPSF_GlobalInsurance

CHANGE_ADD_DmdSlNo_DemandRecords:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN DmdSlNo LONG;"

   szSQL = "SELECT DmdSlNo FROM DemandRecords WHERE TransactionType = 1 ORDER BY DemandID;"
   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
   i = 1
   
   While Not Rst1.EOF
      Rst1.Fields.Item(0).Value = i
      i = i + 1
      Rst1.Update
      Rst1.MoveNext
   Wend
   Rst1.Close

   szSQL = "SELECT DmdSlNo FROM DemandRecords WHERE TransactionType = 2 ORDER BY DemandID;"
   Rst1.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic
   i = 1
   
   While Not Rst1.EOF
      Rst1.Fields.Item(0).Value = i
      i = i + 1
      Rst1.Update
      Rst1.MoveNext
   Wend
   Rst1.Close

   UpdateDatabase = 1
   Exit Function

'   Add TWO columnS PPSF & SCArea on 12/11/09 GlobalInsurance
'###############################################################################################################
ADD_PPSF_GlobalInsurance:
   On Error GoTo CHANGE_ADD_PPSF_GlobalInsurance

   Rst1.Open "SELECT PPSF FROM GlobalInsurance;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo RESIZE_BankCode_RptTransactions

CHANGE_ADD_PPSF_GlobalInsurance:
   Conn1.Execute "ALTER TABLE GlobalInsurance ADD COLUMN PPSF CURRENCY;"
   Conn1.Execute "ALTER TABLE GlobalInsurance ADD COLUMN SCArea SINGLE;"

   UpdateDatabase = 1
   Exit Function

'   Extend the field size of BankCode on 30/11/2009 RptTransactions
'###############################################################################################################
RESIZE_BankCode_RptTransactions:

   On Error GoTo EXTEND_RESIZE_BankCode_RptTransactions

   Rst1.Open "SELECT BankCode FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 15 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE RptTransactions ALTER COLUMN BankCode TEXT(15);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo ADDNEW_COL_ChargingMethod_DemandSplitRecords

EXTEND_RESIZE_BankCode_RptTransactions:
   UpdateDatabase = 1
   Exit Function

'   Add new column ChargingMethod on 01/11/09 DemandSplitRecords
'###############################################################################################################
ADDNEW_COL_ChargingMethod_DemandSplitRecords:
   On Error GoTo MISSING_ADDNEW_COL_ChargingMethod_DemandSplitRecords

   Rst1.Open "SELECT ChargingMethod FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_StPath_GlobalData

MISSING_ADDNEW_COL_ChargingMethod_DemandSplitRecords:
   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN ChargingMethod BYTE;"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new column StPath on 01/11/09 GlobalData
'###############################################################################################################
ADDNEW_COL_StPath_GlobalData:
   On Error GoTo MISSING_ADDNEW_COL_StPath_GlobalData

   Rst1.Open "SELECT StPath FROM GlobalData;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_BankMemo_tlbClientBanks

MISSING_ADDNEW_COL_StPath_GlobalData:
   Conn1.Execute "ALTER TABLE GlobalData ADD COLUMN StPath TEXT(254);"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add new column BankMemo on 08/12/09 tlbClientBanks
'###############################################################################################################
ADDNEW_COL_BankMemo_tlbClientBanks:
   On Error GoTo MISSING_ADDNEW_COL_BankMemo_tlbClientBanks

   Rst1.Open "SELECT BankMemo FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_SuppText1_LeaseDetails

MISSING_ADDNEW_COL_BankMemo_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN BankMemo MEMO;"
'   MsgBox "This company database has been updated. Please restart the program.", vbInformation + vbOKOnly, "ADDNEW_COL_DETAILS_DR"
   UpdateDatabase = 1
   Exit Function

'   Add few columns SuppText1 on 15/12/09 LeaseDetails
'###############################################################################################################
ADDNEW_COL_SuppText1_LeaseDetails:
   On Error GoTo MISSING_ADDNEW_COL_SuppText1_LeaseDetails

   Rst1.Open "SELECT SuppText1 FROM LeaseDetails;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_SuppText1_Property

MISSING_ADDNEW_COL_SuppText1_LeaseDetails:
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN SuppText1 TEXT(254);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN SuppText2 TEXT(254);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN SuppText3 TEXT(254);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN DateFlagDt2 TEXT(21);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN DateFlagDt3 TEXT(21);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN DateFlagDescription2 TEXT(50);"
   Conn1.Execute "ALTER TABLE LeaseDetails ADD COLUMN DateFlagDescription3 TEXT(50);"
   
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare9;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare10;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare11;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare12;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare13;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare14;"
   Conn1.Execute "ALTER TABLE LeaseDetails DROP COLUMN spare15;"

   UpdateDatabase = 1
   Exit Function

'   Add few columns SuppText1 on 18/12/09 Property
'###############################################################################################################
ADDNEW_COL_SuppText1_Property:
   On Error GoTo MISSING_ADDNEW_COL_SuppText1_Property

   Rst1.Open "SELECT SuppText1 FROM Property;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADDNEW_COL_SuppText1

MISSING_ADDNEW_COL_SuppText1_Property:
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppCaption1 TEXT(50);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppCaption2 TEXT(50);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppCaption3 TEXT(50);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppText1 TEXT(254);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppText2 TEXT(254);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN SuppText3 TEXT(254);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDt1 TEXT(21);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDt2 TEXT(21);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDt3 TEXT(21);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDescription1 TEXT(50);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDescription2 TEXT(50);"
   Conn1.Execute "ALTER TABLE Property ADD COLUMN DateFlagDescription3 TEXT(50);"
   
   UpdateDatabase = 1
   Exit Function

'   Add new column SuppText1 on 15/08/07 DemandRecords
'###############################################################################################################
ADDNEW_COL_SuppText1:
   On Error GoTo MissingTable_ADDNEW_COL_SuppText1

   Rst1.Open "SELECT SuppText1 FROM DemandRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Reconciled_tlbBankPayment

MissingTable_ADDNEW_COL_SuppText1:
   Conn1.Execute "ALTER TABLE DemandRecords ADD COLUMN SuppText1 TEXT(254)"
   Conn1.Execute "ALTER TABLE DemandRecords DROP COLUMN spare6;"
   
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Col - RptAmtType) - DemandRecords"
   UpdateDatabase = 1
   Exit Function

'   Add a columns Reconciled on 19/01/10 tlbBankPayment
'###############################################################################################################
ADD_Reconciled_tlbBankPayment:
   On Error GoTo CHANGE_ADD_Reconciled_tlbBankPayment

   Rst1.Open "SELECT Reconciled FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Reconciled_tlbReceipt

CHANGE_ADD_Reconciled_tlbBankPayment:
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN Reconciled Currency;"
   Conn1.Execute "ALTER TABLE tlbBankPayment ADD COLUMN ReconNow TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add columns Reconciled&ReconNow on 20/10/10 tlbReceipt
'###############################################################################################################
ADD_Reconciled_tlbReceipt:
   On Error GoTo CHANGE_ADD_Reconciled_tlbReceipt

   Rst1.Open "SELECT Reconciled FROM tlbReceipt;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Reconciled_tlbPayment

CHANGE_ADD_Reconciled_tlbReceipt:
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN Reconciled Currency;"
   Conn1.Execute "ALTER TABLE tlbReceipt ADD COLUMN ReconNow TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add columns Reconciled&ReconNow on 20/10/10 tlbPayment
'###############################################################################################################
ADD_Reconciled_tlbPayment:
   On Error GoTo CHANGE_ADD_Reconciled_tlbPayment

   Rst1.Open "SELECT Reconciled FROM tlbPayment;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_Exp2Sage_RptTransactions

CHANGE_ADD_Reconciled_tlbPayment:
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN Reconciled Currency;"
   Conn1.Execute "ALTER TABLE tlbPayment ADD COLUMN ReconNow TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns Exp2Sage on 22/01/2010 RptTransactions
'###############################################################################################################
ADD_Exp2Sage_RptTransactions:
   On Error GoTo CHANGE_ADD_Exp2Sage_RptTransactions

   Rst1.Open "SELECT Exp2Sage FROM RptTransactions;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo NEW_tblBatchTransaction

CHANGE_ADD_Exp2Sage_RptTransactions:
   Conn1.Execute "ALTER TABLE RptTransactions ADD COLUMN Exp2Sage TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   New table on 20/01/XX tblBatchTransaction
'###############################################################################################################
NEW_tblBatchTransaction:
   On Error GoTo MissingTable_NEW_tblBatchTransaction

   Rst1.Open "SELECT * FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   GoTo ADD_Fund_tblBatchTransaction

MissingTable_NEW_tblBatchTransaction:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Add New Table - tblBatchTransaction"
   UpdateDatabase = -1
   Exit Function

'   Add a columns Fund on 08/02/2010 tblBatchTransaction
'###############################################################################################################
ADD_Fund_tblBatchTransaction:

   On Error GoTo RESIZE_MY_ID_tlbBankPayment
   Rst1.Open "SELECT * FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   On Error GoTo CHANGE_ADD_Fund_tblBatchTransaction

   Rst1.Open "SELECT Fund FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo RESIZE_MY_ID_tlbBankPayment

CHANGE_ADD_Fund_tblBatchTransaction:
   Conn1.Execute "ALTER TABLE tblBatchTransaction ADD COLUMN Fund TEXT(50);"
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of MY_ID on 15/02/2010 tlbBankPayment
'###############################################################################################################
RESIZE_MY_ID_tlbBankPayment:

   On Error GoTo MissingTable_RESIZE_MY_ID_tlbBankPayment

   Rst1.Open "SELECT MY_ID FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN MY_ID TEXT(50);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo RESIZE_PROJ_REF_tlbBankPayment

MissingTable_RESIZE_MY_ID_tlbBankPayment:
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of PROJ_REF on 17/02/2010 tlbBankPayment
'###############################################################################################################
RESIZE_PROJ_REF_tlbBankPayment:

   On Error GoTo MissingTable_RESIZE_PROJ_REF_tlbBankPayment

   Rst1.Open "SELECT PROJ_REF FROM tlbBankPayment;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 20 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE tlbBankPayment ALTER COLUMN PROJ_REF TEXT(20);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo RESIZE_INV_NO_tblPurInv

MissingTable_RESIZE_PROJ_REF_tlbBankPayment:
   UpdateDatabase = 1
   Exit Function

'   Extend the field size of INV_NO on 17/02/2010 tblPurInv
'###############################################################################################################
RESIZE_INV_NO_tblPurInv:

   On Error GoTo MissingTable_RESIZE_INV_NO_tblPurInv

   Rst1.Open "SELECT INV_NO FROM tblPurInv;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item(0).DefinedSize <> 20 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE tblPurInv ALTER COLUMN INV_NO TEXT(20);"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

   GoTo MODIFY_DT_SageDepartment_DemandSplitRecords

MissingTable_RESIZE_INV_NO_tblPurInv:
   UpdateDatabase = 1
   Exit Function
   
'   Change datatype of SageDepartment on 18/02/10 DemandSplitRecords
'###############################################################################################################
MODIFY_DT_SageDepartment_DemandSplitRecords:
   On Error GoTo CHANGE_MODIFY_DT_SageDepartment_DemandSplitRecords

   Rst1.Open "SELECT SageDepartment FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.Fields(0).Type = adInteger Then
      Rst1.Close
   Else
      Rst1.Close
      Conn1.Execute "ALTER TABLE DemandSplitRecords ALTER COLUMN SageDepartment LONG;"
      GoTo CHANGE_MODIFY_DT_SageDepartment_DemandSplitRecords
   End If

   GoTo ADD_REC_SECONDARYCODE_RAT_DATA

CHANGE_MODIFY_DT_SageDepartment_DemandSplitRecords:
   UpdateDatabase = 1
   Exit Function

'   Change datatype of SageDepartment on 23/02/10 DemandSplitRecords
'###############################################################################################################
ADD_REC_SECONDARYCODE_RAT_DATA:
   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'RAT';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "RAT"
         !Code = "BACS"
         !Value = "BACS"
         .Update
         'Secondary code to keep track of Alarm Y/N
         .AddNew
         !PrimaryCode = "RAT"
         !Code = "CHQ"
         !Value = "Cheque"
         .Update
      End With
   End If
   Rst1.Close

'   Extend the field size of Field7 on 01/07/10 ShoppingCentre
'###############################################################################################################
   Rst1.Open "SELECT Field7 FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.Fields.Item("Field7").DefinedSize = 50 Then
      Rst1.Close
      Set Rst1 = Nothing

      Conn1.Execute "ALTER TABLE ShoppingCentre ALTER COLUMN Field7 TEXT(255)"
   Else
      Rst1.Close
      Set Rst1 = Nothing
   End If

'   Add a columns CurrentBalance on 22/01/2010 tlbClientBanks
'###############################################################################################################
ADD_CurrentBalance_tlbClientBanks:
   On Error GoTo CHANGE_ADD_CurrentBalance_tlbClientBanks

   Rst1.Open "SELECT CurrentBalance FROM tlbClientBanks;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo Check_Bank_Balance

CHANGE_ADD_CurrentBalance_tlbClientBanks:
   Conn1.Execute "ALTER TABLE tlbClientBanks ADD COLUMN CurrentBalance Currency;"
   UpdateDatabase = 1
   Exit Function

'   Check Bank account Balance on 23/02/10 DemandSplitRecords
'###############################################################################################################
Check_Bank_Balance:
   CheckBankBalance

''   Add new records on 21/04/10 NLType
''###############################################################################################################
'ADDNEW_REC_NLT:
'
'   Rst1.Open "SELECT NLTypeCode FROM NLType WHERE NLTypeCode = 1;", Conn1, adOpenStatic, adLockReadOnly
'
'   If Rst1.EOF Then
'      Rst1.Close
'      Rst1.Open "SELECT * FROM NLType;", Conn1, adOpenDynamic, adLockOptimistic
'      Rst1.AddNew
'      Rst1!NLTypeCode = 1
'      Rst1!TypeValue = "Balance Sheet"
'      Rst1.Update
'      Rst1.AddNew
'      Rst1!NLTypeCode = 2
'      Rst1!TypeValue = "Profit and Loss"
'      Rst1.Update
'   End If
'   Rst1.Close

'   Add new record on 21/04/10 PrimaryCode
'###############################################################################################################
   On Error GoTo MissingTable_ADDNEW_REC_NCDC

   Rst1.Open "SELECT CODE FROM PRIMARYCODE WHERE CODE = 'NCDC';", Conn1, adOpenStatic, adLockReadOnly

   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM PRIMARYCODE;", Conn1, adOpenDynamic, adLockOptimistic
      Rst1.AddNew
      Rst1!Code = "NCDC"
      Rst1!Value = "NOMINAL CODE POS"
      Rst1!Flexible = False
      Rst1.Update
   End If
   Rst1.Close

   GoTo ADD_REC_SECONDARYCODE_NCDC_DATA

MissingTable_ADDNEW_REC_NCDC:
'   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database (Add Record - RAT) - tlbReceipt"
   UpdateDatabase = 1
   Exit Function

'   Add new records on 21/04/10 SecondaryCode
'###############################################################################################################
ADD_REC_SECONDARYCODE_NCDC_DATA:
   Rst1.Open "SELECT PrimaryCode FROM SecondaryCode WHERE PrimaryCode = 'NCDC';", Conn1, adOpenStatic, adLockReadOnly
   If Rst1.EOF Then
      Rst1.Close
      Rst1.Open "SELECT * FROM SecondaryCode;", Conn1, adOpenDynamic, adLockOptimistic
      With Rst1
         .AddNew
         !PrimaryCode = "NCDC"
         !Code = "Dr"
         !Value = "Debit"
         .Update
         .AddNew
         !PrimaryCode = "NCDC"
         !Code = "Cr"
         !Value = "Credit"
         .Update
      End With
   End If
   Rst1.Close

'   Add a columns SMTP on 02/02/2011 ShoppingCentre
'###############################################################################################################
ADD_SMTP_ShoppingCentre:
   On Error GoTo CHANGE_ADD_SMTP_ShoppingCentre

   Rst1.Open "SELECT SMTP FROM ShoppingCentre;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_SLControl_Tenants

CHANGE_ADD_SMTP_ShoppingCentre:
   Conn1.Execute "ALTER TABLE ShoppingCentre ADD COLUMN SMTP TEXT(15);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns SLControl on 03/07/2013 Tenants
'###############################################################################################################
ADD_SLControl_Tenants:
   On Error GoTo CHANGE_ADD_SLControl_Tenants

   Rst1.Open "SELECT SLControl FROM Tenants;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_JobID_DemandSplitRecords

CHANGE_ADD_SLControl_Tenants:
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN SLControl TEXT(50);"
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN DefaultNC TEXT(10);"
   Conn1.Execute "ALTER TABLE Tenants ADD COLUMN VAT_CODE TEXT(5);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns JobID on 18/08/10 DemandSplitRecords
'###############################################################################################################
ADD_JobID_DemandSplitRecords:
   On Error GoTo CHANGE_ADD_JobID_DemandSplitRecords

   Rst1.Open "SELECT JobID FROM DemandSplitRecords;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo ADD_JobID_tlbReceiptSplit

CHANGE_ADD_JobID_DemandSplitRecords:
   Conn1.Execute "ALTER TABLE DemandSplitRecords ADD COLUMN JobID TEXT(10);"
   UpdateDatabase = 1
   Exit Function

'   Add a columns JobID on 18/08/10 tlbReceiptSplit
'###############################################################################################################
ADD_JobID_tlbReceiptSplit:
   On Error GoTo CHANGE_ADD_JobID_tlbReceiptSplit

   Rst1.Open "SELECT JobID FROM tlbReceiptSplit;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close

   GoTo tlbReceiptSplit

CHANGE_ADD_JobID_tlbReceiptSplit:
   Conn1.Execute "ALTER TABLE tlbReceiptSplit ADD COLUMN JobID TEXT(10);"
   UpdateDatabase = 1
   Exit Function

'   New table on 13/05/2010 tlbReceiptSplit
'###############################################################################################################
tlbReceiptSplit:
   On Error GoTo MissingTable_tlbReceiptSplit

   Rst1.Open "SELECT * FROM tlbReceiptSplit;", Conn1, adOpenStatic, adLockReadOnly

   Rst1.Close
   Exit Function

MissingTable_tlbReceiptSplit:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tlbReceiptSplit"
   UpdateDatabase = -1
   Exit Function
End Function

Private Sub CheckBankBalance()
   Dim rstBank As New ADODB.Recordset, rstBalance As New ADODB.Recordset, rstSC As New ADODB.Recordset
   Dim szSQL As String, cBalan As Currency

   On Error GoTo ErrHandler

   szSQL = "SELECT Field7 " & _
           "FROM ShoppingCentre;"
   rstSC.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, CB.CurrentBalance AS BAL " & _
           "FROM tlbClientBanks AS CB " & _
           "WHERE CB.CLIENT_ID <> '';"
'Debug.Print szSQL
   rstBank.Open szSQL, Conn1, adOpenDynamic, adLockOptimistic

   While Not rstBank.EOF
      szSQL = "SELECT SUM(R.Amount) AS AMT, 'SR' AS TID " & _
              "FROM tlbReceipt AS R, tlbTransactionTypes AS T " & _
              "WHERE R.BankCode = '" & rstBank.Fields.Item(0).Value & "' AND " & _
                  "R.Type = T.TYPE_ID "

      szSQL = szSQL + " UNION "

      szSQL = szSQL + _
              "SELECT SUM(P.Amount) AS AMT, 'PP' AS TID " & _
              "FROM tlbPayment AS P, tlbTransactionTypes AS T " & _
              "WHERE P.BankCode = '" & rstBank.Fields.Item(0).Value & "' AND " & _
                  "P.Type = T.TYPE_ID "

      szSQL = szSQL + " UNION "

      szSQL = szSQL + _
              "SELECT SUM( (BP.NET_AMOUNT + BP.VAT)) AS AMT, MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3) AS TID " & _
              "FROM tlbBankPayment AS BP, tlbTransactionTypes AS T " & _
              "WHERE BP.BANK_AC = '" & rstBank.Fields.Item(0).Value & "' AND " & _
                  "BP.TransactionType = T.TYPE_ID " & _
              "GROUP BY MID(T.CONSTANT, 4, LEN(T.CONSTANT)-3);"
      rstBalance.Open szSQL, Conn1, adOpenStatic, adLockReadOnly

      If rstBalance.EOF Then
         Set rstSC = Nothing
         Set rstBank = Nothing
         Set rstBalance = Nothing
         Exit Sub
      End If

      While Not rstBalance.EOF
         If Not IsNull(rstBalance.Fields.Item(0).Value) Then
            If rstBalance.Fields.Item(1).Value = "SR" Then cBalan = cBalan + rstBalance.Fields.Item(0).Value
            If rstBalance.Fields.Item(1).Value = "PP" Then cBalan = cBalan - rstBalance.Fields.Item(0).Value
            If rstBalance.Fields.Item(1).Value = "BP" Then cBalan = cBalan - rstBalance.Fields.Item(0).Value
            If rstBalance.Fields.Item(1).Value = "BR" Then cBalan = cBalan + rstBalance.Fields.Item(0).Value
         End If
         rstBalance.MoveNext
      Wend
      rstBalance.Close
      
      If rstBank.Fields.Item("BAL").Value = cBalan Then
         If Right(rstSC.Fields.Item(0).Value, 1) <> "S" Then
            rstSC.Fields.Item(0).Value = Right(rstSC.Fields.Item(0).Value, 49) & "#" & "S"
            rstSC.Update
         End If
      Else

         rstSC.Fields.Item(0).Value = Right(rstSC.Fields.Item(0).Value, 41) & "#" & "F" & Format(Now, "yyyymmdd")
         rstSC.Update
         
         rstBank.Fields.Item("BAL").Value = cBalan
         rstBank.Update
      End If
      
      rstBank.MoveNext
      cBalan = 0
   Wend

   rstSC.Close
   Set rstSC = Nothing

   rstBank.Close
   Set rstBank = Nothing
   Exit Sub

ErrHandler:
   ShowMsgInTaskBar Err.Number & " " & Err.description, , "N"
End Sub

Private Function Leaseid_RentAnalysis() As Boolean
   'if there are more than 1 default bank details for a client then mark them not default
   Dim adoRst As New ADODB.Recordset
   Dim objTable As ADODB.Field

   adoRst.Open "SELECT * FROM RentAnalysis;", Conn1, adOpenStatic, adLockReadOnly

   If adoRst.Fields.Item("LeaseID").DefinedSize <> 20 Then
      Leaseid_RentAnalysis = False
      adoRst.Close
      Set adoRst = Nothing

      Conn1.Execute "ALTER TABLE RentAnalysis ALTER COLUMN LeaseID text(20) NOT NULL"
   End If

   Leaseid_RentAnalysis = True
End Function

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  I AM TESTING HERE

'   If Button = 4 Then
'      Load frmLeaseViewSummary
'      frmLeaseViewSummary.Show
'   End If
   If Shift = 2 Then                         'Ctrl + Leftbutton
   Dim adoconn As New ADODB.Connection
   
   adoconn.Open getConnectionString
   adoconn.BeginTrans
   
   adoconn.Execute "update client set ClientOfficeEmail = 'yyyy';"
   
   adoconn.RollbackTrans
   
   adoconn.Close
   Set adoconn = Nothing
   End If
   
   If Shift = 1 Then                         'Shift + Leftbutton
   
   adoconn.Open getConnectionString
   adoconn.BeginTrans
   
   adoconn.Execute "update client set ClientOfficeEmail = 'yyyy';"
   
   adoconn.CommitTrans
   
   adoconn.Close
   Set adoconn = Nothing
   End If
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = ""
   If Not rtxtMessageDisplay.Visible Then StatusMsgBarLoad
   Me.MousePointer = vbArrow
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 And Shift = 3 Then
'      frmTemp1.Show
   End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim szTemp() As String

   If Not bLogOFF Then
          If MsgBox("Are you sure you want to quit?", vbOKCancel + vbQuestion, "Exit") = vbCancel Then
             Cancel = 1
          End If
          If IsLoadedAndVisible("frmPopUpMenu") Then Unload frmPopUpMenu
          If IsLoadedAndVisible("frmReport") Then Unload frmReport
        

          If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
                 Conn1.Open getConnectionString
                 Rst1.Open "SELECT UserName FROM UserNames WHERE UserName = 'samrat';", Conn1, adOpenStatic, adLockReadOnly
        
                 If Not Rst1.EOF Then
                    Rst1.Close
                    Conn1.Execute "UPDATE Tenants SET Email1 = spare8, Email2 = spare9;"
                    Conn1.Execute "UPDATE Supplier SET SupplierOfficeEmail = SageSuppAC;"
                    Conn1.Execute "DELETE * FROM UserNames WHERE UserName = 'samrat';"
                    Conn1.Execute "UPDATE tlbClientBanks SET FileLoc = FileLoc_;"
        
                    Rst1.Open "SELECT Field8 FROM ShoppingCentre WHERE Field8 <> '';", Conn1, adOpenStatic, adLockReadOnly
                    If Not Rst1.EOF Then
                       szTemp = Split(Rst1.Fields.Item(0).Value, "#")
                       Rst1.Close
        
                       Conn1.Execute "UPDATE ShoppingCentre " & _
                                     "SET SMTP = '" & szTemp(0) & "', " & _
                                         "UName = '" & szTemp(1) & "', " & _
                                         "Pws = '" & szTemp(2) & "', " & _
                                         "Port = " & Val(szTemp(3)) & ";"
                    Else
                       Rst1.Close
                    End If
                 Else
                    Rst1.Close
                 End If
        
                 Conn1.Close
          End If
          LastTenBackup 'implemnted by anol 20170130
          
   End If
   'Code added by mahboob 10/07/2023
   End
End Sub


Private Sub mmuLetters_Click()
   LoadForm frmTemplate
   frmTemplate.szLetter = "LT"
'   frmTemplate.Show
End Sub

Public Sub mmuVacUnits_Click()
    LoadForm frmAvailableUnitsReport
    'frmAvailableUnitsReport.Show
End Sub
Public Sub mmuPropertyList_Click()
    LoadForm frmRepPropertyList
    'frmRepPropertyList.Show
End Sub
Private Sub mnuAbout_Click()
   LoadForm frmAbout
   'frmAbout.Show
   'frmAbout.ZOrder 0
End Sub

Private Sub mnuAgedCreditors_Click()
   LoadForm frmAgedReports
   frmAgedReports.szWhoIsCalling = "Creditors"
   'frmAgedReports.Show
'   frmMMain.Arrange vbCascade
   'frmAgedReports.ZOrder 0
End Sub

Private Sub mnuAgedDebtors_Click()
   LoadForm frmAgedReports
   frmAgedReports.szWhoIsCalling = "Debtors"
   'frmAgedReports.Show
   'frmAgedReports.ZOrder 0
End Sub

Private Sub mnuBackup_Click()
   If CheckAnyFormOpen Then
        Exit Sub
   End If
   If MsgBox("All users should be out of the system when a backup is made." & Chr(13) & _
             "Do you wish to continue taking a backup?", vbYesNo + vbQuestion, "Data Backup") = vbYes Then
      If BackupDB Then
         MsgBox "Backup successful.", vbInformation, "The Backup has been successful"
      Else
         MsgBox "Backup failed. Please try again", vbInformation, "Warning"
      End If
   End If
End Sub
Private Function CheckAnyFormOpen() As Boolean
'"You must close all screens before continuing with this backup! <PC Name> <User Name> <Screen Name>
    Dim szpcName As String
    Dim szUserName As String
    Dim szScreenName  As String
    Dim rstPayment As New ADODB.Recordset
    Dim rstReceipt As New ADODB.Recordset
    Dim rstBankPayment As New ADODB.Recordset
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    
     rstPayment.Open "Select distinct MachineName,WindowsUserName,Module from (Select Distinct MachineName,WindowsUserName,Module from tlbPayment where ServerIPaddress<>''" & _
     " UNION ALL " & _
     "Select Distinct MachineName,WindowsUserName,Module from tlbReceipt where ServerIPaddress<>''" & _
      " UNION ALL " & _
      "Select Distinct MachineName,WindowsUserName,Module from tlbBankPayment where ServerIPaddress<>'')", adoconn, adOpenStatic, adLockReadOnly
   While Not rstPayment.EOF
        szpcName = szpcName & "  " & rstPayment("MachineName").Value
        szUserName = szUserName & "  " & rstPayment("WindowsUserName").Value
        szScreenName = szScreenName & "  " & rstPayment("Module").Value
        rstPayment.MoveNext
   Wend
   rstPayment.Close
   Set rstPayment = Nothing
   
   
'    rstPayment.Open "Select Distinct MachineName,WindowsUserName,Module from tlbPayment where ServerIPaddress<>''", adoconn, adOpenStatic, adLockReadOnly
'   While Not rstPayment.EOF
'        szpcName = szpcName & "  " & rstPayment("MachineName").Value
'        szUserName = szUserName & "  " & rstPayment("WindowsUserName").Value
'        szScreenName = szScreenName & "  " & rstPayment("Module").Value
'        rstPayment.MoveNext
'   Wend
'   rstPayment.Close
'   Set rstPayment = Nothing
'    rstReceipt.Open "Select Distinct MachineName,WindowsUserName,Module from tlbReceipt where ServerIPaddress<>''", adoconn, adOpenStatic, adLockReadOnly
'    While Not rstReceipt.EOF
'        szpcName = szpcName & "  " & rstReceipt("MachineName").Value
'        szUserName = szUserName & "  " & rstReceipt("WindowsUserName").Value
'        szScreenName = szScreenName & "  " & rstReceipt("Module").Value
'        rstReceipt.MoveNext
'   Wend
'   rstReceipt.Close
'   Set rstReceipt = Nothing
'    rstBankPayment.Open "Select Distinct MachineName,WindowsUserName,Module from tlbBankPayment where ServerIPaddress<>''", adoconn, adOpenStatic, adLockReadOnly
'    While Not rstBankPayment.EOF
'        szpcName = szpcName & "  " & rstBankPayment("MachineName").Value
'        szUserName = szUserName & "  " & rstBankPayment("WindowsUserName").Value
'        szScreenName = szScreenName & "  " & rstBankPayment("Module").Value
'        rstBankPayment.MoveNext
'  Wend
'  rstBankPayment.Close
'  Set rstBankPayment = Nothing
    adoconn.Close
    Set adoconn = Nothing
    If szScreenName <> "" Then
        MsgBox "Please close the following module(s) before continuing with this backup! " & vbCrLf & "PC Name: " & szpcName & " , User Name: " & szUserName & " , Module Name:" & szScreenName & ".", vbInformation, "Warning"
        CheckAnyFormOpen = True
    Else
        CheckAnyFormOpen = False
    End If
'      Conn1.Execute "UPDATE tlbPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress=''"
'      Conn1.Execute "UPDATE tlbReceipt P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
'      Conn1.Execute "UPDATE tlbBankPayment P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
'      Conn1.Execute "UPDATE NJ_Header P Set DateTimeStamp='',Module='',UserSessionID='',WindowsUserName='',MachineName='',PrestigeUserName='',ServerIPaddress='' "
End Function
Private Sub mnuBACSEmailTemaplate_Click()
   LoadForm frmEmailTemplate
   frmEmailTemplate.Caption = "BACS Email Template"
'   frmEmailTemplate.Show
'   frmEmailTemplate.ZOrder 0
End Sub

Public Sub mnuBAE_Click()
'Modified by anol 08 July 2015
'(a) Budget v Actual - Add input form to select client and property
'   Load frmPreSCBudVsAE
'   frmPreSCBudVsAE.Show
    LoadForm frmPreSCExpBudRpt
    frmPreSCExpBudRpt.Caption = "Budget Versus Actual Comparison"
    frmPreSCExpBudRpt.LOOKUPCommand = "A"
'    frmPreSCExpBudRpt.Left = 0
'    frmPreSCExpBudRpt.Top = 0
'    frmPreSCExpBudRpt.Show
'    frmPreSCExpBudRpt.ZOrder 0
End Sub

Private Sub mnuBank_Click()
   LoadForm frmBank
'   frmBank.Show
'   frmBank.SetFocus
'   frmBank.ZOrder 0
End Sub

Private Sub mnuBankTransactions_Click()
 'issue 496  Bank Receipt and Payment
   'Added by anol 13 Nov 2014
'         Dim strTemp As String
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If
   LoadForm frmBankTransactions
'   frmBankTransactions.Show
'   frmBankTransactions.ZOrder 0
End Sub

Private Sub mnuBatchPayment__Click()
'Batch Payments
   'issue 496 Batch Payments
   'Added by anol 13 Nov 2014
'         Dim strTemp As String
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If
Rem by anol 20190408 issue 749 Locking records
'   If IsFormLoaded("frmPurchaseExpense") Or IsLoadedAndVisible("frmPurchaseExpense") Then
'      MsgBox "Please close the Purchase and Expenses Screen before running Batch Payments.", vbInformation + vbOKOnly, "Batch Payment"
'      Exit Sub
'   End If
   
   On Error GoTo MissingTable_GlobalRC

   Conn1.Open getConnectionString

   Rst1.Open "SELECT * FROM tblBatchPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Rst1.Open "SELECT * FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Conn1.Close
   Set Conn1 = Nothing

   If frmBatchPayment.bBPPreForm Then
      'ShowMsgInTaskBar "Batch process is already open.", , "N"
      frmBatchPayment.Top = IsLoadedAndVisibleCount * 100
      frmBatchPayment.Left = IsLoadedAndVisibleCount * 100
      frmBatchPayment.Visible = True
      frmBatchPayment.ZOrder 0
      
   Else
      LoadForm frmBPPreForm
'      frmBPPreForm.Show
'      frmBPPreForm.ZOrder 0
   End If

   Exit Sub
MissingTable_GlobalRC:
   MsgBox "The database is not prepared for batch payment. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBatchPayment, tblBatchTransaction"
   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub mnuBPP_Click()
   If IsLoadedAndVisible("frmBACSFiles") Then
      If frmBACSFiles.Caption = "Payment Processed" Then
         ShowMsgInTaskBar "Payment Processed window is open.", "Y", "N"
         Exit Sub
      End If
   End If

   On Error GoTo MissingTable_GlobalRC

   Conn1.Open getConnectionString

   Rst1.Open "SELECT * FROM tblBatchPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Rst1.Open "SELECT * FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Conn1.Close
   Set Conn1 = Nothing

   LoadForm frmBACSFiles
'   frmBACSFiles.Show
'   frmBACSFiles.ZOrder 0
   Exit Sub
MissingTable_GlobalRC:
   MsgBox "The database is not prepared for batch payment. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBatchPayment, tblBatchTransaction"
   Conn1.Close
   Set Conn1 = Nothing
End Sub

Public Sub mnuBR_Click()
'   Load frmPreBudgetPnL
'   frmPreBudgetPnL.Show
'Modified by anol 23 Aug 2015
'(a) Budget v profit and loss - Add input form to select client and property
    LoadForm frmPreSCExpBudRpt
    frmPreSCExpBudRpt.Caption = "Budget profit and loss report"
    frmPreSCExpBudRpt.LOOKUPCommand = "B"
'    frmPreSCExpBudRpt.Left = 0
'    frmPreSCExpBudRpt.Top = 0
    frmPreSCExpBudRpt.Show
    frmPreSCExpBudRpt.ZOrder 0
End Sub

Private Sub mnuBS_Click()
   LoadForm frmPreBalnceSheet
'   frmPreBalnceSheet.Show
'   frmPreBalnceSheet.ZOrder 0
End Sub

Public Sub mnuCashbookHistoryReport_Click()
   LoadForm frmPreCBHistory
'   frmPreCBHistory.Show
'   frmPreCBHistory.ZOrder 0
End Sub

Public Sub mnuCashbookTransactionReport_Click()
   LoadForm frmPreCBTransactions
'   frmPreCBTransactions.Show
'   frmPreCBTransactions.ZOrder 0
End Sub

Public Sub mnuChangeClientID_Click()
   frmChangeLesseeID.WhoAmI = "ClientID"
   LoadForm frmChangeLesseeID
'   frmChangeLesseeID.Show
'   frmChangeLesseeID.ZOrder 0
End Sub

Public Sub mnuChangeLesseeID_Click()
   frmChangeLesseeID.WhoAmI = "LesseeID"
   LoadForm frmChangeLesseeID
   'frmChangeLesseeID.Show
End Sub

Public Sub mnuChangePropertyID_Click()
   frmChangeLesseeID.WhoAmI = "PropertyID"
   LoadForm frmChangeLesseeID
   'frmChangeLesseeID.Show
End Sub

Public Sub mnuChangeSupplierID_Click()
   frmChangeLesseeID.WhoAmI = "SupplierID"
   LoadForm frmChangeLesseeID
'   frmChangeLesseeID.Show
'   frmChangeLesseeID.ZOrder 0
End Sub

Private Sub mnuChangePassword_Click()
    Load frmChangePassword1
    frmChangePassword1.Show
    frmChangePassword1.ZOrder
End Sub 'mnuChangePassword_Click()

Public Sub mnuChangeUnitID_Click()
   frmChangeLesseeID.WhoAmI = "UnitID"
   LoadForm frmChangeLesseeID
'   frmChangeLesseeID.Show
'   frmChangeLesseeID.ZOrder 0
End Sub

Private Sub mnuChargeTypes_Click()
'   Load frmDCTypesPre
'   frmDCTypesPre.szMenu = "CHARGE_TYPE"
'   frmDCTypesPre.Show
'   frmDCTypesPre.ZOrder 0

   LoadForm frmChargeTypes
'   frmFeeChargeTypes.Show
'   frmFeeChargeTypes.ZOrder 0
End Sub

Private Sub mnuChartAccountsReport_Click()
   LoadForm frmPreChartOfAccunts
'   frmPreChartOfAccunts.Show
'   frmPreChartOfAccunts.ZOrder 0
   
Exit Sub
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\NL_List.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub mnuCodes_Click()
   If IsAnyFormVisible Then
      ShowMsgInTaskBar "Please close all opened forms to continue.", "Y", "N"
      Exit Sub
   End If
   
   Load frmSecondaryCode
   frmSecondaryCode.Show 1
End Sub

Private Sub mnuCompanySetup_Click()
   Load frmShoppingCentre
   frmShoppingCentre.Show
End Sub

Private Sub mnuControlCode_Click()
   LoadForm frmControlCode
   'frmControlCode.Show
End Sub

Private Sub mnuCR_Click()
    'Batch Receipt
    'issue 496 Batch Receipt
   'Added by anol 13 Nov 2014
'         Dim strTemp As String
'         strTemp = isControlAccountSet
'         If Len(strTemp) > 0 Then
'            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'            Exit Sub
'         End If

'   If IsFormLoaded("frmDemands3") Or IsLoadedAndVisible("frmDemands3") Then
'      If MsgBox("This option cannot be run with the Demand Window Open." & Chr(13) & _
'                "Do you wish to close the demand Window?" & Chr(13) & Chr(13) & _
'                "WARNING: Any unsaved data will be lost.", vbInformation + vbYesNo, "Batch Receipt") = vbYes Then
'         Unload frmDemands3
'      Else
'         Exit Sub
'      End If
'   End If

   On Error GoTo MissingTable_GlobalRC

   If frmBatchPayment.bBPPreForm Then
      ShowMsgInTaskBar "Batch process is already open.", vbInformation + vbOKOnly, "Batch Process"
   Else
      Dim adoconn As New ADODB.Connection
      Dim adoRst  As New ADODB.Recordset
      adoconn.Open getConnectionString
   
      adoRst.Open "SELECT PropertyID FROM DemandTypes WHERE spare1='';", adoconn, adOpenStatic, adLockReadOnly

      If Not adoRst.EOF Then
         adoRst.Close
         Set adoRst = Nothing
         adoconn.Close
         Set adoconn = Nothing
         ShowMsgInTaskBar "Bank has not been setup in the demand type, for PropertyID: " & adoRst("PropertyID").Value, "Y", "N"
         Exit Sub
      End If
      adoRst.Close
      Set adoRst = Nothing
      adoconn.Close
      Set adoconn = Nothing
      
      LoadForm frmBRPreForm
'      frmBRPreForm.Show
'      frmBRPreForm.ZOrder 0
   End If

   Exit Sub
MissingTable_GlobalRC:
   MsgBox "The database is not prepared for batch receipt. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBatchReceipt & tblBtRptTran"
   adoconn.Close
   Set Conn1 = Nothing
End Sub

Public Sub mnuDemandAnalysisReport_Click()
   LoadForm frmViewInvDtRange
   'frmViewInvDtRange.Show
End Sub

Private Sub mnuDemandEmailTemplate_Click()
   frmEmailTemplate.Caption = "Demand Email Template"
   LoadForm frmEmailTemplate
'   frmEmailTemplate.Show
'   frmEmailTemplate.ZOrder 0
End Sub
Private Sub mnuClientStatementTemplate_Click()
   frmEmailTemplate.Caption = "Client Statement Email Template"
   LoadForm frmEmailTemplate

End Sub
Private Sub mnufrmAuditLog_Click()
   
   LoadForm frmAuditLog

End Sub
Private Sub mnufrmViewCashBalance_Click()
   LoadForm frmViewBankBalance
End Sub

Private Sub mnuStatementTemplate_Click()
   frmEmailTemplate.Caption = "Statement Email Template"
   LoadForm frmEmailTemplate
'   frmEmailTemplate.Show
'   frmEmailTemplate.ZOrder 0
End Sub
'
'Private Sub mnuDemandPreviewReport_Click()
''     ************* #################################### This menu is hidden.
''     ************* #################################### B4 visible this menu, make sure 'GenSngDmdPreView' bug free
'   Dim szChoice As String, szaChoice() As String
'   Dim adoConn As New ADODB.Connection
'   Dim szSQL As String, i As Integer
'
'   MousePointer = vbHourglass
'
''   connect to database
'   adoConn.Open getConnectionString
'
'   GenSngDmdPreView adoConn
'
'   szSQL = "UPDATE DemandRecPreview " & _
'           "SET DemandRecPreview.spare2 = 'Y';"
'
'   adoConn.Execute szSQL
'
'   MousePointer = vbDefault
'
'   ShowReport App.Path & szReportPath & "\PrintPreGeneratedDemandList.rpt"
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub
'
'Private Sub GenSngDmdPreView(adoConn As ADODB.Connection)
'   Dim szaNCode() As String, szaNCName() As String, szaPrefix() As String, szaTemp() As String
'   Dim szDes As String, CutOffDate As String, szSQLStr As String
'
'   Dim BRcount As Integer, SCcount As Integer, IPcount As Integer, ICcount As Integer
'   Dim NextUniqueRefNo As Long, iProp As Integer, iSerial As Integer
'   Dim lDemand As Long, iChildId As Integer
'   Dim DaysB4Due As Integer
'
'   Dim dtEndDate As Date, dtNtDueDate As Date
'
'   Dim adoRstDemandRec As ADODB.Recordset, adoRstDmdTyp As ADODB.Recordset
'   Dim adoRstLeaseDtl As ADODB.Recordset, adoRstSplitDemand As ADODB.Recordset
'   Dim adoRstRC As ADODB.Recordset, adoRstSC As ADODB.Recordset
'   Dim adoRstPro As ADODB.Recordset, adoRstIns As ADODB.Recordset
'
'   If MsgBox("Please ensure your global data has been correctly inputted.", vbYesNo + vbQuestion, _
'             "Generate Automatic Demands") = vbNo Then Exit Sub
'
''   On Error GoTo ErrH
'
'   MousePointer = vbHourglass    'change the mouse cursor to show program is busy/working
'
''  Empty the temporary table
'   adoConn.Execute "DELETE * FROM DemandRecPreview;"
'   adoConn.Execute "DELETE * FROM DemandSplPreview;"
'
''   Connect to Demands table to add new demands.
'   Set adoRstDemandRec = New ADODB.Recordset
'   szSQLStr = "SELECT * FROM DemandRecPreview"
'   adoRstDemandRec.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
'
'   Set adoRstSplitDemand = New ADODB.Recordset
'   szSQLStr = "SELECT * FROM DemandSplPreview"
'   adoRstSplitDemand.Open szSQLStr, adoConn, adOpenDynamic, adLockPessimistic
'
''   get nominal codes and prefix for base rent from demand types.
'   Set adoRstDmdTyp = New ADODB.Recordset
'   szSQLStr = "SELECT * FROM DemandTypes;"
'   adoRstDmdTyp.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRstDmdTyp.EOF Then
'      ShowMsgInTaskBar "There are no Demand Type in the database", , "N"
'
'      adoRstSplitDemand.Close
'      adoRstDemandRec.Close
'      adoRstDmdTyp.Close
'
'      Set adoRstSplitDemand = Nothing
'      Set adoRstDemandRec = Nothing
'      Set adoRstDmdTyp = Nothing
'
'      Exit Sub
'   End If
'
'   ReDim szaPrefix(adoRstDmdTyp.RecordCount) As String
'   ReDim szaNCName(adoRstDmdTyp.RecordCount) As String
'   ReDim szaNCode(adoRstDmdTyp.RecordCount) As String
'
''**Saving all Nominal Codes and Prefixes in the array
'   While Not adoRstDmdTyp.EOF
'      szaPrefix(adoRstDmdTyp!ID) = adoRstDmdTyp!prefix
'      szaNCName(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalNameforAmount & " # " & adoRstDmdTyp!NominalNameforVAT & " # " & adoRstDmdTyp!NominalNameforTotal
'      szaNCode(adoRstDmdTyp!ID) = adoRstDmdTyp!NominalCodeforAmount & " # " & adoRstDmdTyp!NominalCodeForVAT & " # " & adoRstDmdTyp!NominalCodeForTotal
'      adoRstDmdTyp.MoveNext
'   Wend
'   adoRstDmdTyp.Close
'   Set adoRstDmdTyp = Nothing
'
'   Set adoRstLeaseDtl = New ADODB.Recordset
'   Set adoRstPro = New ADODB.Recordset
'
'   szSQLStr = "SELECT LeaseDetails.*, Units.PropertyID " & _
'              "FROM LeaseDetails, Units " & _
'              "WHERE LeaseDetails.Status = TRUE AND " & _
'                  "(OLED = TRUE OR DATEDIFF('D', NOW, ENDDATE) >= 0) AND " & _
'                  "LeaseDetails.UnitNumber = Units.UnitNumber;"
'   adoRstLeaseDtl.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
'   iSerial = 1
'
'   If adoRstLeaseDtl.EOF Then
'      adoRstLeaseDtl.Close
'      Set adoRstLeaseDtl = Nothing
'   Else
'      While Not adoRstLeaseDtl.EOF
'
''*********************** SAMRAT 25/11/2005***************************************
''*** Determin the date boundray in future, in this boundary
''*** we have to find those demands' due date to calcuate
''*** demands from lease table and we have to calculate &
''*** collect all those demands in DemandRecPreview table.
''*********************************************************************************
'         DaysB4Due = GlbDaysBeforeDue(adoRstLeaseDtl.Fields("LeaseID").Value, adoConn)
'         CutOffDate = DateAdd("d", DaysB4Due, Date) 'Date boundary
'
''************************************************************************************
''*********************        Interest Charge            ****************************
''************************************************************************************
'         If adoRstLeaseDtl.Fields("InterestChargeable").Value = "Y" Then
'            Dim cTotalOSAmt As Currency, sIntCalDays As Single
''**** Insert the Header info in the DemandRecPreview table
'            lDemand = lDemand + 1
'
'            With adoRstDemandRec
'               .AddNew
'               .Fields("DemandID").Value = lDemand
'               .Fields("BatchID").Value = 1
'               .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'               .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'               .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'               .Fields("Source").Value = 1
'               .Fields("TransactionType").Value = 1
''*** Here my thinking is, all type of demands due date is on the same day
''*** If its not correct then i have to change the manual demands grid and the demand table & split table
'               .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
'               .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'               .Fields("IsPrinted").Value = False
''               .Fields("UPDATE_SAGE").Value = False
'               .Fields("Spare1").Value = ClientDefaultBankDts(adoRstLeaseDtl!PropertyID, adoConn)
'               .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'               .Update
'            End With
'
''*** Insert the split records in the DemandSplPreview table
'            iChildId = 1
'
''*** Add new demand IN demand table.
'            adoRstSplitDemand.AddNew
'            adoRstSplitDemand!DSR = UniqueID()
'            adoRstSplitDemand!SplitID = iChildId
'            adoRstSplitDemand!DemandId = lDemand
'            If adoRstLeaseDtl!ServiceChargeDept = "AUTO" Then
'               adoRstSplitDemand!A_M = "A"
'            Else
'               adoRstSplitDemand!A_M = "M"
'            End If
'            adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 0, " # ")
'            adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 0, " # ")
'            adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstLeaseDtl!IntDemandType), 2, " # ")
'            adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstLeaseDtl!IntDemandType), 2, " # ")
'            If adoRstLeaseDtl!ServiceChargeDept = "AUTO" Then
'               adoRstSplitDemand!Amount = CalAutoInterest(CSng(adoRstLeaseDtl!AdditionalInterest), CInt(adoRstLeaseDtl!DaysAfterInterestPayable), adoRstLeaseDtl!UnitNumber, adoRstLeaseDtl!SageAccountNumber, adoConn, cTotalOSAmt, sIntCalDays)
'            Else
'               adoRstSplitDemand!Amount = adoRstLeaseDtl!InterestAmount
'            End If
'            adoRstSplitDemand!VATAmount = 0
'            adoRstSplitDemand!TotalAmount = CCur(adoRstSplitDemand!Amount) + _
'                                            CCur(adoRstSplitDemand!VATAmount)
'            If adoRstLeaseDtl!Text1 = "" Then
'               szDes = "Interest Charges For " & adoRstLeaseDtl!DaysAfterInterestPayable & _
'                       " Days on " & cTotalOSAmt
'            Else
'               szDes = adoRstLeaseDtl!Text1
'            End If
'            adoRstSplitDemand!SageRef = szaPrefix(adoRstLeaseDtl!IntDemandType) & Right(UniqueID(), 18)
'            adoRstSplitDemand!DueDate = Format(Date, "dd/mm/yyyy")
'            adoRstSplitDemand!VATMonth = Month(Date)
'            adoRstSplitDemand!TypeOfDemand = adoRstLeaseDtl!IntDemandType
'            adoRstSplitDemand!description = szDes
'            adoRstSplitDemand!DemandStatement = True
'            adoRstSplitDemand!FrequencyID = ""
'            adoRstSplitDemand.Update
'
'            IPcount = IPcount + 1
'            iSerial = iSerial + 1
'         End If
''*********************************************************************************************************
''         Rent Charges Demands
''*********************************************************************************************************
'         szSQLStr = "SELECT LRentCharges.RentCharges, LRentCharges.LeaseID, " & _
'                        "LRentCharges.RentChargeDept, LRentCharges.BRFrequency, " & _
'                        "LRentCharges.BRStartDate, LRentCharges.BRNextDueDate, " & _
'                        "LRentCharges.BRTotal, LRentCharges.BRAmount, " & _
'                        "LRentCharges.BRDemandType, LRentCharges.RentDesc, " & _
'                        "Units.PropertyID, DemandTypes.Spare1 as ClientBankID, " & _
'                        "Frequencies.CalDays " & _
'                    "FROM LRentCharges, LeaseDetails, Units, DemandTypes, Frequencies " & _
'                    "WHERE LRentCharges.LeaseID = '" & adoRstLeaseDtl!LeaseID & "' AND " & _
'                        "LRentCharges.LeaseID = LeaseDetails.LeaseID AND " & _
'                        "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
'                        "LRentCharges.BRTotal > 0 AND " & _
'                        "LRentCharges.BRDemandType = DemandTypes.ID AND " & _
'                        "LRentCharges.BRFrequency = Frequencies.ID;"
'
'         Set adoRstRC = New ADODB.Recordset
'         adoRstRC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
'         While Not adoRstRC.EOF
'            If adoRstLeaseDtl!BRPayable = "Y" And _
'               DateDiff("d", Date, IIf(adoRstRC!BRNextDueDate = "", _
'                  DateAdd("d", -1, Date), adoRstRC!BRNextDueDate)) <= DaysB4Due Then
'   '**** Insert the Header info in the DemandRecPreview table
'               lDemand = lDemand + 1
'               With adoRstDemandRec
'                  .AddNew
'                  .Fields("DemandID").Value = lDemand
'                  .Fields("BatchID").Value = 1
'                  .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'                  .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'                  .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'                  .Fields("Source").Value = 1
'                  .Fields("TransactionType").Value = 1
'   '***Here my thinking is, all type of demands due date is on the same day
'   '***If its not correct then i have to change the manual demands grid and the demand table & split table
'                  .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
'                  .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'                  .Fields("IsPrinted").Value = False
'                  .Fields("Spare1").Value = adoRstRC!ClientBankID
'                  .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'                  .Update
'               End With
'
'               iChildId = 1
'
'               szDes = IIf(IIf(IsNull(adoRstRC!RentDesc), "", adoRstRC!RentDesc) = "", DemandTypeName(adoRstRC!BRDemandType, adoConn), adoRstRC!RentDesc)
'
'               dtNtDueDate = FindNextDueDate(adoRstRC!BRNextDueDate, _
'                                 adoRstRC!BRfrequency, adoRstRC!BRDemandType, adoRstRC!PropertyID, adoConn)
'
''******* if the override lease end date is false then the lease is open, lease is not expairing.
''******* if the lease end date is open then we dont need to compare the NextDueDate with the lease expaire date.
'               If Not adoRstLeaseDtl.Fields("OLED").Value Then
'                  dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
'                  If DateDiff("d", dtEndDate, dtNtDueDate) > 0 Then dtNtDueDate = dtEndDate
'               End If
'
'               adoRstSplitDemand.AddNew
'               adoRstSplitDemand!DSR = UniqueID()
'               adoRstSplitDemand!SplitID = iChildId
'               adoRstSplitDemand!DemandId = lDemand
'               adoRstSplitDemand!A_M = "A"
'               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstRC!BRDemandType), 0, " # ")
'               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstRC!BRDemandType), 0, " # ")
'               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstRC!BRDemandType), 1, " # ")
'               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstRC!BRDemandType), 1, " # ")
'               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstRC!BRDemandType), 2, " # ")
'               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstRC!BRDemandType), 2, " # ")
'               adoRstSplitDemand!Amount = adoRstRC!BRAmount
'               adoRstSplitDemand!VATAmount = Round(adoRstRC!BRAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn) / 100, 2)
'               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
'               adoRstSplitDemand!SageRef = szaPrefix(adoRstRC!BRDemandType) & Right(UniqueID(), 18)
'               adoRstSplitDemand!DueDate = Format(adoRstRC!BRNextDueDate.Value, "dd/mm/yyyy")
'               adoRstSplitDemand!VATMonth = Month(adoRstRC!BRNextDueDate)
'               adoRstSplitDemand!TypeOfDemand = adoRstRC!BRDemandType
'               adoRstSplitDemand!description = szDes
'               adoRstSplitDemand!DemandStatement = True
'               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn)
'               If Left(adoRstRC!CalDays, 1) <> "-" Then
''                          ADVANCE
'                  adoRstSplitDemand!DateFrom = CDate(adoRstRC!BRNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, dtNtDueDate)
'               Else
''                          ARREARS
'                  adoRstSplitDemand!DateFrom = DateAdd(Right(adoRstRC!CalDays, 1), Left(adoRstRC!CalDays, Len(adoRstRC!CalDays) - 1), adoRstRC!BRNextDueDate) 'CDate(adoRstRC!BRNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, adoRstRC!BRNextDueDate)
'               End If
'               adoRstSplitDemand!SageDepartment = adoRstRC!RentChargeDept ' DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Rent Charges", adoConn)
'               adoRstSplitDemand!FrequencyID = adoRstRC!BRfrequency
'               adoRstSplitDemand.Update
'
'               BRcount = BRcount + 1
'               iSerial = iSerial + 1
'            End If
'            adoRstRC.MoveNext
'         Wend
'         adoRstRC.Close
'         Set adoRstRC = Nothing
''************************************************************************************************
''   Service Charge demands
''************************************************************************************************
'         szSQLStr = "SELECT LServiceCharges.ServiceCharge, LServiceCharges.LeaseID, " & _
'                        "LServiceCharges.SCFrequency, LServiceCharges.SCPayableFrom, " & _
'                        "LServiceCharges.SCNextDueDate, LServiceCharges.ChargingMethod, " & _
'                        "LServiceCharges.CMFigure, LServiceCharges.SCTotal, " & _
'                        "LServiceCharges.SCAmount, LServiceCharges.SCTOLimit, " & _
'                        "LServiceCharges.SCDemandType, LServiceCharges.ServiceChargeDept, " & _
'                        "LServiceCharges.SCDesc, Units.PropertyID, " & _
'                        "DemandTypes.Spare1 as ClientBankID, Frequencies.CalDays " & _
'                    "FROM LServiceCharges, LeaseDetails, Units, DemandTypes, Frequencies " & _
'                    "WHERE LServiceCharges.LeaseID = '" & adoRstLeaseDtl!LeaseID & "' AND " & _
'                        "LServiceCharges.LeaseID = LeaseDetails.LeaseID AND " & _
'                        "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
'                        "LServiceCharges.SCAmount > 0 AND " & _
'                        "LServiceCharges.SCDemandType = DemandTypes.ID AND " & _
'                        "LServiceCharges.SCFrequency = Frequencies.ID;"
'
'         Set adoRstSC = New ADODB.Recordset
'
'         adoRstSC.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
'         While Not adoRstSC.EOF
'            If adoRstLeaseDtl!SCPayable = "Y" And _
'               DateDiff("d", Date, IIf(adoRstSC!SCNextDueDate = "", _
'                  DateAdd("d", -1, Date), adoRstSC!SCNextDueDate)) <= DaysB4Due Then
''**** Insert the Header info in the DemandRecPreview table
'               lDemand = lDemand + 1
'               With adoRstDemandRec
'                  .AddNew
'                  .Fields("DemandID").Value = lDemand
'                  .Fields("BatchID").Value = 1
'                  .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'                  .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'                  .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'                  .Fields("Source").Value = 1
'                  .Fields("TransactionType").Value = 1
'   '***Here my thinking is, all type of demands due date is on the same day
'   '***If its not correct then i have to change the manual demands grid and the demand table & split table
'                  .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
'                  .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'                  .Fields("IsPrinted").Value = False
'                  .Fields("Spare1").Value = adoRstSC!ClientBankID
'                  .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'                  .Update
'               End With
'
'               iChildId = 1
'               szDes = IIf(IIf(IsNull(adoRstSC!SCDesc), "", adoRstSC!SCDesc) = "", DemandTypeName(adoRstSC!SCDemandType, adoConn), adoRstSC!SCDesc)
'
'               dtNtDueDate = FindNextDueDate(adoRstSC!SCNextDueDate, _
'                                 adoRstSC!SCFrequency, adoRstSC!SCDemandType, adoRstSC!PropertyID, adoConn)
'
'               If Not adoRstLeaseDtl.Fields("OLED").Value Then
'                  dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
'                  If DateDiff("d", dtEndDate, dtNtDueDate) > 0 Then dtNtDueDate = dtEndDate
'               End If
'
'               adoRstSplitDemand.AddNew
'               adoRstSplitDemand!DSR = UniqueID()
'               adoRstSplitDemand!SplitID = iChildId
'               adoRstSplitDemand!DemandId = lDemand
'               adoRstSplitDemand!A_M = "A"
'               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstSC!SCDemandType), 0, " # ")
'               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstSC!SCDemandType), 0, " # ")
'               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstSC!SCDemandType), 1, " # ")
'               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstSC!SCDemandType), 1, " # ")
'               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstSC!SCDemandType), 2, " # ")
'               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstSC!SCDemandType), 2, " # ")
'               adoRstSplitDemand!Amount = CCur(adoRstSC!SCAmount)
'               adoRstSplitDemand!VATAmount = Round(adoRstSC!SCAmount * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn) / 100, 2)
'               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
'               adoRstSplitDemand!SageRef = szaPrefix(adoRstSC!SCDemandType) & Right(UniqueID(), 18)
'               adoRstSplitDemand!DueDate = Format(adoRstSC!SCNextDueDate.Value, "dd/mm/yyyy")
'               adoRstSplitDemand!VATMonth = Month(adoRstSC!SCNextDueDate)
'               adoRstSplitDemand!TypeOfDemand = adoRstSC!SCDemandType
'               adoRstSplitDemand!description = szDes
'               adoRstSplitDemand!DemandStatement = True
'               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn)
'               If Left(adoRstSC!CalDays, 1) <> "-" Then
''                          ADVANCE
'                  adoRstSplitDemand!DateFrom = CDate(adoRstSC!SCNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, dtNtDueDate)
'               Else
''                          ARREARS
'                  adoRstSplitDemand!DateFrom = DateAdd(Right(adoRstSC!CalDays, 1), Left(adoRstSC!CalDays, Len(adoRstSC!CalDays) - 1), adoRstSC!SCNextDueDate) 'CDate(adoRstSC!SCNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, adoRstSC!SCNextDueDate)
'               End If
'
'               adoRstSplitDemand!SageDepartment = adoRstSC!ServiceChargeDept ' DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Service Charge", adoConn)
'               adoRstSplitDemand!FrequencyID = adoRstSC!SCFrequency
'               adoRstSplitDemand.Update
'
'               SCcount = SCcount + 1
'               iSerial = iSerial + 1
'            End If
'            adoRstSC.MoveNext
'         Wend
'         adoRstSC.Close
'         Set adoRstSC = Nothing
''************************************************************************************************
''   Insurance Charge demands
''************************************************************************************************
'         szSQLStr = "SELECT LInsuranceCharges.InsCharges, LInsuranceCharges.LeaseID, " & _
'                        "LInsuranceCharges.InsuranceStartDate, LInsuranceCharges.InsuranceFrequency, " & _
'                        "LInsuranceCharges.InsuranceEndDate, LInsuranceCharges.InsuranceDemandType, " & _
'                        "LInsuranceCharges.InsuranceEachPeriod, LInsuranceCharges.InsuranceNextDueDate, " & _
'                        "LInsuranceCharges.ChargingType, " & _
'                        "LInsuranceCharges.ChargingFigure, LInsuranceCharges.TotalYearlyInsurance, " & _
'                        "LInsuranceCharges.InsuranceDept, LInsuranceCharges.InsDesc, " & _
'                        "Units.PropertyID, DemandTypes.Spare1 as ClientBankID, " & _
'                        "Frequencies.CalDays " & _
'                    "FROM LInsuranceCharges, LeaseDetails, Units, DemandTypes, Frequencies " & _
'                    "WHERE LInsuranceCharges.LeaseID = '" & adoRstLeaseDtl!LeaseID & "' AND " & _
'                        "LInsuranceCharges.LeaseID = LeaseDetails.LeaseID AND " & _
'                        "LeaseDetails.UnitNumber = Units.UnitNumber AND " & _
'                        "LInsuranceCharges.InsuranceEachPeriod > 0 AND " & _
'                        "LInsuranceCharges.InsuranceDemandType = DemandTypes.ID AND " & _
'                        "LInsuranceCharges.InsuranceFrequency = Frequencies.ID;"
'
'         Set adoRstIns = New ADODB.Recordset
'         adoRstIns.Open szSQLStr, adoConn, adOpenStatic, adLockReadOnly
'
'         While Not adoRstIns.EOF
'               If adoRstLeaseDtl!InsurancePayable = "Y" And _
'                  DateDiff("d", Date, IIf(adoRstIns!InsuranceNextDueDate = "", _
'                     DateAdd("d", -1, Date), adoRstIns!InsuranceNextDueDate)) <= DaysB4Due Then
'   '**** Insert the Header info in the DemandRecPreview table
'               lDemand = lDemand + 1
'               With adoRstDemandRec
'                  .AddNew
'                  .Fields("DemandID").Value = lDemand
'                  .Fields("BatchID").Value = 1
'                  .Fields("SageAccountNumber").Value = adoRstLeaseDtl!SageAccountNumber
'                  .Fields("TenantCompanyName").Value = adoRstLeaseDtl!CompanyName
'                  .Fields("UnitNumber").Value = adoRstLeaseDtl!UnitNumber
'                  .Fields("Source").Value = 1
'                  .Fields("TransactionType").Value = 1
'   '***Here my thinking is, all type of demands due date is on the same day
'   '***If its not correct then i have to change the manual demands grid and the demand table & split table
'                  .Fields("IssueDate").Value = Format(Date, "dd/mm/yyyy")
'                  .Fields("SageText").Value = "S/L " & adoRstLeaseDtl!SageAccountNumber
'                  .Fields("IsPrinted").Value = False
'                  .Fields("Spare1").Value = adoRstIns!ClientBankID
'                  .Fields.Item("LeaseRef").Value = adoRstLeaseDtl!LeaseID
'                  .Update
'               End With
'
'               iChildId = 1
'               szDes = IIf(IIf(IsNull(adoRstIns!InsDesc), "", adoRstIns!InsDesc) = "", DemandTypeName(adoRstIns!InsuranceDemandType, adoConn), adoRstIns!InsDesc)
'
'               dtNtDueDate = FindNextDueDate(adoRstIns!InsuranceNextDueDate, _
'                              adoRstIns!InsuranceFrequency, _
'                              adoRstIns!InsuranceDemandType, adoRstIns!PropertyID, adoConn)
'
'               If Not adoRstLeaseDtl.Fields("OLED").Value Then
'                  dtEndDate = Format(adoRstLeaseDtl!EndDate, "dd/mm/yyyy")
'                  If DateDiff("d", dtEndDate, dtNtDueDate) > 0 Then dtNtDueDate = dtEndDate
'               End If
'
'               adoRstSplitDemand.AddNew
'               adoRstSplitDemand!DSR = UniqueID()
'               adoRstSplitDemand!SplitID = iChildId
'               adoRstSplitDemand!DemandId = lDemand
'               adoRstSplitDemand!A_M = "A"
'               adoRstSplitDemand!NominalCodeforAmount = PartString(szaNCode(adoRstIns!InsuranceDemandType), 0, " # ")
'               adoRstSplitDemand!NominalNameforAmount = PartString(szaNCName(adoRstIns!InsuranceDemandType), 0, " # ")
'               adoRstSplitDemand!NominalCodeForVAT = PartString(szaNCode(adoRstIns!InsuranceDemandType), 1, " # ")
'               adoRstSplitDemand!NominalNameforVAT = PartString(szaNCName(adoRstIns!InsuranceDemandType), 1, " # ")
'               adoRstSplitDemand!NominalCodeForTotal = PartString(szaNCode(adoRstIns!InsuranceDemandType), 2, " # ")
'               adoRstSplitDemand!NominalNameforTotal = PartString(szaNCName(adoRstIns!InsuranceDemandType), 2, " # ")
'               adoRstSplitDemand!Amount = adoRstIns!InsuranceEachPeriod
'               adoRstSplitDemand!VATAmount = Round(adoRstIns!InsuranceEachPeriod * GetVAT_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn) / 100, 2)
'               adoRstSplitDemand!TotalAmount = adoRstSplitDemand!Amount + adoRstSplitDemand!VATAmount
'               adoRstSplitDemand!SageRef = szaPrefix(adoRstIns!InsuranceDemandType) & Right(UniqueID(), 18)
'               adoRstSplitDemand!DueDate = Format(adoRstIns!InsuranceNextDueDate, "dd/mm/yyyy")
'               adoRstSplitDemand!VATMonth = Month(adoRstIns!InsuranceNextDueDate)
'               adoRstSplitDemand!TypeOfDemand = adoRstIns!InsuranceDemandType
'               adoRstSplitDemand!description = szDes
'               adoRstSplitDemand!DemandStatement = True
'               adoRstSplitDemand!VAT_CODE = GetVATCode_Tenant(adoRstLeaseDtl!SageAccountNumber, adoConn)
'               If Left(adoRstIns!CalDays, 1) <> "-" Then
'               '           ADVANCE
'                  adoRstSplitDemand!DateFrom = CDate(adoRstIns!InsuranceNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, dtNtDueDate)
'               Else
'               '           ARREARS
'                  adoRstSplitDemand!DateFrom = DateAdd(Right(adoRstIns!CalDays, 1), Left(adoRstIns!CalDays, Len(adoRstIns!CalDays) - 1), adoRstIns!InsuranceNextDueDate) 'CDate(adoRstIns!InsuranceNextDueDate)
'                  adoRstSplitDemand!DateTo = DateAdd("d", -1, adoRstIns!InsuranceNextDueDate)
'               End If
'
'               adoRstSplitDemand!SageDepartment = adoRstIns!InsuranceDept ' DepartmentID(adoRstLeaseDtl!SageAccountNumber, adoRstLeaseDtl!UnitNumber, "Insurance Charge", adoConn)
'               adoRstSplitDemand!FrequencyID = adoRstIns!InsuranceFrequency
'               adoRstSplitDemand.Update
'
'               ICcount = ICcount + 1
'               iSerial = iSerial + 1
'            End If
'            adoRstIns.MoveNext
'         Wend
'         adoRstIns.Close
'         Set adoRstIns = Nothing
''MsgBox adoRstLeaseDtl.RecordCount
'         adoRstLeaseDtl.MoveNext
'      Wend
'
'      adoRstLeaseDtl.Close
'      adoRstDemandRec.Close
'      adoRstSplitDemand.Close
'
'      Set adoRstLeaseDtl = Nothing
'      Set adoRstDemandRec = Nothing
'      Set adoRstSplitDemand = Nothing
'   End If
'
'   MousePointer = vbDefault
'
'   Exit Sub
'ErrH:
'       'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
'       ShowMsgInTaskBar ERR.Number & " - (pcm_001)" & ERR.description, , "N"
'End Sub

Private Function PartString(szString As String, iStringPartNum As Integer, szDelemeter As String) As String
   Dim szaString() As String

   szaString = Split(szString, szDelemeter)
   PartString = szaString(iStringPartNum)
End Function

Private Function CalAutoInterest(ByVal sIntRate As Single, ByVal iDays As Integer, ByVal szUnitNum As String, ByVal szSAN As String, adoconn As ADODB.Connection, ByRef cTotalAmt As Currency, ByRef sDaysCalculated) As Currency
   Dim szSQL As String, dtStartDt As Date, iLoop As Integer, szaIntRate() As String
   Dim iNumOfdays As Integer
   Dim adoRst As New ADODB.Recordset, adoInt As New ADODB.Recordset

   szSQL = "SELECT TransactionID, DDate, OSAmount, IntCalDate " & _
           "FROM tlbReceipt " & _
           "WHERE ReceiptView = TRUE AND " & _
               "(DATEDIFF('D',DDate,DATE())-1)>= " & iDays & " AND Type = 1 AND " & _
               "SageAccountNumber = '" & szSAN & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

   szSQL = "SELECT InterestRates.BaseRate, InterestRates.AdditionalRate, " & _
               "InterestRates.DateFrom " & _
           "FROM InterestRates, Units " & _
           "WHERE Units.UnitNumber = '" & szUnitNum & "' AND " & _
               "Units.PropertyID = InterestRates.PropertyID AND " & _
               "InterestRates.Active = TRUE " & _
           "ORDER BY RateID;"
   adoInt.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   ReDim szaIntRate(adoInt.RecordCount, 1) As String

   While Not adoInt.EOF
      szaIntRate(iLoop, 0) = Format(adoInt!dateFrom, "dd/mmmm/yyyy")
      szaIntRate(iLoop, 1) = CStr(CCur(adoInt!BaseRate) + IIf(sIntRate = 0, CCur(adoInt!AdditionalRate), sIntRate))
      iLoop = iLoop + 1
      adoInt.MoveNext
   Wend
   szaIntRate(iLoop, 0) = Format(Date, "dd/mmmm/yyyy")
   adoInt.MoveFirst
   iLoop = 0

   CalAutoInterest = 0
   cTotalAmt = 0
   sDaysCalculated = 0

   While Not adoRst.EOF
      If adoRst!IntCalDate = "" Or IsNull(adoRst!IntCalDate) Then            'NO INT HAS BEEN CALCULATED BEFORE
         dtStartDt = IIf(DateDiff("d", CDate(adoRst!dDate), CDate(adoInt!dateFrom)) >= 0, CDate(adoInt!dateFrom), CDate(adoRst!dDate))
         For iLoop = 1 To adoInt.RecordCount
            iNumOfdays = DateDiff("d", dtStartDt, szaIntRate(iLoop, 0))
            If iNumOfdays >= 0 Then
               CalAutoInterest = CalAutoInterest + CCur(adoRst!OSAmount) * (szaIntRate(iLoop - 1, 1) / 100) * (iNumOfdays - 1) / 365
               dtStartDt = CDate(szaIntRate(iLoop, 0))
               cTotalAmt = cTotalAmt + CCur(adoRst!OSAmount)
               sDaysCalculated = sDaysCalculated + iNumOfdays
            End If
         Next iLoop
      Else
         dtStartDt = CDate(adoRst!IntCalDate)
         For iLoop = 1 To adoInt.RecordCount
            iNumOfdays = DateDiff("d", dtStartDt, szaIntRate(iLoop, 0))
            If iNumOfdays >= 0 Then
               CalAutoInterest = CalAutoInterest + CCur(adoRst!OSAmount) * (szaIntRate(iLoop - 1, 1) / 100) * (iNumOfdays - 1) / 365
               dtStartDt = CDate(szaIntRate(iLoop, 0))
               cTotalAmt = cTotalAmt + CCur(adoRst!OSAmount)
               sDaysCalculated = sDaysCalculated + iNumOfdays
            End If
         Next iLoop
      End If

      adoRst!IntCalDate = Format(Date, "dd/mmmm/yyyy")
      adoRst.Update
      adoRst.MoveNext
   Wend

   adoRst.Close
   Set adoRst = Nothing
End Function
'
'Private Function FindNextDueDate(dtNtDueDate As Date, iFreq As Integer, szDemandType As String, szPropertyID As String, adoConn As ADODB.Connection) As Date
''     this method is also in demand form. therefore, if there is any change here then change in demand form too.
'   Call GetGlobalDataPropertyWise(szPropertyID, adoConn, szDemandType)
'
'   If szGDYearly = "AUTOMATIC" Then
'      Select Case iFreq
'         Case 1:                                                   'Weekly in advance
'            FindNextDueDate = dtNtDueDate
'         Case 2:                                                   'Weekly in arrears
'            FindNextDueDate = DateAdd("d", 7, dtNtDueDate)
'         Case 3:                                                   'Fortnightly in advance
'            FindNextDueDate = DateAdd("d", 14, dtNtDueDate)
'         Case 4:                                                   'Fortnightly in arrears
'            FindNextDueDate = DateAdd("d", 14, dtNtDueDate)
'         Case 5:                                                   'Monthly in advance
'            FindNextDueDate = DateAdd("m", 1, dtNtDueDate)
'         Case 6:                                                   'Monthly in arrears
'            FindNextDueDate = DateAdd("m", 1, dtNtDueDate)
'         Case 7:                                                   'Quarterly in advance
'            FindNextDueDate = DateAdd("m", 3, dtNtDueDate)
'         Case 8:                                                   'Quarterly in arrears
'            FindNextDueDate = DateAdd("m", 3, dtNtDueDate)
'         Case 9:                                                   'Half yearly in advance
'            FindNextDueDate = DateAdd("m", 6, dtNtDueDate)
'         Case 10:                                                  'Half yearly in arrears
'            FindNextDueDate = DateAdd("m", 6, dtNtDueDate)
'         Case 11:                                                  'yearly in advance
'            FindNextDueDate = DateAdd("m", 12, dtNtDueDate)
'         Case 12:                                                  'yearly in arrears
'            FindNextDueDate = DateAdd("m", 12, dtNtDueDate)
'         Case 13:                                                  'Daily
'            FindNextDueDate = DateAdd("d", 1, dtNtDueDate)
'         Case 14:                                                  '4 Weekly in Advance
'            FindNextDueDate = DateAdd("d", 28, dtNtDueDate)
'         Case 15:                                                  '4 Weekly in Arrears
'            FindNextDueDate = DateAdd("d", -28, dtNtDueDate)
'         Case 16:                                                  '4 Monthly in Advance
'            FindNextDueDate = DateAdd("m", 4, dtNtDueDate)
'         Case 17:                                                  '4 Monthly in Arrears
'            FindNextDueDate = DateAdd("m", -4, dtNtDueDate)
'      End Select
'
'      Exit Function
'   End If
'
'   Select Case iFreq
'      Case 1:                                                   'Weekly in advance
'         FindNextDueDate = dtNtDueDate
'      Case 2:                                                   'Weekly in arrears
'         FindNextDueDate = DateAdd("d", 7, dtNtDueDate)
'      Case 3:                                                   'Fortnightly in advance
'         FindNextDueDate = dtNtDueDate
'      Case 4:                                                   'Fortnightly in arrears
'         FindNextDueDate = DateAdd("d", 14, dtNtDueDate)
'      Case 5:                                                   'Monthly in advance
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Monthly)
'      Case 6:                                                   'Monthly in arrears
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Monthly)
'      Case 7:                                                   'Quarterly in advance
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Quarterly)
'      Case 8:                                                   'Quarterly in arrears
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Quarterly)
'      Case 9:                                                   'Half yearly in advance
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Half_Yearly)
'      Case 10:                                                  'Half yearly in arrears
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Half_Yearly)
'      Case 11:                                                  'yearly in advance
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InAdv, Pay_Yearly)
'      Case 12:                                                  'yearly in arrears
'         FindNextDueDate = NextPayingDate(DateAdd("d", 1, dtNtDueDate), InArr, Pay_Yearly)
'      Case 13:                                                  'Daily
'         FindNextDueDate = DateAdd("d", 1, dtNtDueDate)
'      Case 14:                                                  '4 Weekly in Advance
'         FindNextDueDate = DateAdd("d", 28, dtNtDueDate)
'      Case 15:                                                  '4 Weekly in Arrears
'         FindNextDueDate = DateAdd("d", -28, dtNtDueDate)
'      Case 16:                                                  '4 Monthly in Advance
'         FindNextDueDate = DateAdd("m", 4, dtNtDueDate)
'      Case 17:                                                  '4 Monthly in Arrears
'         FindNextDueDate = DateAdd("m", -4, dtNtDueDate)
'   End Select
'End Function

Public Sub mnuDemandHistoryReport_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PrintDemandHistory.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Public Sub mnuDemandTransactionReport_Click()
   Load frmPreDemandTransactions
   frmPreDemandTransactions.Show
End Sub

Private Sub mnuDemandTypes_Click()
'   Load frmDCTypesPre
'   frmDCTypesPre.szMenu = "DEMAND_TYPE"
'   frmDCTypesPre.Show

   LoadForm frmDemandTypes
'   frmDemandTypes.Show
'   frmDemandTypes.ZOrder 0
End Sub

Public Sub mnuEBR_Click()
   LoadForm frmPreSCExpBudRpt
   frmPreSCExpBudRpt.LOOKUPCommand = ""
'   frmPreSCExpBudRpt.Left = 0
'   frmPreSCExpBudRpt.Top = 0
'   frmPreSCExpBudRpt.Show
'   frmPreSCExpBudRpt.ZOrder 0
End Sub

Public Sub mnuEBSR_Click()
   LoadForm frmPreSCExpBudSt
'   frmPreSCExpBudSt.Show
'   frmPreSCExpBudSt.ZOrder 0
End Sub

Private Sub mnuchartOfAccounts_Click()
    LoadForm frmNominalLedger
End Sub

Private Sub mnuEditUserNames_Click()
   LoadForm frmUsers
'   frmUsers.Show
'   frmUsers.ZOrder 0
End Sub
'
'Private Sub mnuElectricityCharge_Click()
'   Load frmElectricityUsage
'   frmElectricityUsage.Show
'End Sub

Public Sub mnuES_Click()
   LoadForm frmPreSCExpSt
'   frmPreSCExpSt.Show
'   frmPreSCExpSt.ZOrder 0
End Sub

Public Sub mnuExpLease_Click()
   Dim adoconn As New ADODB.Connection

'   connect to database
   adoconn.Open getConnectionString

   If PrintExpiredLeases(adoconn) Then
      ShowReport App.Path & szReportPath & "\RecentlyExpLease.rpt"
   End If

   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub mnuFund_Click()
   LoadForm frmFund
'   frmFund.Show
'   frmFund.ZOrder 0
End Sub

Public Sub mnuFundSummary_Click()
   LoadForm frmFundSumry
'   If Not frmFundSumry.bProceed Then
'      Unload frmFundSumry
'   Else
'      frmFundSumry.Show
'      frmFundSumry.ZOrder 0
'   End If
End Sub

Private Sub mnuGeneralLetters_Click()
   If IsLoadedAndVisible("frmSendLetters") Then Unload frmSendLetters

   frmSendLetters.szLetter = "GL"
   LoadForm frmSendLetters
'   frmSendLetters.Show
'   frmSendLetters.ZOrder 0
End Sub

Private Sub mnuGFY_Click()
   LoadForm frmFinancialYear
'   frmFinancialYear.Show
End Sub

Private Sub mnuGLU_Click()
   LoadForm frmGLU
'   frmGLU.Show
'   frmGLU.ZOrder 0
End Sub

Private Sub mnuImportTransaction_Click()
   Dim adoconn As New ADODB.Connection
   Dim szTable As String

   adoconn.Open getConnectionString

   On Error GoTo MissingTable

   szTable = "tblBankStatement"
   Rst1.Open "SELECT * FROM " & szTable & ";", adoconn, adOpenStatic, adLockReadOnly
   Rst1.Close

   adoconn.Close
   Set adoconn = Nothing

   LoadForm frmAutoBankReconciliation
'   frmAutoBankReconciliation.Show
'   frmAutoBankReconciliation.ZOrder 0
   Exit Sub

MissingTable:
   MsgBox "This company database is not up to date. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - " & szTable
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub mnuInsSchedule_Click()
   LoadForm frmInsScheduleCriteria
   
'   frmInsScheduleCriteria.Show
'   frmInsScheduleCriteria.ZOrder 0
End Sub

Private Sub mnuITBS_Click()
   LoadForm frmImpTransFromSt
'   frmImpTransFromSt.Show
'   frmImpTransFromSt.SetFocus
'   frmImpTransFromSt.ZOrder 0
End Sub

Private Sub mnuJBCR_Click()
   LoadForm frmPJRptCriteria
   frmPJRptCriteria.CallingFrom = "Budget"
'   frmPJRptCriteria.Show
'   frmPJRptCriteria.ZOrder 0
End Sub

Public Sub mnuLandlordSummaryStatement_Click()
   LoadForm frmLSS
'   frmLSS.Show
'   frmLSS.ZOrder 0
End Sub

Public Sub mnuLeaseDetails_Click()
   LoadForm frmLeaseReport
'   frmLeaseReport.Show
'   frmLeaseReport.ZOrder 0
End Sub

Public Sub mnuLesseeDetails_Click()
' Lessee details report
'issue 483 note 997 added by anol 24 Mar 2015
'Not showing all lessee address details. Input needs to be added that allows user to select client and property when running this report.
    LoadForm frmLDPre
'    frmLDPre.ZOrder 0
'   ShowReport App.Path & szReportPath & "\TenantDetailsReport.rpt"
End Sub

Public Sub mnuLesseeInfo_Click()
   ShowReport App.Path & szReportPath & "\TenantInfoReport.rpt"
End Sub

Private Sub mnuLesseePortal_Click()
   Dim r As Long
   
   r = ShellExecute(0, "open", "http://lesseeportal.prestige-property-software.co.uk", 0, 0, 1)
End Sub

Public Sub mnuLesseeStatement_Click()
'   Load frmLesseeStatement
'   frmLesseeStatement.Show
   LoadForm frmEmailLesseeSt
'   frmEmailLesseeSt.Show
'   frmEmailLesseeSt.ZOrder 0
End Sub

Public Sub mnuLofSL_Click()
   'issue 589 I am marking the leasedetail first on spare8 field
   'Dim adocon As New ADODB.Connection
   'adocon.Open getConnectionString
   'adocon.Execute "Update leasedetails set spare8=0"
   'adocon.Execute "Update leasedetails L INNER JOIN LRentCharges R ON L.LeaseID = R.LeaseID set L.spare8=1 where DateDiff ('d', {R.BRNextDueDate}, CDate({R.StopRC})) <= 0"
   'adocon.Execute "Update leasedetails L INNER JOIN LInsuranceCharges R ON L.LeaseID = R.LeaseID set L.spare8=1 where DateDiff ('d', {R.InsuranceNextDueDate}, CDate({R.StopIC})) <= 0 "
   'adocon.Execute "Update leasedetails L INNER JOIN LServiceCharges R ON L.LeaseID = R.LeaseID set L.spare8=1 where DateDiff ('d', {R.SCNextDueDate}, CDate({R.StopSC})) "
   ShowReport App.Path & szReportPath & "\StopLeases1.rpt"
   'adocon.Close
   'Set adocon = Nothing
End Sub



Private Sub mnuLogOut_Click()
   If MsgBox("Are you sure you wish to log out?", vbYesNo + vbQuestion, "Log Out") = vbNo Then
       Exit Sub
   Else
      If IsLoadedAndVisible("frmReport") Then
         MsgBox "There are open reports found. Please must close all open reports before login another company.", vbCritical + vbOKOnly, "Login"
         Exit Sub
      End If
      'issue 625
            Leasee1_LesseList_isUptoDate = False
            Leasee4_LesseList_isUptoDate = False
            frmDemand3_LesseList_isUptoDate = False
            frmSupplier_SupplierList_isUptoDate = False
            frmSupplier_SupplierListBCL_isUptoDate = False
            
      bLogOFF = True
      Load frmLogin2
      LastTenBackup
      Unload Me
      frmLogin2.Show
   End If
End Sub 'mnuLogout_Click()

Private Sub mnuExit_Click()
   Unload frmMMain
End Sub 'mnuExit_Click()

Private Sub mnuLSummary_Click()
   LoadForm frmLeaseViewSummary
'   frmLeaseViewSummary.Show
'   frmLeaseViewSummary.ZOrder 0
End Sub

Private Sub mnuManagementFees_Click()
   'issue 496 Management Fees
   'Added by anol 13 Nov 2014
'   Dim strTemp As String
'   strTemp = isControlAccountSet
'   If Len(strTemp) > 0 Then
'      MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'      Exit Sub
'   End If
   LoadForm frmManagementFees
'   frmFees.Show
'   frmFees.ZOrder 0
End Sub

Private Sub mnuNHR_Click()
   LoadForm frmNLAnalysis
   frmNLAnalysis.LOOKUPparam = ""
'   frmNLAnalysis.Show
'   frmNLAnalysis.ZOrder 0
End Sub

Private Sub mnuNominalLedger_Click()
'   LoadForm frmNominalLedger1
'   frmNominalLedger1.Show
'   frmNominalLedger1.ZOrder 0
End Sub

Private Sub mnuPayableTypes_Click()
'   Load frmDCTypesPre
'   frmDCTypesPre.szMenu = "PAYABLE_TYPE"
'   frmDCTypesPre.Show
'   frmDCTypesPre.ZOrder 0
'
   LoadForm frmRentPayableTypes
'   frmRentPayableTypes.Show
'   frmRentPayableTypes.ZOrder 0
End Sub

Private Sub mnuPayPro_Click()
   If IsLoadedAndVisible("frmBACSFiles") Then
      If frmBACSFiles.Caption <> "Payment Processed" Then
         MsgBox "Please close the 'View BACS File' window.", vbInformation + vbOKOnly, "Payment Processed"
         Exit Sub
      End If
   End If

   On Error GoTo MissingTable_GlobalRC

   Conn1.Open getConnectionString

   Rst1.Open "SELECT * FROM tblBatchPayment;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Rst1.Open "SELECT * FROM tblBatchTransaction;", Conn1, adOpenStatic, adLockReadOnly
   Rst1.Close

   Conn1.Close
   Set Conn1 = Nothing

   Load frmBACSFiles
   frmBACSFiles.cmdPayPro.Visible = True
   frmBACSFiles.Caption = "Payment Processed"
   frmBACSFiles.Show
   frmBACSFiles.ZOrder 0

   Exit Sub
MissingTable_GlobalRC:
   MsgBox "The database is not prepared for batch payment. Please contact PCM Consulting Ltd.", vbInformation + vbOKOnly, "Database - tblBatchPayment, tblBatchTransaction"
   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub mnuPCF_Click()
   LoadForm frmPCF
'   frmPCF.Show
'   frmPCF.ZOrder 0
End Sub

Private Sub mnuPendingJobs_Click()
   LoadForm frmJobReports
   'frmJobReports.Show
End Sub

Private Sub mnuPJR_Click()
   LoadForm frmPJRptCriteria
'   frmPJRptCriteria.CallingFrom = "Pending"
'   frmPJRptCriteria.Show
'   frmPJRptCriteria.ZOrder 0
End Sub

Private Sub mnuProfitLoss_Click()
   LoadForm frmPreProfit_n_Loss
'   frmPreProfit_n_Loss.Show
End Sub

Public Sub mnuPurchaseHistoryReport_Click()
'   Dim adoConn As New ADODB.Connection
   Dim i As Integer, szMY_ID As String
   Dim rep As frmReport
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'   adoConn.Open getConnectionString
'
'   adoConn.Execute "UPDATE tblPurInv SET Prn = 'N';"
'
'   szMY_ID = ""
'   For i = 1 To flxPurchHistory.Rows - 2
'      If flxPurchHistory.RowHeight(i) > 0 Then _
'         szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "', "
'   Next i
'   szMY_ID = szMY_ID + "'" + flxPurchHistory.TextMatrix(i, 0) + "'"
'
'   adoConn.Execute "UPDATE tblPurInv SET Prn = 'Y' WHERE MY_ID IN (" & szMY_ID & ");"
'   adoConn.Close
'   Set adoConn = Nothing

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PI_List1.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue "Y"

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
End Sub

Public Sub mnuPurchaseTransactionReport_Click()
   LoadForm frmPrePurchaseTransactions
'   frmPrePurchaseTransactions.Show
'   frmPrePurchaseTransactions.ZOrder 0
End Sub

Public Sub mnuRAS_Click()
   LoadForm frmRAS
'   frmRAS.Show
'   frmRAS.ZOrder 0
End Sub

Public Sub mnuRCC_Click()
   LoadForm frmRCC
'   frmRCC.Show
'   frmRCC.ZOrder 0
End Sub

Public Sub mnuReceiptAnalysis_Click()
   LoadForm frmViewReceiptDtRange
'   frmViewReceiptDtRange.Show
'   frmViewReceiptDtRange.ZOrder 0
End Sub

Private Sub mnuRefreshTree_Click()
   MousePointer = vbHourglass

   tvwLandLord.Nodes.Clear

   Conn1.Open getConnectionString

   Rst1.Open "SELECT ClientID FROM Client", Conn1, adOpenStatic, adLockReadOnly

   While Not Rst1.EOF
      DrawLandLordTree tvwLandLord, imgList, Rst1!ClientID, False, ""
      Rst1.MoveNext
   Wend
   Rst1.Close
   Conn1.Close
   MousePointer = vbDefault
End Sub

Private Sub mnuReminderLetters_Click()
   If IsLoadedAndVisible("frmSendLetters") Then Unload frmSendLetters

   frmSendLetters.szLetter = "RL"
   LoadForm frmSendLetters
'   frmSendLetters.Show
'   frmSendLetters.ZOrder 0
End Sub

Private Sub mnuReminderTemplates_Click()
   frmTemplate.szLetter = "RT"
   LoadForm frmTemplate
'   frmTemplate.Show
'   frmTemplate.ZOrder 0
End Sub

Private Sub mnuRentPayable_Click()
   'issue 496  Rent Payable
   'Added by anol 13 Nov 2014
'   Dim strTemp  As String
'   strTemp = isControlAccountSet
'   If Len(strTemp) > 0 Then
'      MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
'      Exit Sub
'   End If
   LoadForm frmRentPayable
'   frmRentPayable.Show
'   frmRentPayable.ZOrder 0
End Sub

Public Sub mnuRentReceivedReport_Click()
   LoadForm frmRRR
'   frmRRR.Show
'   frmRRR.ZOrder 0
End Sub

Public Sub mnuRentReviews_Click()
   LoadForm frmPreDtRange
   'frmPreDtRange.Show
End Sub

Private Sub mnuReportCategories_Click()
   LoadForm frmRptCategory
'   frmRptCategory.Show
'   frmRptCategory.ZOrder 0
End Sub

Private Sub mnuNominalHistoryReport_Click()
    LoadForm frmNLAnalysis
End Sub

Private Sub mnuNominalJournal_Click()
    LoadForm frmPurchaseExpense
End Sub

Private Sub mnuNominalListing_Click()
    LoadForm frmNominalLedger
End Sub

Private Sub mnuProfitAndLoss_Click()
    mnuProfitLoss_Click
End Sub

Private Sub mnuRestoreDB_Click()
   If MsgBox("Please ensure all other users are logged out of Prestige.", vbYesNo + vbInformation, "Data Restore") = vbNo Then Exit Sub
   If MsgBox("WARNING: NEWLY ADDED DATA WILL BE LOST.", vbYesNo + vbInformation, "Data Restore") = vbNo Then Exit Sub

   If RestoreDB Then
      Dim Conn1 As New ADODB.Connection
      Conn1.Open getConnectionString
      Call AllUpdateFunctions(Conn1) 'added by anol 2023-08-15
      Conn1.Close
      MsgBox "Database restore has been successfull.", vbInformation, "Information"
   Else
      MsgBox "System could not restore the database successfully. Please try again.", vbInformation, "Information"
   End If
End Sub

Public Sub mnuRP_Click()     'Receipt Payment
   LoadForm frmRPCriteria
   frmRPCriteria.Show
   frmRPCriteria.ZOrder 0
End Sub

Private Sub mnuSchedule_Click()
   LoadForm frmSchedule
'   frmSchedule.Show
'   frmSchedule.ZOrder 0
End Sub

Private Sub mnuSCY_Click()
   LoadForm frmSCYE
'   frmSCYE.Show
'   frmSCYE.ZOrder 0
End Sub

Private Sub mnuSetting_Click()
   LoadForm frmOptions
'   frmOptions.Show
'   frmOptions.ZOrder 0
End Sub

Private Sub mnuSupplierPortal_Click()
   Dim r As Long
   
   r = ShellExecute(0, "open", "http://supplierportal.prestige-property-software.co.uk", 0, 0, 1)
End Sub

Private Sub mnuThirdParty_Click()
   LoadForm frmExp3rdParty
'   frmExp3rdParty.Show
'   frmExp3rdParty.ZOrder 0
End Sub

Private Sub mnuTreeView_Click()
'   If mnuTreeView.Checked = True Then
'      mnuTreeView.Checked = False
'      picTreeView.Width = 0
'      Picture2.Width = 0
'   Else
'      mnuTreeView.Checked = True
'      picTreeView.Width = 1600
'      Picture2.Width = 55
'   End If
End Sub

Private Sub mnuTrialBalance_Click()
     LoadForm frmPreTrialBalnce
'   Load frmPreTrialBalnce
'   frmPreTrialBalnce.Show
'   frmPreTrialBalnce.ZOrder 0
End Sub

Public Sub mnuUnitDetails_Click()
   ShowReport App.Path & szReportPath & "\UnitDetailsReport.rpt"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub 'cmdExit_Click()
'Private Function isControlAccountSet() As String
'     'Resolved by BOSL
'      'issue 496
'      'Added by anol 11 Nov 2014
'
'      'This function will check if control account is set for all client.
'      'It will return message containing the client name that does not have control accounts
'      Dim szSQL1 As String
'      Dim rsCheck As New ADODB.Recordset
'      Dim adoConn1 As New ADODB.Connection
'
'      If adoConn1.State = 0 Then
'            adoConn1.Open getConnectionString
'      End If
'      Dim rsClient As New ADODB.Recordset
'      rsClient.Open "Select ClientID from Client", adoConn1, adOpenStatic, adLockReadOnly
'      While rsClient.EOF = False
'
'          szSQL1 = "SELECT * FROM NominalLedger WHERE ClientID='" & rsClient("ClientID").Value & "'"
'          rsCheck.Open szSQL1, adoConn1, adOpenStatic, adLockReadOnly
'
'           rsCheck.Find ("CAName = 'Input VAT'"), , , 1
'            If rsCheck.EOF = True Then
'                     rsCheck.Close
'                     isControlAccountSet = isControlAccountSet & vbCrLf & rsClient("ClientID").Value
'                     GoTo XX:
'            End If
'
'            rsCheck.Find ("CAName = 'Purchase Ledger Control'"), , , 1
'            If rsCheck.EOF = True Then
'                     rsCheck.Close
'                     isControlAccountSet = isControlAccountSet & vbCrLf & rsClient("ClientID").Value
'                     GoTo XX:
'            End If
'
'            rsCheck.Find ("CAName = 'Retained Earnings'"), , , 1
'            If rsCheck.EOF = True Then
'                     rsCheck.Close
'                     isControlAccountSet = isControlAccountSet & vbCrLf & rsClient("ClientID").Value
'                     GoTo XX:
'            End If
'
'             rsCheck.Find ("CAName = 'Output VAT'"), , adSearchForward, 1
'            If rsCheck.EOF = True Then
'                     rsCheck.Close
'                     isControlAccountSet = isControlAccountSet & vbCrLf & rsClient("ClientID").Value
'                     GoTo XX:
'            End If
'
'            rsCheck.Find ("CAName = 'Sales Ledger Control'"), , , 1
'            If rsCheck.EOF = True Then
'                     rsCheck.Close
'                     isControlAccountSet = isControlAccountSet & vbCrLf & rsClient("ClientID").Value
'                    GoTo XX:
'            End If
'
'            rsCheck.Close
'            Set rsCheck = Nothing
'XX:
'            rsClient.MoveNext
'
'        Wend
'      If adoConn1.State = 1 Then
'            adoConn1.Close
'      End If
'   'End of modification
'End Function
Private Sub cmdDemands_Click()
'   If IsLoadedAndVisible("frmBRPreForm") Or _
'      IsLoadedAndVisible("frmBatchRpt") Then
'      MsgBox "Please close the batch receipt before open the demand module.", vbInformation + vbOKOnly, "Demands"
'      Exit Sub
'   End If
'
'   MousePointer = vbHourglass

   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   adoconn.Open getConnectionString

   adoRst.Open "SELECT COUNT(spare1) FROM DemandTypes WHERE spare1='';", adoconn, adOpenStatic, adLockReadOnly

   If adoRst.Fields.Item(0).Value > 0 Then
      adoRst.Close
      Set adoRst = Nothing
      adoconn.Close
      Set adoconn = Nothing
      ShowMsgInTaskBar "Bank has not been setup in the demand type", "Y", "N"
      Exit Sub
   End If
   adoRst.Close
   Set adoRst = Nothing
''   'issue 496 Demand
''   'Added by anol 13 Nov 2014
''         Dim strTemp As String
''         strTemp = isControlAccountSet
''         If Len(strTemp) > 0 Then
''            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
''            Exit Sub
''         End If
'New cascade function introduced to load forms by anol 2021-05-17
''   Load frmDemands3
''   If Not frmDemands3.FormLoad Then
''      Unload frmDemands3
''   Else
''      frmDemands3.Show
''      frmDemands3.ZOrder 0
''    'I have found this procedure on 20170423 which was making the operations slow I have rem it for that reason anol
'''      Call frmDemands3.UpdateLesseeAccountBalance(adoConn)  'Update all lessee's balance
''   End If
    LoadForm frmDemands3

   adoconn.Close
   Set adoconn = Nothing
   'MousePointer = vbDefault
End Sub

Private Sub cmdtenants_Click()
'   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
'      Load frmLeasee2
'      frmLeasee2.Show
'   Else
   LoadForm frmLeasee1
'   frmLeasee1.Show
'   frmLeasee1.ZOrder 0
'   End If
End Sub 'cmdtenants_Click()

Private Sub cmdUnits_Click()
    LoadForm frmUnits2
'    MousePointer = vbHourglass
'    Load frmUnits2
'    frmUnits2.Show
'    frmUnits2.ZOrder 0
'
'    MousePointer = vbDefault
End Sub 'cmdUnits_Click()

Private Sub cmdLease_Click()
    LoadForm frmLease4
'   Load frmLease4
'   If Not frmLease4.FormLoad Then
'      Unload frmLease4
'   Else
'      frmLease4.Show
'      frmLease4.ZOrder 0
'   End If
End Sub 'cmdLease_Click()

Private Sub cmdShopCentre_Click()
    LoadForm frmProperty2
'    Load frmProperty2
'    frmProperty2.Show
'    frmProperty2.ZOrder 0
End Sub 'cmdShopCentre_Click()

Private Sub cmdGlobal_Click()
        LoadForm frmGlobalx
'   If RibbonVersion Then
'      Load frmGlobalx
'      frmGlobalx.Show
'      frmGlobalx.ZOrder 0
'   Else
'      Load frmGlobal1
'      frmGlobal1.Show
'      frmGlobal1.ZOrder 0
'   End If
End Sub 'cmdGlobal_Click()

Private Sub printReport(szReportPath As String)
'    Declare the application object used to open the rpt file
   Dim crxApplication As New CRAXDRT.Application

'   Declare the report object
   Dim Report As CRAXDRT.Report

   Set Report = crxApplication.OpenReport(szReportPath, 1)
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Dim strInvoiceAmt As String
   Dim rep As frmReport

   Set rep = New frmReport
   rep.LoadReportViewer Report
End Sub

Private Sub mnuVat_Click()
   Load frmVat
   frmVat.Show
   frmVat.ZOrder 0
End Sub

Private Sub mnuVATDetailReport_Click()
   LoadForm frmPreVAT_Details
'   frmPreVAT_Details.Show
'   frmPreVAT_Details.ZOrder 0
End Sub

Private Sub mnuVATSummaryReport_Click()
   LoadForm frmPreVAT_Summary
   'frmPreVAT_Summary.Show
   'frmPreVAT_Summary.ZOrder 0
End Sub

Private Sub mnuViewPayEmails_Click()
   Dim szTemp     As String
   Dim szLine     As String

   szTemp = Dir$(App.Path & "\Logs\moved_PI.dat")

'   If szTemp <> "moved_PI.dat" Then
'      If MsgBox("Email archiving process is being updated." & Chr(13) & _
'                "Please make sure no one is using Prestige." & Chr(13) & _
'                "Do you want to proceed now?", vbQuestion + vbYesNo, "Email Archive") = vbYes Then
''        Move Email_PI.dat to server
'         szTemp = Dir$(DB_PATH & "\AllStuff\Logs\Email_PI.dat")
'
'         If szTemp <> "Email_PI.dat" Then
''            CreateNonExistsFolder DB_PATH & "\AllStuff\Logs"
'            Open DB_PATH & "\AllStuff\Logs\Email_PI.dat" For Append As #1
'         End If
''        Read the old log file
'         Open App.Path & "\Logs\Email_PI.dat" For Input As #2
'
''        Transfer data to the new location
'         While Not EOF(2)
'            Line Input #2, szLine
'            Print #1, szLine
'         Wend
'
'         Close #1
'         Close #2
'
''        Mark the system that archive has been transfered to server
'         Open App.Path & "\Logs\moved_PI.dat" For Output As #1
'         Close #1
'
'         MoveTempFolder2Server
'      Else
'         Exit Sub
'      End If
'   End If

   Load frmEmailPayRemitt
   frmEmailPayRemitt.Show
   frmEmailPayRemitt.ZOrder 0
End Sub
'
'Private Sub mnuYEServiceCharge_Click()
'   Load frmSCYRR
'   frmSCYRR.Show
'End Sub

Private Sub picTreeView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub picTreeView_Resize()
   tvwLandLord.Width = picTreeView.Width
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   g = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If g = True Then Me.picTreeView.Width = IIf(Me.picTreeView.Width + X < 0, 0, Me.picTreeView.Width + X)
   stbStatusBar.Panels(1).text = "Drag the bar to change the width of the tree view"

   Me.MousePointer = 9
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   g = False
End Sub

Private Sub stbStatusBar_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
   If Panel = "Calculator" Then
'      Load frmCalculator
'      frmCalculator.Top = Me.Height - frmCalculator.Height - 2190
'      frmCalculator.Left = Me.Width - frmCalculator.Width - 1905
'      frmCalculator.Show
      Dim retval

      retval = Shell("C:\WINDOWS\System32\calc.exe")
   End If
End Sub

Private Sub tmDisplayTimer_Timer()
   iSecCount = iSecCount + 1
   
   If iSecCount = 2 Then
      rtxtMessageDisplay.text = ""
      tmDisplayTimer.Enabled = False
   End If
End Sub

Private Sub tvwLandLord_DblClick()
   Dim szaKey() As String

   MousePointer = vbHourglass

   szaKey = Split(tvwLandLord.SelectedItem.key, "@")

   Select Case UCase(szaKey(1))
      Case "CLIENT":
         Load frmClientNew4
         frmClientNew4.LOAD_CLINT_CLIENTID = szaKey(0)
            LoadForm frmClientNew4
'         frmClientNew4.Show
'         frmClientNew4.ZOrder 0
      Case "PROPERTY":
         Load frmProperty2
         frmProperty2.LOAD_PROPERTY_PROPERTYID = szaKey(0)
         szaKey = Split(tvwLandLord.SelectedItem.Parent, " / ")
         frmProperty2.CLIENT_NAME = szaKey(0)
         LoadForm frmProperty2
'         frmProperty2.Show
'         frmProperty2.ZOrder 0
      Case "UNITS":
         Load frmUnits2
         frmUnits2.LOAD_UNIT_UNITID = szaKey(0)
         LoadForm frmUnits2
'         frmUnits2.Show
'         frmUnits2.ZOrder 0
      Case "TENANT":
         If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
            Load frmLeasee2
            szaKey = Split(szaKey(0), "$")
            frmLeasee2.LOAD_TENANT_TENANTID = szaKey(1)
            LoadForm frmLeasee2
'            frmLeasee2.Show
'            frmLeasee2.ZOrder 0
         Else
            Load frmLeasee1
            szaKey = Split(szaKey(0), "$")
            frmLeasee1.LOAD_TENANT_TENANTID = szaKey(1)
            LoadForm frmLeasee1
'            frmLeasee1.Show
'            frmLeasee1.ZOrder 0
         End If
      Case "LEASE":
            LoadForm frmLease4
'         frmLease4.Show
'         frmLease4.ZOrder 0
         frmLease4.LoadFormFromTree szaKey(0)
   End Select
   
   MousePointer = vbDefault
End Sub

Private Sub tvwLandLord_Expand(ByVal Node As MSComctlLib.Node)
'   MsgBox Node.key
'   Node.Children.
End Sub

Private Sub tvwLandLord_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 116 And Shift = 0 Then mnuRefreshTree_Click
End Sub

Private Sub tvwLandLord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   stbStatusBar.Panels(1).text = "Tree View of the system"

   Me.MousePointer = vbArrow
End Sub

Private Sub UpdatePI_ClientID()
   Dim szSQL      As String

   szSQL = "UPDATE tblPurInv AS I, Property AS P " & _
           "SET    I.CL_ID = P.ClientID " & _
           "WHERE  I.PropertyID = P.PropertyID AND " & _
                  "(ISNULL(I.CL_ID) OR I.CL_ID = '');"
   Conn1.Execute szSQL
End Sub

Private Sub LoadTestCode()
'   Load frmPurchaseExpense
'   frmPurchaseExpense.Show


'   frmNJ.cmdAddNew_Click
'   frmNJ_Entry.cmbClient.ListIndex = 0
'   frmNJ_Entry.cmbProperty.ListIndex = 0
'   frmNJ_Entry.txtDateFrom.SetFocus
'   frmNJ_Entry.txtTitle.text = "XXXX"
'   frmNJ_Entry.cmbNC.ListIndex = 0
'   frmNJ_Entry.cmbFund.ListIndex = 0
'   frmNJ_Entry.txtDr.SetFocus
End Sub


