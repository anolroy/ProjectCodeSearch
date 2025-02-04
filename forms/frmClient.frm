VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Client / Landlord"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   FillColor       =   &H00C0C000&
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   9825
   Begin VB.CommandButton Command6 
      Caption         =   "Reprint Statement"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   39
      Top             =   7680
      Width           =   1455
   End
   Begin TabDlg.SSTab tabAddNewClient 
      Height          =   3885
      Left            =   45
      TabIndex        =   40
      Top             =   45
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6853
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Client"
      TabPicture(0)   =   "frmClient.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contact Address"
      TabPicture(1)   =   "frmClient.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Bank Details"
      TabPicture(2)   =   "frmClient.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBankDetails"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdNewBack(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraAccDetails"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraAddNewBankInfo"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdNewSave"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdNewSave 
         BackColor       =   &H00F1F1F1&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   3855
         Left            =   -75000
         TabIndex        =   61
         Top             =   -60
         Width           =   9615
         Begin VB.Frame Frame19 
            BackColor       =   &H00F1F1F1&
            Caption         =   "Office:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3255
            Left            =   5040
            TabIndex        =   68
            Top             =   120
            Width           =   4575
            Begin VB.TextBox txtNewOffAdd3 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1320
               TabIndex        =   69
               Top             =   960
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAdd2 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   14
               Top             =   600
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffEmail 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1320
               TabIndex        =   18
               Top             =   2760
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAdd1 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   13
               Top             =   240
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffAddPC 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   15
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtNewOffTel 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   16
               Top             =   2040
               Width           =   3000
            End
            Begin VB.TextBox txtNewOffPos 
               BackColor       =   &H00FFFFEA&
               Height          =   330
               Left            =   1300
               TabIndex        =   17
               Top             =   2400
               Width           =   3000
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Email:"
               Height          =   195
               Left            =   75
               TabIndex        =   74
               Top             =   2760
               Width           =   885
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   75
               TabIndex        =   73
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code"
               Height          =   195
               Left            =   75
               TabIndex        =   72
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Telephone:"
               Height          =   195
               Left            =   75
               TabIndex        =   71
               Top             =   2040
               Width           =   810
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Position:"
               Height          =   195
               Left            =   75
               TabIndex        =   70
               Top             =   2400
               Width           =   600
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00F1F1F1&
            Caption         =   "Home:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   3255
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   4575
            Begin VB.TextBox txtNewHomeAdd1 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1365
               TabIndex        =   6
               Top             =   240
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeAdd3 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1365
               TabIndex        =   8
               Top             =   960
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomePC 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1365
               TabIndex        =   9
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox txtNewHomeAdd2 
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1365
               TabIndex        =   7
               Top             =   600
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeEmail 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   12
               Top             =   2640
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeTel 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   10
               Top             =   1920
               Width           =   3000
            End
            Begin VB.TextBox txtNewHomeMob 
               BackColor       =   &H00FFFFEA&
               Height          =   345
               Left            =   1365
               TabIndex        =   11
               Top             =   2280
               Width           =   3000
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Address:"
               Height          =   195
               Left            =   120
               TabIndex        =   67
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Post Code:"
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   1440
               Width           =   780
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Personal Email:"
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   2640
               Width           =   1080
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Tel:"
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile:"
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   2280
               Width           =   510
            End
         End
         Begin VB.CommandButton cmdNewNext 
            BackColor       =   &H00F1F1F1&
            Caption         =   "&Next >>"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3405
            Width           =   1455
         End
         Begin VB.CommandButton cmdNewBack 
            BackColor       =   &H00F1F1F1&
            Caption         =   "<< &Back"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3405
            Width           =   1455
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   3735
         Left            =   0
         TabIndex        =   55
         Top             =   -60
         Width           =   9615
         Begin VB.Frame Frame8 
            BackColor       =   &H00F1F1F1&
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   5175
            Begin VB.TextBox txtNewClientID 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   3
               Top             =   2505
               Width           =   1455
            End
            Begin VB.TextBox txtNewClinetName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFEA&
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1530
               TabIndex        =   2
               Top             =   1800
               Width           =   2535
            End
            Begin VB.ComboBox cboSageCustAcc 
               BackColor       =   &H00FFFFEA&
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   120
               Width           =   3550
            End
            Begin VB.ComboBox cboSageSupAcc 
               BackColor       =   &H00FFFFEA&
               Height          =   315
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   960
               Width           =   3550
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Supplier A/C:"
               Height          =   195
               Left            =   75
               TabIndex        =   60
               Top             =   960
               Width           =   1365
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sage Customer A/C:"
               Height          =   195
               Left            =   75
               TabIndex        =   59
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Name:"
               Height          =   195
               Left            =   75
               TabIndex        =   58
               Top             =   1800
               Width           =   465
            End
            Begin VB.Label Label27 
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Client/Landlord ID:"
               Height          =   195
               Left            =   75
               TabIndex        =   57
               Top             =   2520
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdNewNext 
            BackColor       =   &H00F1F1F1&
            Caption         =   "&Next >>"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtAddNewNote 
            BackColor       =   &H00FFFFEA&
            Height          =   3015
            Left            =   5325
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame fraAddNewBankInfo 
         BackColor       =   &H00F1F1F1&
         Caption         =   "Add New Bank Info:"
         Height          =   3495
         Left            =   -74040
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   5295
         Begin VB.TextBox txtAddNewBankName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtAddNewBankAdd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   33
            Top             =   1380
            Width           =   3255
         End
         Begin VB.TextBox txtAddNewBankAdd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   1740
            Width           =   3255
         End
         Begin VB.TextBox txtAddNewBankAdd3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   35
            Top             =   2100
            Width           =   3255
         End
         Begin VB.TextBox txtAddNewBankPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   36
            Top             =   2580
            Width           =   1335
         End
         Begin VB.CommandButton cmdBankInfoSave 
            BackColor       =   &H00F1F1F1&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox txtAddNewBankID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton cmdAddNewBankCancel 
            BackColor       =   &H00F1F1F1&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank/Building Socity:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   2580
            Width           =   780
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Socity ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "To create a new Bank info you need to enter bank name and unique ID."
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraAccDetails 
         BackColor       =   &H00F1F1F1&
         Caption         =   "Account Details:"
         Height          =   2895
         Left            =   -69360
         TabIndex        =   45
         Top             =   0
         Width           =   3975
         Begin VB.TextBox txtAddNewBankAcc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   1320
            Width           =   2400
         End
         Begin VB.TextBox txtAddNewBankSC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   2280
            Width           =   2400
         End
         Begin VB.TextBox txtAddNewBankACName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   480
            Width           =   2400
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   2280
            Width           =   750
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   1110
         End
      End
      Begin VB.CommandButton cmdNewBack 
         BackColor       =   &H00F1F1F1&
         Caption         =   "<< &Back"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -68880
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Frame fraBankDetails 
         BackColor       =   &H00F1F1F1&
         Caption         =   "Bank Details:"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   41
         Top             =   0
         Width           =   5295
         Begin VB.TextBox txtBankNewAdd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1020
            Width           =   3255
         End
         Begin VB.TextBox txtBankNewAdd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1380
            Width           =   3255
         End
         Begin VB.TextBox txtBankNewAdd3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1740
            Width           =   3255
         End
         Begin VB.TextBox txtBankNewPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank/Building Socity:"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   480
            Width           =   1530
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   2340
            Width           =   780
         End
         Begin MSForms.ComboBox cboBankList 
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Top             =   480
            Width           =   3225
            VariousPropertyBits=   746604571
            BackColor       =   16777194
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5689;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szTextBox As Control
Dim szCurrentBankID As String
Dim bLoadAggrements As Boolean

'Dim szEditBankID As String

Private Sub cboBankList_Click()
   Dim aBank() As String
   
   If cboBankList.text = "ADD NEW BANK" Then
      fraAddNewBankInfo.Top = fraBankDetails.Top
      fraAddNewBankInfo.Left = fraBankDetails.Left
      fraAddNewBankInfo.Visible = True
      fraAddNewBankInfo.ZOrder 0
      txtAddNewBankID.SetFocus
      fraAccDetails.Enabled = False
      cmdNewBack(1).Enabled = False
      cmdNewSave.Enabled = False
      
      Exit Sub
   End If
   
   aBank = Split(cboBankList.text, " / ")
   cboBankList.text = aBank(0)
   txtBankNewPC.text = aBank(1)
   szCurrentBankID = aBank(2)
   
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT BANK_ADDRESS1,BANK_ADDRESS2,BANK_ADDRESS3 " & _
           "FROM TLBBANK " & _
           "WHERE BANK_ID='" & aBank(2) & "';"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   szEditBankID = aBank(2)
   txtBankNewAdd1.text = rstBank!BANK_ADDRESS1
   txtBankNewAdd2.text = rstBank!BANK_ADDRESS2
   txtBankNewAdd3.text = rstBank!BANK_ADDRESS3
   
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   
End Sub

Private Sub cmdAddNewBankCancel_Click()
   fraAddNewBankInfo.Visible = False
   txtAddNewBankID.text = ""
   txtAddNewBankName.text = ""
   txtAddNewBankAdd1.text = ""
   txtAddNewBankAdd2.text = ""
   txtAddNewBankAdd3.text = ""
   txtAddNewBankPC.text = ""
   fraAccDetails.Enabled = True
   cmdNewBack(1).Enabled = True
   cmdNewSave.Enabled = True
End Sub

Private Sub cmdBankInfoSave_Click()
   If txtAddNewBankName.text = "" Then
      MsgBox "Bank name can't be empty.", vbCritical, "Bank Name Error"
      Exit Sub
   End If
   If txtAddNewBankAdd1.text = "" Then
      MsgBox "Please Provide first line of Bank Address.", vbCritical, "Bank Address Error"
      Exit Sub
   End If
   If txtAddNewBankPC.text = "" Then
      MsgBox "Bank Post Code can't be empty.", vbCritical, "Bank Post Code Error"
      Exit Sub
   End If
'
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String
'
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt
'
   szSQL = "SELECT * " & _
           "FROM tlbBank;"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
   With rstBank
      .AddNew
      !BANK_ID = txtAddNewBankID.text
      !BANK_NAME = txtAddNewBankName.text
      !BANK_ADDRESS1 = txtAddNewBankAdd1.text
      !BANK_ADDRESS2 = txtAddNewBankAdd2.text
      !BANK_ADDRESS3 = txtAddNewBankAdd3.text
      !BANK_POST_CODE = txtAddNewBankPC.text
      .Update
      .Close
   End With
   Set rstBank = Nothing
   conBank.Close
   Set conBank = Nothing
'
   cboBankList.RemoveItem cboBankList.ListCount - 1
   cboBankList.AddItem txtAddNewBankName.text & " / " & txtAddNewBankPC.text & " / " & txtAddNewBankID.text
   cboBankList.text = txtAddNewBankName.text
   szCurrentBankID = txtAddNewBankID.text
   txtBankNewAdd1.text = txtAddNewBankAdd1.text
   txtBankNewAdd2.text = txtAddNewBankAdd2.text
   txtBankNewAdd3.text = txtAddNewBankAdd3.text
   txtBankNewPC.text = txtAddNewBankPC.text

   fraAddNewBankInfo.Visible = False
   fraAccDetails.Enabled = True
   cmdNewBack(1).Enabled = True
   cmdNewSave.Enabled = True

   MsgBox "Bank information has been save successfully", vbOKOnly, "Success"
End Sub

Private Sub cmdNewBack_Click(Index As Integer)
   tabAddNewClient.Tab = Index
End Sub

Private Sub cmdNewNext_Click(Index As Integer)
   tabAddNewClient.Tab = Index + 1
End Sub

Private Sub cmdNewNext_GotFocus(Index As Integer)
   If Index = 0 Then
      If cboSageCustAcc.text = "" Then
         MsgBox "Please choose Client's Sage Customer Account.", vbOKOnly + vbCritical, "Error"
         cboSageCustAcc.SetFocus
         Exit Sub
      End If
      If cboSageSupAcc.text = "" Then
         MsgBox "Please choose Client's Sage Supplier Account.", vbOKOnly + vbCritical, "Error"
         cboSageSupAcc.SetFocus
         Exit Sub
      End If
      If txtNewClinetName.text = "" Then
         MsgBox "Please type Client's Name.", vbOKOnly + vbCritical, "Error"
         txtNewClinetName.SetFocus
         Exit Sub
      End If
      If txtNewClientID.text = "" Then
         MsgBox "Please type Client's ID.", vbOKOnly + vbCritical, "Error"
         txtNewClientID.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub cmdNewSave_Click()
   If cboBankList.text = "" Then
      MsgBox "Please choose a Bank.", vbCritical, "Bank Info Error"
      cboBankList.SetFocus
      Exit Sub
   End If
   If txtAddNewBankAcc.text = "" Then
      If MsgBox("Do you want to leave Bank Account number Blank?", vbYesNo + vbQuestion, "Bank Account") = vbNo Then
         txtAddNewBankAcc.SetFocus
         Exit Sub
      End If
   End If
   If txtAddNewBankSC.text = "" Then
      If MsgBox("Do you want to leave Bank Sort Code number Blank?", vbYesNo + vbQuestion, "Bank Account") = vbNo Then
         txtAddNewBankSC.SetFocus
         Exit Sub
      End If
   End If

   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset, rstBank As rdoResultset
   Dim szSQL As String

   On Error Resume Next

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt
   
   szSQL = "SELECT * " & _
           "FROM CLIENT;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   With rstClient
      .AddNew
      !CLIENTID = txtNewClientID.text
      !ClientName = txtNewClinetName.text
      !ClientAddressLine1 = txtNewHomeAdd1.text
      !ClientAddressLine2 = txtNewHomeAdd2.text
      !ClientAddressLine3 = txtNewHomeAdd3.text
      !ClientPostCode = txtNewHomePC.text
      !ClientOfficeEmail = txtNewOffEmail.text
      !ClientPersonalEmail = txtNewHomeEmail.text
      !ClientHomeTel = txtNewHomeTel.text
      !ClientMobile = txtNewHomeMob.text
'      !ClientOffice = txtNewOff.text
      !ClientOfficeAddressLine1 = txtNewOffAdd1.text
      !ClientOfficeAddressLine2 = txtNewOffAdd2.text
      !ClientOfficeAddressLine3 = txtNewOffAdd3.text
      !ClientOfficePostCode = txtNewOffAddPC.text
      !ClientOfficeTel = txtNewOffTel.text
      !ClientOfficePos = txtNewOffPos.text
      !LandLordSageCustAC = cboSageCustAcc.text
      !LandLordSageSuppAC = cboSageSupAcc.text
      !Note = txtAddNewNote.text
      !BANK_ID = szCurrentBankID

      .Update
      .Close
   End With
   Set rstClient = Nothing
   
   szSQL = "SELECT * " & _
           "FROM tlbClientBanks;"
   Set rstBank = conClient.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   With rstBank
      .AddNew
      !CLIENT_ID = txtNewClientID.text
      !BANK_ID = szCurrentBankID
      !Bank_AC_Name = txtAddNewBankACName.text
      !BANK_AC_NUM = txtAddNewBankAcc.text
      !BANK_SC = txtAddNewBankSC.text
      !DEFAULT_AC = True
      .Update
      .Close
   End With
   Set rstBank = Nothing

   conClient.Close
   Set conClient = Nothing

   MsgBox "New customer informations have been entered successfully", vbInformation + vbOKOnly, "Success"
   MsgBox "Please add Property and Unit information for this client.", vbInformation + vbOKOnly, "Extra Info"

   txtNewClientID.text = ""
   txtNewClinetName.text = ""
   txtAddNewNote.text = ""
   txtNewHomeAdd1.text = ""
   txtNewHomeAdd2.text = ""
   txtNewHomeAdd3.text = ""
   txtNewHomePC.text = ""
   txtNewHomeTel.text = ""
   txtNewHomeMob.text = ""
   txtNewHomeEmail.text = ""
   txtNewOffAdd1.text = ""
   txtNewOffAdd2.text = ""
   txtNewOffAdd3.text = ""
   txtNewOffAddPC.text = ""
   txtNewOffTel.text = ""
   txtNewOffPos.text = ""
   txtNewOffEmail.text = ""
   cboBankList.text = ""
   txtBankNewAdd1.text = ""
   txtBankNewAdd2.text = ""
   txtBankNewAdd3.text = ""
   txtBankNewPC.text = ""
   txtAddNewBankAcc.text = ""
   txtAddNewBankSC.text = ""
   
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Top = 850
   Me.Left = 850
'   bLoadAggrements = False

   tabAddNewClient.Tab = 0
'   ConfigurFlexGrid

   If cboSageCustAcc.ListCount = 0 Then
      SageCustomerAccCombo
      SageSupplierAccCombo
   End If
   If cboBankList.ListCount = 0 Then
      BankAccList
   End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmClientNew.Enabled = True
    Unload Me
End Sub

Private Sub BankAccList()
   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   'Get the record for the client id
   szSQL = "SELECT BANK_NAME,BANK_POST_CODE,BANK_ID " & _
           "FROM TLBBANK;"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   
   While Not rstBank.EOF
      cboBankList.AddItem rstBank!BANK_NAME & " / " & rstBank!BANK_POST_CODE & _
                           " / " & rstBank!BANK_ID
      rstBank.MoveNext
   Wend
   
   cboBankList.AddItem "ADD NEW BANK"

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub SageCustomerAccCombo()
'
   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oSalesRecord As SageDataObject120.SalesRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
'   oSDO.Workspaces.Clear
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oSalesRecord = oWS.CreateObject("SalesRecord")

         ' Move to the First Record
         oSalesRecord.MoveFirst
         Dim rRow As Integer
         For rRow = 1 To oSalesRecord.Count
            cboSageCustAcc.AddItem CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value) & _
                                  " \ " & _
                                  CStr(oSalesRecord.Fields.Item("NAME").Value)
            oSalesRecord.MoveNext
         Next rRow

         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   MsgBox "(pcm_002) The SDO generated the following error: " & oSDO.LastError.text
   Set oSalesRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub SageSupplierAccCombo()
'
   ' Error Handler
   On Error GoTo Error_Handler

   ' Declare Objects
   Dim oSDO As SageDataObject120.SDOEngine
   Dim oWS As SageDataObject120.Workspace
   Dim oPurchaseRecord As SageDataObject120.PurchaseRecord

   ' Declare Variables
   Dim szDataPath As String

   ' Create the SDOEngine Object
   Set oSDO = New SageDataObject120.SDOEngine

   ' Create the Workspace
   Set oWS = oSDO.Workspaces.Add("Example")

   'read datapath from registr
   szDataPath = GetSetting("PropertyManagement", "SageCompany", CompanyDatapath)
   If szDataPath = "" Then
      ' Select Company. The SelectCompany method takes the program install
      ' folder as a parameter
      szDataPath = oSDO.SelectCompany(sageDirPath)
      'Save company name in the registry
      SaveSetting "PropertyManagement", "SageCompany", CompanyDatapath, szDataPath
   Else
     ' Try to Connect - Will Throw an Exception if it Fails
      If oWS.Connect(szDataPath, sageUserName, sagePassword, "Example") Then

         Set oPurchaseRecord = oWS.CreateObject("PurchaseRecord")

         ' Move to the First Record
         oPurchaseRecord.MoveFirst
         Dim rRow As Integer
         For rRow = 1 To oPurchaseRecord.Count
            cboSageSupAcc.AddItem CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value) & _
                                 " \ " & _
                                 CStr(oPurchaseRecord.Fields.Item("NAME").Value)
            oPurchaseRecord.MoveNext
         Next rRow

         'Disconnect
         oWS.Disconnect
      End If
   End If

   ' Destroy Objects
   Set oPurchaseRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   MsgBox "(pcm_003) The SDO generated the following error: " & oSDO.LastError.text

   Set oPurchaseRecord = Nothing
   Set oWS = Nothing
   Set oSDO = Nothing
End Sub

Private Sub txtSageCustomerAC_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtSageSupplierAC_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtAddNewBankID_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClientID.text = ""

   If (Len(txtAddNewBankID.text) = 10 And KeyAscii <> 8) Or _
      KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtAddNewBankID_LostFocus()
   If txtAddNewBankID.text = "" Then
      MsgBox "Please provide Bank ID", vbCritical, "Error"
      Exit Sub
   End If

   If IsBankID(txtAddNewBankID.text) Then
      MsgBox "Bank ID already exits, Please provide another ID.", vbCritical, "Duplicate ID"
      txtAddNewBankID.SelStart = 0
      txtAddNewBankID.SelLength = Len(txtAddNewBankID.text)
      txtAddNewBankID.SetFocus
   End If
End Sub

Private Sub txtAddNewBankName_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtAddNewBankName.text = ""
   
   If Len(txtAddNewBankName.text) = 50 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtNewClientID_GotFocus()
   txtNewClientID.SelStart = 0
   txtNewClientID.SelLength = Len(txtNewClientID.text)
End Sub

Private Sub txtNewClientID_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClientID.text = ""

   If (Len(txtNewClientID.text) = 10 And KeyAscii <> 8) Or _
      KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtNewClinetName_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
   If KeyAscii = 27 Then txtNewClinetName.text = ""
End Sub

Private Sub txtNewClinetName_LostFocus()
   Dim szTemp As String, i As Integer
   
On Error GoTo ErrorHandler

   i = 1
   While Len(szTemp) < 10
      If Asc(UCase(Mid(txtNewClinetName.text, i, 1))) > 64 And Asc(UCase(Mid(txtNewClinetName.text, i, 1))) < 91 Then
         szTemp = szTemp + Mid(txtNewClinetName.text, i, 1)
      End If
      i = i + 1
   Wend
ErrorHandler:
   txtNewClientID = szTemp
End Sub

'Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   cmdSelected.Enabled = True
'End Sub

'Private Sub lblAddress_Click()
'   If lblAddress.Caption = "Details >>" Then
'      fraOfficeAdd.Visible = True
'      fraOfficeAdd.ZOrder 0
'      lblAddress.Caption = "Details <<"
'   Else
'      fraOfficeAdd.Visible = False
'      lblAddress.Caption = "Details >>"
'   End If
'End Sub
'
'Private Sub lblAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblAddress_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub lblHomeAddress_Click()
'   fraOfficeAdd.Visible = False
'   lblAddress.Caption = "Details >>"
'End Sub
'
'Private Sub lblHomeAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblHomeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblHomeAddress_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblHomeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub lblOfficeAddress_Click()
'   fraOfficeAdd.Visible = True
'   fraOfficeAdd.ZOrder 0
'End Sub
'
'Private Sub lblOfficeAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''   lblOfficeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblOfficeAddress_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
''   lblOfficeAddress.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub lblSave_Click()
'   MsgBox "Under Construction"
'End Sub
'
'Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSave.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblSave.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub lblTenantIDLink_Click()
'   MsgBox "under construction"
'End Sub
'
'Private Sub tabAddNewClient_Click(PreviousTab As Integer)
'   If tabAddNewClient.Tab = 2 Then
'      If cboBankList.ListCount = 0 Then BankAccList
'   End If
'
'   If tabAddNewClient.Tab = 1 Or tabAddNewClient.Tab = 2 Then
'      If cboSageCustAcc.text = "" Then
'         MsgBox "Plseas Choose Sage Customer Account ID."
'         tabAddNewClient.Tab = 0
'         cboSageCustAcc.SetFocus
'         Exit Sub
'      End If
'      If cboSageSupAcc.text = "" Then
'         MsgBox "Plseas Choose Sage Supplier Account ID."
'         tabAddNewClient.Tab = 0
'         cboSageSupAcc.SetFocus
'         Exit Sub
'      End If
'      If txtNewClientID.text = "" Then
'         MsgBox "Plseas type Client ID."
'         tabAddNewClient.Tab = 0
'         txtNewClientID.SetFocus
'         Exit Sub
'      End If
'      If txtNewClinetName.text = "" Then
'         MsgBox "Plseas type Client Name."
'         tabAddNewClient.Tab = 0
'         txtNewClinetName.SetFocus
'         Exit Sub
'      End If
'   End If
'End Sub
'
'Private Sub tabClient_Click(PreviousTab As Integer)
'   Select Case tabClient.Tab
'   Case Is = 1
'      If txtSageCust.text = "" And txtSageSupp.text = "" Then
'         MsgBox "Select a Customer first.", vbCritical, "No Customer"
'         tabClient.Tab = 0
'      End If
'   Case Is = 2
'      If Not bLoadAggrements And txtClientID.text <> "" Then
'         LoadAggrements txtClientID.text
'      End If
'   Case Is = 6
'      tabAddNewClient.Tab = 0
'      If cboSageCustAcc.ListCount = 0 Then
'         SageCustomerAccCombo
'         SageSupplierAccCombo
'      End If
'   Case Else
'      'clicked anyother tab except state above
'   End Select
'End Sub
'
'Private Sub LoadAggrements(szID As String)
'   On Error Resume Next
'   bLoadAggrements = True
'
'   Dim Conn As New RDO.rdoConnection
'   Dim Rst As rdoResultset
'   Dim szStr As String, szaTemp() As String
'
'   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn.CursorDriver = rdUseIfNeeded
'   Conn.EstablishConnection rdDriverNoPrompt
'
'   szStr = "SELECT * " & _
'         "FROM tlbAggreement " & _
'         "WHERE tlbAggreement.CLIENT_ID='" & szID & "';"
'   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
'
'   If Rst!AGG_DATE <> "" Then txtAggDt.text = Format(Rst!AGG_DATE, "DD/MM/YYYY")
'   If Rst!START_DATE <> "" Then txtAggStDt.text = Format(Rst!START_DATE, "DD/MM/YYYY")
'   If Rst!END_DATE <> "" Then txtAggEndDt.text = Format(Rst!END_DATE, "DD/MM/YYYY")
'   If Rst!REVIEW_DATE <> "" Then txtAggReviewDt.text = Format(Rst!REVIEW_DATE, "DD/MM/YYYY")
'   If Rst!NOTICE_DATE <> "" Then txtAggNoticeDt.text = Format(Rst!NOTICE_DATE, "DD/MM/YYYY")
'   If Rst!BGRP_DATE <> "" Then txtBGRPDt.text = Format(Rst!BGRP_DATE, "DD/MM/YYYY")
'
'   Rst.Close
'   Set Rst = Nothing
'
'   szStr = "SELECT CommissionType,CommissionAmt,BGRPayable " & _
'           "FROM CLIENT " & _
'           "WHERE ClientID='" & szID & "';"
'   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
'
'   If Rst!COMMISSIONTYPE = 2 Then
'      txtCommissionAmt.text = Format(Rst!COMMISSIONAMT, "0.00")
'      optFixed.Value = True
'   Else
'      txtCommissionAmt.text = Format(Rst!COMMISSIONAMT, "0.00") & "%"
'      optComPercReceivable.Value = Not Rst!COMMISSIONTYPE
''      optPercReveived.Value = Rst!COMMISSIONTYPE
'   End If
'
'   Rst.Close
'   Set Rst = Nothing
'   Conn.Close
'   Set Conn = Nothing
'End Sub

'Private Sub DrawLandLordTree()
'   Dim nodX As Node   ' Declare the object variable.
'   Dim i As Integer, j As Integer   ' Declare a counter variable.
'   Dim szProperty As String, szProArray() As String
'   Dim szUnits As String, szUtArray() As String
'   Dim szaUtNmID() As String, szaProNmID() As String
'   Dim szLLID As String
'
'   tvwLandLord.ImageList = imgList
'   szLLID = txtClientID.text + "@" + "LANDLORD"
'   'Landlord ID
'   Set nodX = tvwLandLord.Nodes.Add(, , szLLID, txtClientName.text + " / " + txtClientID.text, 1, 1)
'
'   'Collece all property ID
'   szProperty = LLPropertyList(txtClientID.text)
'   szProArray = Split(szProperty, " # ")
'
'   If szProArray(0) <> "NULL" Then
'      For i = 0 To UBound(szProArray) 'Property Loop
'         szaProNmID = Split(szProArray(i), " / ")
'         Set nodX = tvwLandLord.Nodes.Add(szLLID, tvwChild, szaProNmID(0) & "@" & "PROPERTY", szaProNmID(1), 2, 2)
'
'         'Collect all Units for current Property
'         szUnits = LLUnitList(szaProNmID(0))
'         szUtArray = Split(szUnits, " # ")
'         If szUtArray(0) <> "NULL" Then
'            For j = 0 To UBound(szUtArray) 'Unit Loop
'               szaUtNmID = Split(szUtArray(j), " / ")
'               Set nodX = tvwLandLord.Nodes.Add(szaProNmID(0) & "@" & "PROPERTY", tvwChild, szaUtNmID(0) & "@" & "UNIT", szaUtNmID(1), 3, 3)
'            Next j
'         End If
'      Next i
'   End If
'End Sub
'
'Private Sub tvwLandLord_Click()
'   Dim szaPremisisIDType() As String
'
'   szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
'   fraType.Caption = szaPremisisIDType(1)
'
'   PremisisImageLoader imgPremises, szaPremisisIDType(0), szaPremisisIDType(1)
''                                   ID                         TYPE
'   txtTVInfoName.text = tvwLandLord.SelectedItem.text
'
'   If szaPremisisIDType(1) = "PROPERTY" Then
'      PropertyDetails szaPremisisIDType(0)
'   End If
'   If szaPremisisIDType(1) = "UNIT" Then
'      UnitDetails szaPremisisIDType(0)
'   End If
'End Sub
''
'Private Sub PropertyDetails(szID As String)
'   Dim Conn As New RDO.rdoConnection
'   Dim Rst As rdoResultset
'   Dim szStr As String, szaTemp() As String
'
'   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn.CursorDriver = rdUseIfNeeded
'   Conn.EstablishConnection rdDriverNoPrompt
'
'   szStr = "SELECT * " & _
'         "FROM PROPERTY " & _
'         "WHERE PROPERTY.PROPERTYID='" & szID & "';"
'   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
'
'   txtTVInfoName.text = Rst!PropertyName
'   txtTVInfoAdd1.text = Rst!ProAddressLine1
'   txtTVInfoAdd2.text = Rst!ProAddressLine2
'   txtTVInfoAdd3.text = Rst!ProAddressLine3
'   txtTVInfoPC.text = Rst!PROPOSTCODE
'
'   fraOccupied.Enabled = False
'
'   Rst.Close
'   Set Rst = Nothing
'   Conn.Close
'   Set Conn = Nothing
'End Sub

'Private Sub UnitDetails(szID As String)
'   Dim Conn As New RDO.rdoConnection
'   Dim Rst As rdoResultset
'   Dim szStr As String, szaTemp() As String
'
'   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn.CursorDriver = rdUseIfNeeded
'   Conn.EstablishConnection rdDriverNoPrompt
'
'   szStr = "SELECT * " & _
'         "FROM UNITS " & _
'         "WHERE UNITS.UnitNumber='" & szID & "';"
'   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
'
'   If Rst.EOF Then
'      MsgBox "Error in Database, Please contact with vendor", vbCritical, "Serious Error"
'   Else
'      If Rst!UNITNAME <> "" Then txtTVInfoName.text = Rst!UNITNAME
'      If Rst!UnitAddressLine1 <> "" Then
'         txtTVInfoAdd1.text = Rst!UnitAddressLine1
'      Else
'         txtTVInfoAdd1.text = ""
'      End If
'      If Rst!UnitAddressLine2 <> "" Then
'         txtTVInfoAdd2.text = Rst!UnitAddressLine2
'      Else
'         txtTVInfoAdd2.text = ""
'      End If
'      If Rst!UnitAddressLine3 <> "" Then
'         txtTVInfoAdd3.text = Rst!UnitAddressLine3
'      Else
'         txtTVInfoAdd3.text = ""
'      End If
'      If Rst!UnitPostCode <> "" Then
'         txtTVInfoPC.text = Rst!UnitPostCode
'      Else
'         txtTVInfoPC.text = ""
'      End If
'      If Rst!OCCUPIED = "Y" Then
'         lblTenantIDLink.Caption = Rst!SageAccountNumber
'         lblTenantNameLink.Caption = IIf(IsNull(Rst!TenantCompanyName), "", Rst!TenantCompanyName)
'         Rst.Close
'         Conn.Close
'         Set Rst = Nothing
'         Set Conn = Nothing
''
'         szStr = LeaseDetails(szID)
'         If szStr = "NULL" Then
'            MsgBox "Please update lease information of this unit.", vbInformation + vbOKOnly, "Error"
'         Else
'            szaTemp = Split(szStr, " # ")
'
'            txtPreOccupiedFr.text = szaTemp(0)
'            txtPreOccupiedTo.text = szaTemp(1)
'            txtPreTenancyType.text = szaTemp(2)
'            txtPreRentRvw.text = szaTemp(3)
'         End If
'      Else
'         lblTenantIDLink.Caption = "NOT OCCUPIED"
'         lblTenantNameLink.Caption = "NOT OCCUPIED"
'         fraOccupied.Enabled = False
'      End If
'   End If
'End Sub
'
'Private Sub lblTenantIDLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblTenantIDLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblTenantIDLink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
''   MsgBox "hi"
'   lblTenantIDLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub lblTenantNameLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblTenantNameLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\hmove.cur")
'End Sub
'
'Private Sub lblTenantNameLink_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   lblTenantNameLink.MouseIcon = LoadPicture(App.Path + "\" + "Package1\harrow.cur")
'End Sub
'
'Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtAdd3_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub

'Private Sub txtAggDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtAggDt
'End Sub
'
'Private Sub txtAggEndDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtAggEndDt
'End Sub
'
'Private Sub txtAggNoticeDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtAggNoticeDt
'End Sub
'
'Private Sub txtAggReviewDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtAggReviewDt
'End Sub
'
'Private Sub txtAggStDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtAggStDt
'End Sub
'
'Private Sub txtBGRPDt_Click()
'   dtDate.Visible = True
'   dtDate.ZOrder 0
'   Set szTextBox = txtBGRPDt
'End Sub

'Private Sub txtClientID_Change()
''   cmdClientID.Default = True
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTID LIKE '" & Trim(txtClientID.text) & "%' " & _
'              "ORDER BY CLIENTNAME;"
'
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   Dim iRow As Integer
'   iRow = 1
'
'   flxSearchResult.Clear
'   flxSearchResult.Rows = 2
'   ConfigurFlexGrid
'   While Not rstClient.EOF
'      flxSearchResult.TextMatrix(iRow, 0) = rstClient!CLIENTID
'      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
'      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
'      rstClient.MoveNext
'      If Not rstClient.EOF Then flxSearchResult.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'
'   cmdSelected.Enabled = True
'End Sub
'
'Private Sub txtClientID_KeyPress(KeyAscii As Integer)
'   Dim iTxtLen As Integer
'
'   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'   If KeyAscii = 27 Then txtClientID.text = ""
'   iTxtLen = Len(txtClientID.text)
''   If iTxtLen = 0 Then cmdClientID.Default = False
'   If (iTxtLen = 10 And KeyAscii <> 8) Or KeyAscii = 32 Then KeyAscii = 0
'End Sub

'Private Sub txtClientName_Change()
''   cmdClientName.Default = True
''   cmdClientID.Default = True
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTNAME LIKE '" & Trim(txtClientName.text) & "%' " & _
'              "ORDER BY CLIENTNAME;"
'
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   Dim iRow As Integer
'   iRow = 1
'
'   flxSearchResult.Clear
'   flxSearchResult.Rows = 2
'   ConfigurFlexGrid
'   While Not rstClient.EOF
'      flxSearchResult.TextMatrix(iRow, 0) = rstClient!CLIENTID
'      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
'      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
'      rstClient.MoveNext
'      If Not rstClient.EOF Then flxSearchResult.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'
'   cmdSelected.Enabled = True
'End Sub
'
'Private Sub txtClientName_KeyPress(KeyAscii As Integer)
'   Dim iTxtLen As Integer
'
'   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'   If KeyAscii = 27 Then txtClientID.text = ""
'   iTxtLen = Len(txtClientName.text)
''   If iTxtLen = 0 Then cmdClientName.Default = False
'   If iTxtLen = 10 And KeyAscii <> 8 Then KeyAscii = 0
'End Sub
'
'Private Sub txtHomePC_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtHomePh_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtMobile_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub

'Private Sub txtOffAdd1_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffAdd2_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffAdd3_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffEmail_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffice_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffPC_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffPh_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtOffPos_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtPerEmail_KeyPress(KeyAscii As Integer)
'   If cmdContactEdit.Enabled Then KeyAscii = 0
'End Sub
'
'Private Sub txtSearch_Change()
'   cmdSearch.Default = True
'End Sub
'
'Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'   Dim iTxtLen As Integer
'
'   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
'   If KeyAscii = 27 Then txtSearch.text = ""
'   iTxtLen = Len(txtSearch.text)
'   If iTxtLen = 0 Then cmdSearch.Default = False
'   If iTxtLen = 10 And KeyAscii <> 8 Then KeyAscii = 0
'End Sub
'
'Private Sub OpenFile()
'   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   Conn1.CursorDriver = rdUseOdbc
'   Conn1.EstablishConnection rdDriverNoPrompt
'
'   SQLStr1 = "SELECT * FROM AttachedFile WHERE FileName= '" & cmbFiles & "' AND SageAccountNumber='" & cboSage.text & "' AND Tenant='" & cboTenants.text & "'"
'   Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
''                                          rdOpenStatic, rdConcurReadOnly)  use this options
'
'   Dim filePath As String
'   Dim Filename As String
'   filePath = Rst1!filePath
'   Filename = Rst1!Filename
'   filePath = Mid(filePath, 1, Len(filePath) - Len(Filename))
'   If Rst1.EOF Then MsgBox Rst1.RowCount
'
'   Rst1.Close
'   Conn1.Close
'   Dim errorLevel As Long
'
'   errorLevel = ShellExecute(hWnd, "open", Filename, vbNullString, filePath, SW_SHOW)
'   If errorLevel < 32 Then MsgBox "File has been moved from original location.", vbExclamation
'End Sub
            
'Private Sub ResetFields()
'   bLoadAggrements = False
'   txtCommissionAmt.text = ""
'   txtAggDt.text = ""
'   txtAggStDt.text = ""
'   txtAggEndDt.text = ""
'   txtAggReviewDt.text = ""
'   txtAggNoticeDt.text = ""
'   txtBGRPDt.text = ""
'
'   txtBankName.text = ""
'   txtBankAdd1.text = ""
'   txtBankAdd2.text = ""
'   txtBankAdd3.text = ""
'   txtBankPC.text = ""
'   txtBankAccount.text = ""
'   txtBankSC.text = ""
'
'   'REFRESH THE TREE FOR NEW LANDLORD
'   tvwLandLord.Nodes.Clear
'   txtTVInfoName.text = ""
'   txtTVInfoAdd1.text = ""
'   txtTVInfoAdd2.text = ""
'   txtTVInfoAdd3.text = ""
'   txtTVInfoPC.text = ""
'   txtPreOccupiedFr.text = ""
'   txtPreOccupiedTo.text = ""
'   txtPreTenancyType.text = ""
'   txtPreRentRvw.text = ""
'   lblTenantIDLink.Caption = ""
'   lblTenantNameLink.Caption = ""
'   imgPremises.Picture = LoadPicture("")
'End Sub
'
'Private Function DoBrowse() As String
'' Browse for a file
'    Dim obj As New CDialog
'    With obj                'find a graphics file
'        .Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNFileMustExist + cdlOFNExplorer
'        .Filter = "Graphics File|*.bmp;*.ico;*.emf;*.wmf;*.jpg;*.gif|All Files|*.*"
'        .DialogTitle = "Select a File"
'        .InitDir = Trim$(App.Path)
'        .lHwnd = Me.hWnd
'        .ShowOpen
'        If Not .Cancelled Then
'            DoBrowse = .Filename
'        Else
'            DoBrowse = "NONE"
'        End If
'    End With
'    Set obj = Nothing
'End Function

'Private Sub cmdAgEdit_Click()
'   fraAgreement.Enabled = True
'   fraCommission.Enabled = True
'   cmdAgSave.Enabled = True
'   cmdAgEdit.Enabled = False
'End Sub

'Private Sub cmdAgSave_Click()
'   fraAgreement.Enabled = False
'   fraCommission.Enabled = False
'   cmdAgSave.Enabled = False
'   cmdAgEdit.Enabled = True
'End Sub

'Private Sub cmdClientID_Click()
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'Get the record for the client id
'   szSQL = "SELECT * " & _
'           "FROM CLIENT, TLBBANK " & _
'           "WHERE CLIENTID = '" & Trim(txtClientID.text) & "' AND " & _
'               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   If rstClient.EOF Then
'      MsgBox "No such ID found in the Database", vbCritical, "Error"
'      rstClient.Close
'      conClient.Close
'      Set rstClient = Nothing
'      Set conClient = Nothing
'      Exit Sub
'   End If
'
'   txtClientName.text = rstClient!ClientName
'   txtSageCust.text = rstClient!LandLordSageCustAC
'   txtSageSupp.text = rstClient!LandLordSageSuppAC
'   txtAdd1.text = rstClient!ClientAddressLine1
'   txtAdd2.text = rstClient!ClientAddressLine2
'   txtAdd3.text = rstClient!ClientAddressLine3
'   txtHomePC.text = rstClient!ClientPostCode
'   txtHomePh.text = rstClient!ClientHomeTel
'   txtMobile.text = rstClient!ClientMobile
'   txtPerEmail.text = rstClient!ClientPersonalEmail
'   txtOffice.text = rstClient!ClientOffice
'   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
'   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
'   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
'   txtOffPC.text = rstClient!ClientOfficePostCode
'   txtOffPh.text = rstClient!ClientOfficeTel
'   txtOffPos.text = rstClient!ClientOfficePos
'   txtOffEmail.text = rstClient!ClientOfficeEmail
'   txtClinetNote.text = rstClient!Note
'
'   txtBankName.text = rstClient!BANK_NAME
'   txtBankAdd1.text = rstClient!BANK_ADDRESS1
'   txtBankAdd2.text = rstClient!BANK_ADDRESS2
'   txtBankAdd3.text = rstClient!BANK_ADDRESS3
'   txtBankPC.text = rstClient!BANK_POST_CODE
'   txtBankAccount.text = rstClient!BANK_AC_NUM
'   txtBankSC.text = rstClient!BANK_SC
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'End Sub

'Private Sub cmdClientName_Click()
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'Get the record for the client id
'   szSQL = "SELECT * " & _
'           "FROM CLIENT, TLBBANK " & _
'           "WHERE CLIENTNAME = '" & Trim(txtClientName.text) & "' AND " & _
'               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   If rstClient.EOF Then
'      MsgBox "No such ID found in the Database", vbCritical, "Error"
'      rstClient.Close
'      conClient.Close
'      Set rstClient = Nothing
'      Set conClient = Nothing
'      Exit Sub
'   End If
'
'   txtClientID.text = rstClient!CLIENTID
'   txtSageCust.text = rstClient!LandLordSageCustAC
'   txtSageSupp.text = rstClient!LandLordSageSuppAC
'   txtAdd1.text = rstClient!ClientAddressLine1
'   txtAdd2.text = rstClient!ClientAddressLine2
'   txtAdd3.text = rstClient!ClientAddressLine3
'   txtHomePC.text = rstClient!ClientPostCode
'   txtHomePh.text = rstClient!ClientHomeTel
'   txtMobile.text = rstClient!ClientMobile
'   txtPerEmail.text = rstClient!ClientPersonalEmail
'   txtOffice.text = rstClient!ClientOffice
'   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
'   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
'   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
'   txtOffPC.text = rstClient!ClientOfficePostCode
'   txtOffPh.text = rstClient!ClientOfficeTel
'   txtOffPos.text = rstClient!ClientOfficePos
'   txtOffEmail.text = rstClient!ClientOfficeEmail
'   txtClinetNote.text = rstClient!Note
'
'   txtBankName.text = rstClient!BANK_NAME
'   txtBankAdd1.text = rstClient!BANK_ADDRESS1
'   txtBankAdd2.text = rstClient!BANK_ADDRESS2
'   txtBankAdd3.text = rstClient!BANK_ADDRESS3
'   txtBankPC.text = rstClient!BANK_POST_CODE
'   txtBankAccount.text = rstClient!BANK_AC_NUM
'   txtBankSC.text = rstClient!BANK_SC
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'End Sub
'
'Private Sub cmdClinetAddAtch_Click()
'   Dim szImageFileName As String
'   szImageFileName = DoBrowse()
'   If szImageFileName <> "NONE" Then
'      If DoStoreInDB(szImageFileName) Then
'         MsgBox "SUCCESSFUL"
'      End If
'   Else
'      Exit Sub
'   End If
'End Sub

'Private Function DoStoreInDB(sTemp As String) As Boolean
'   Dim mDB As Database             'open once, close upon unload
'   Dim Rst As Recordset
'   Dim szStr As String, msDBNameFull As String
'   Dim fldLongBinary As Field
'   Dim fldFileName As Field
'   Dim obj As New SaveCreateFile.cStoreCreateFile
'
'   msDBNameFull = szPictureDBPath
'   Set mDB = OpenDatabase(msDBNameFull)
'   szStr = "SELECT * FROM TLBIMAGES;"
'   Set Rst = mDB.OpenRecordset(szStr)  'make sure there is at least one record and store the file name
'
'    With Rst
'      .AddNew         'add a one 'dummy' record if none
''      .Fields("MY_ID") = "xx"
'      .Fields("PREMISIS_ID") = txtClientID.text
'      .Fields("PREMISIS_TYPE") = "CLIENT"
'      .Fields("IMAGE_NAME") = cmbFiles.text
'      .Update         'update the table
'    End With
'    szStr = "SELECT * " & _
'            "FROM TLBIMAGES " & _
'            "WHERE IMAGE_NAME='" & cmbFiles.text & "' AND " & _
'            "TLBIMAGES.PREMISIS_ID='" & txtClientID.text & "';"
'    Set Rst = mDB.OpenRecordset(szStr)
'    With Rst
'        If Not .EOF Then
'            .MoveLast
'            Set fldFileName = .Fields![IMAGE_PATH]   'set reference to this field
'            Set fldLongBinary = .Fields![PREMISIS_IMAGE]  'set a reference to the file field
'            With obj                        'now call the dll
'                If .StoreFileIntoField(Rst, fldFileName, fldLongBinary, sTemp) Then  'call the dll
'                    DoStoreInDB = True
'                End If
'            End With
'        End If
'        .Close
'    End With
'
'    Set Rst = Nothing
'    Set fldFileName = Nothing
'    Set fldLongBinary = Nothing
'    Set obj = Nothing
'

' Store the file in a database field.
'   Dim Conn As New Adodc

'   Dim fldLongBinary As Field
'   Dim fldFileName   As Field
'   Dim obj As New SaveCreateFile.cStoreCreateFile
'
'   Conn.ConnectionString = getConnectionString
'   Conn.RecordSource = "SELECT * FROM TLBIMAGES;"
'   Conn.CommandType = adCmdText
'   Conn.Refresh
'
'   Set Rst = Conn.Recordset
'
'   Rst.AddNew
'   Rst.Update
'
'   szStr = "SELECT * FROM TLBIMAGES;"
'   Conn.RecordSource = szStr
'   Conn.CommandType = adCmdText
'   Conn.Refresh
'   Set Rst = Conn.Recordset
'
'   With Rst
'       If Not .EOF Then
''           .MoveLast
'           Set fldFileName = ![IMAGE_NAME]   'set reference to this field
'           Set fldLongBinary = !PREMISIS_IMAGE  'set a reference to the file field
'           With obj                        'now call the dll
'               If .StoreFileIntoField(Rst, fldFileName, fldLongBinary, sTemp) Then  'call the dll
'                   DoStoreInDB = True
'               End If
'           End With
'       End If
'       .Close
'   End With
'   Set Rst = Nothing
'   Set fldFileName = Nothing
'   Set fldLongBinary = Nothing
'   Set obj = Nothing
'End Function

'Private Sub cmdClose_Click()
'   frmMMain.fraCmdButton.Enabled = True
'   Unload Me
'End Sub
'
'Private Sub cmdContactEdit_Click()
'   cmdContactEdit.Enabled = False
'   cmdContactSave.Enabled = True
'End Sub
'
'Private Sub cmdContactSave_Click()
'   'At the end
'   cmdContactEdit.Enabled = True
'   cmdContactSave.Enabled = False
'End Sub

'Private Sub cmdEdit_Click()
'   LockedComponents False
'End Sub

'Private Sub LockedComponents(bStatus As Boolean)
''Bank Details
'   txtBankName.Locked = bStatus
'   txtBankAdd1.Locked = bStatus
'   txtBankAdd2.Locked = bStatus
'   txtBankAdd3.Locked = bStatus
'   txtBankPC.Locked = bStatus
''Account Number & sort code
'   txtBankAccountName.Locked = bStatus
'   txtBankAccount.Locked = bStatus
'   txtBankSC.Locked = bStatus
'
'   cmdEdit.Enabled = bStatus
'End Sub

'Private Sub cmdEditBank_Click()
'   Dim szaTemp() As String
'
'   fraAddNewBankInfo.Top = fraBankDetails.Top
'   fraAddNewBankInfo.Left = fraBankDetails.Left
'   fraAddNewBankInfo.Visible = True
'   fraAddNewBankInfo.ZOrder 0
'   txtAddNewBankID.SetFocus
'   fraAccDetails.Enabled = False
'   cmdNewBack(1).Enabled = False
'   cmdNewSave.Enabled = False
'
'   txtAddNewBankAdd1.text = txtBankNewAdd1.text
'   txtAddNewBankAdd2.text = txtBankNewAdd2.text
'   txtAddNewBankAdd3.text = txtBankNewAdd3.text
'   txtAddNewBankPC.text = txtBankNewPC.text
'
'   szaTemp = Split(cboBankList.text, " / ")
'   txtAddNewBankID.text = szaTemp(1)
'   txtAddNewBankName.text = szaTemp(0)
'End Sub
'
'Private Sub cmdOpenFile_Click()
'   If cmbFiles.text = "" Then
'       MsgBox "Select a file from list."
'   Else
'       OpenFile
'   End If
'End Sub
'
'Private Sub cmdSearch_Click()
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'CLREATE SQL QUERY ON OPTION BUTTON SELECTION
'   If optSearchID.Value Then
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTID LIKE '" & Trim(txtSearch.text) & "%'"
'   ElseIf optSearchName.Value Then
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CLIENTNAME LIKE '%" & Trim(txtSearch.text) & "%'"
'   Else
'      szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'              "FROM CLIENT " & _
'              "WHERE CILENTPOSTCODE LIKE '" & Trim(txtSearch.text) & "'"
'   End If
'
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   If rstClient.EOF Then
'      MsgBox "No search result found in the Database", vbCritical, "Error"
'      rstClient.Close
'      conClient.Close
'      Set rstClient = Nothing
'      Set conClient = Nothing
'      Exit Sub
'   End If
'
'   Dim iRow As Integer
'   iRow = 1
'
'   flxSearchResult.Clear
'   flxSearchResult.Rows = 2
'   ConfigurFlexGrid
'   While Not rstClient.EOF
'      flxSearchResult.TextMatrix(iRow, 0) = rstClient!CLIENTID
'      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
'      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
'      rstClient.MoveNext
'      If Not rstClient.EOF Then flxSearchResult.AddItem ""
'      iRow = iRow + 1
'   Wend
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'
'   cmdSelected.Enabled = True
'End Sub
'
'Private Sub cmdSelected_Click()
'   ResetFields
'
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
''   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   'Get the record for the client id
'   szSQL = "SELECT * " & _
'           "FROM CLIENT, TLBBANK " & _
'           "WHERE CLIENTID = '" & Trim(flxSearchResult.TextMatrix(flxSearchResult.RowSel, 0)) & "' AND " & _
'               "CLIENT.BANK_ID=TLBBANK.BANK_ID"
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
''Debug.Print szSQL
'   txtClientID.text = flxSearchResult.TextMatrix(flxSearchResult.RowSel, 0)
'   txtClientName.text = rstClient!ClientName
'   txtSageCust.text = rstClient!LandLordSageCustAC
'   txtSageSupp.text = rstClient!LandLordSageSuppAC
'   txtAdd1.text = rstClient!ClientAddressLine1
'   txtAdd2.text = rstClient!ClientAddressLine2
'   txtAdd3.text = rstClient!ClientAddressLine3
'   txtHomePC.text = rstClient!ClientPostCode
'   txtHomePh.text = rstClient!ClientHomeTel
'   txtMobile.text = rstClient!ClientMobile
'   txtPerEmail.text = rstClient!ClientPersonalEmail
'   txtOffice.text = rstClient!ClientOffice
'   txtOffAdd1.text = rstClient!ClientOfficeAddressLine1
'   txtOffAdd2.text = rstClient!ClientOfficeAddressLine2
'   txtOffAdd3.text = rstClient!ClientOfficeAddressLine3
'   txtOffPC.text = rstClient!ClientOfficePostCode
'   txtOffPh.text = rstClient!ClientOfficeTel
'   txtOffPos.text = rstClient!ClientOfficePos
'   txtOffEmail.text = rstClient!ClientOfficeEmail
'   txtClinetNote.text = rstClient!Note
'
'   szCurrentBankID = rstClient!BANK_ID
'   txtBankAccountName.text = rstClient!Bank_AC_Name
'   txtBankName.text = rstClient!BANK_NAME
'   txtBankAdd1.text = rstClient!BANK_ADDRESS1
'   txtBankAdd2.text = rstClient!BANK_ADDRESS2
'   txtBankAdd3.text = rstClient!BANK_ADDRESS3
'   txtBankPC.text = rstClient!BANK_POST_CODE
'   txtBankAccount.text = rstClient!BANK_AC_NUM
'   txtBankSC.text = rstClient!BANK_SC
'
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'
'   DrawLandLordTree
'End Sub
'
'Private Sub Command12_Click()
'
'End Sub
'
'Private Sub cmdUpdate_Click()
'   If cmdEdit.Enabled = True Then Exit Sub
'SAVE DATA
'   If txtAddNewBankName.text = "" Then
'      MsgBox "Bank name can't be empty.", vbCritical, "Bank Name Error"
'      Exit Sub
'   End If
'   If txtAddNewBankAdd1.text = "" Then
'      MsgBox "Please Provide first line of Bank Address.", vbCritical, "Bank Address Error"
'      Exit Sub
'   End If
'   If txtAddNewBankPC.text = "" Then
'      MsgBox "Bank Post Code can't be empty.", vbCritical, "Bank Post Code Error"
'      Exit Sub
'   End If
'
'   Dim conBank As New RDO.rdoConnection
'   Dim rstBank As rdoResultset
'   Dim szSQL As String
''
'   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conBank.CursorDriver = rdUseIfNeeded
'   conBank.EstablishConnection rdDriverNoPrompt
''
'   szSQL = "UPDATE tlbBank " & _
'           "SET BANK_NAME = '" & txtBankName.text & "', " & _
'               "BANK_ADDRESS1 = '" & txtBankAdd1.text & "', " & _
'               "BANK_ADDRESS2 = '" & txtBankAdd2.text & "', " & _
'               "BANK_ADDRESS3 = '" & txtBankAdd3.text & "', " & _
'               "BANK_POST_CODE = '" & txtBankPC.text & "' " & _
'           "WHERE BANK_ID='" & szCurrentBankID & "'"
'   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
''
'   rstBank.Close
''
'   szSQL = "UPDATE Client " & _
'           "SET Bank_AC_Name = '" & txtBankAccountName.text & "', " & _
'               "BANK_AC_NUM = '" & txtBankAccount.text & "', " & _
'               "BANK_SC = '" & txtBankSC.text & "' " & _
'           "WHERE ClientID='" & Trim(txtClientID.text) & "'"
'   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
''
'   rstBank.Close
'   conBank.Close
'   Set rstBank = Nothing
'   Set conBank = Nothing
''
'   MsgBox "Bank information has been update successfully", vbOKOnly, "Success"
''
'   LockedComponents True
'End Sub
''
'Private Sub dtDate_DateClick(ByVal DateClicked As Date)
'   szTextBox.text = dtDate.Value
'   dtDate.Visible = False
'End Sub
'
'Private Sub flxSearchResult_DblClick()
'   cmdSelected_Click
'End Sub

'
'Private Sub PlacedAddressFrame()
'   fraOfficeAdd.Left = flxSearchResult.Left
'   fraOfficeAdd.Top = flxSearchResult.Top
'End Sub
'
'Private Sub LoadAllClientFlxGrd()
'   Dim conClient As New RDO.rdoConnection
'   Dim rstClient As rdoResultset
'   Dim szSQL As String
'
'   On Error Resume Next
'
'   'Set the RDO Connections to the dataset
'   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
'   conClient.CursorDriver = rdUseIfNeeded
'   conClient.EstablishConnection rdDriverNoPrompt
'
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTNAME;"
'
'   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
'
'   If rstClient.EOF Then GoTo NoRes
'
'   Dim iRow As Integer
'   iRow = 1
'
'   flxSearchResult.Clear
'   flxSearchResult.Rows = 2
'   ConfigurFlexGrid
'   While Not rstClient.EOF
'      flxSearchResult.TextMatrix(iRow, 0) = rstClient!CLIENTID
'      flxSearchResult.TextMatrix(iRow, 1) = rstClient!ClientName
'      flxSearchResult.TextMatrix(iRow, 2) = rstClient!ClientPostCode
'      rstClient.MoveNext
'      If Not rstClient.EOF Then flxSearchResult.AddItem ""
'      iRow = iRow + 1
'   Wend
'NoRes:
'   rstClient.Close
'   conClient.Close
'   Set rstClient = Nothing
'   Set conClient = Nothing
'
'   cmdSelected.Enabled = True
'End Sub
'
'Private Sub ConfigurFlexGrid()
'   flxSearchResult.ColWidth(0) = 1400
'   flxSearchResult.TextMatrix(0, 0) = "Client ID"
'
'   flxSearchResult.ColWidth(1) = 2000
'   flxSearchResult.TextMatrix(0, 1) = "Client Name"
'
'   flxSearchResult.ColWidth(2) = 1550
'   flxSearchResult.TextMatrix(0, 2) = "Post Code"
'End Sub
