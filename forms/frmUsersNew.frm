VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsersNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   7020
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8760
   Icon            =   "frmUsersNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabUsers 
      Height          =   6960
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8724
      _ExtentX        =   15399
      _ExtentY        =   12277
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   420
      BackColor       =   16777183
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Users"
      TabPicture(0)   =   "frmUsersNew.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameUserInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Company Access"
      TabPicture(1)   =   "frmUsersNew.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameRolePermission"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameUserInfo 
         BackColor       =   &H00FFFFDF&
         Caption         =   "User Details"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6672
         Left            =   0
         TabIndex        =   15
         Top             =   288
         Width           =   8652
         Begin VB.PictureBox picRole 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4230
            Left            =   1764
            ScaleHeight     =   4200
            ScaleWidth      =   6255
            TabIndex        =   40
            Top             =   2916
            Visible         =   0   'False
            Width           =   6285
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
               TabIndex        =   42
               Top             =   0
               Width           =   255
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRole 
               Height          =   3480
               Left            =   48
               TabIndex        =   41
               Top             =   708
               Width           =   6168
               _ExtentX        =   10874
               _ExtentY        =   6138
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
            Begin MSForms.TextBox txtSearchRoleName 
               Height          =   252
               Left            =   1620
               TabIndex        =   48
               Top             =   372
               Width           =   4620
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "8149;444"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtSearchRoletID 
               Height          =   252
               Left            =   48
               TabIndex        =   47
               Top             =   372
               Width           =   1536
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2709;444"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.Label lblRoleID 
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   120
               Width           =   1410
               VariousPropertyBits=   8388627
               Caption         =   "Role ID"
               Size            =   "2487;344"
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
               TabIndex        =   45
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lblFlxPayee 
               Caption         =   "EMPTY"
               Height          =   255
               Index           =   4
               Left            =   2115
               TabIndex        =   44
               Top             =   1200
               Width           =   1095
            End
            Begin MSForms.Label Label3 
               Height          =   195
               Left            =   1665
               TabIndex        =   43
               Top             =   120
               Width           =   1590
               VariousPropertyBits=   8388627
               Caption         =   "Role Name"
               Size            =   "2805;344"
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
               Index           =   15
               Left            =   48
               Top             =   72
               Width           =   5856
            End
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6264
            TabIndex        =   38
            Top             =   1980
            Width           =   1215
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   288
            TabIndex        =   0
            Top             =   1980
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            TabIndex        =   37
            Top             =   1980
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3312
            TabIndex        =   6
            Top             =   1980
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4752
            TabIndex        =   36
            Top             =   1980
            Width           =   1215
         End
         Begin VB.PictureBox fmeUserLookup 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   4116
            Left            =   36
            ScaleHeight     =   4080
            ScaleWidth      =   8550
            TabIndex        =   24
            Top             =   2520
            Width           =   8580
            Begin VB.TextBox txtSearchUser 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   50
               TabIndex        =   29
               Top             =   240
               Width           =   1020
            End
            Begin VB.TextBox txtSearchName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1110
               TabIndex        =   28
               Top             =   240
               Width           =   1560
            End
            Begin VB.TextBox txtSearchRole 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2700
               TabIndex        =   27
               Top             =   240
               Width           =   2190
            End
            Begin VB.TextBox txtSearchEmail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4944
               TabIndex        =   26
               Top             =   225
               Width           =   1632
            End
            Begin VB.TextBox txtSearchActive 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6615
               TabIndex        =   25
               Top             =   225
               Width           =   1908
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridUserLookup 
               Height          =   3516
               Left            =   36
               TabIndex        =   35
               Top             =   540
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   6191
               _Version        =   393216
               Cols            =   9
               FixedCols       =   0
               BackColorFixed  =   13553358
               ForeColorFixed  =   16777215
               BackColorSel    =   14737632
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               GridColor       =   -2147483638
               WordWrap        =   -1  'True
               HighLight       =   2
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
               _Band(0).Cols   =   9
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   6660
               TabIndex        =   34
               Top             =   36
               Width           =   396
            End
            Begin VB.Label lblSEmail 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   4896
               TabIndex        =   33
               Top             =   36
               Width           =   360
            End
            Begin VB.Label lblSRole 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Role"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   2748
               TabIndex        =   32
               Top             =   36
               Width           =   288
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   1140
               TabIndex        =   31
               Top             =   36
               Width           =   396
            End
            Begin VB.Label lblSUserID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User ID"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   60
               TabIndex        =   30
               Top             =   36
               Width           =   480
            End
            Begin VB.Shape Shape4 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   192
               Index           =   6
               Left            =   48
               Top             =   36
               Width           =   8472
            End
         End
         Begin VB.CommandButton cmdRole 
            Caption         =   ".."
            Enabled         =   0   'False
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
            Left            =   8064
            TabIndex        =   4
            Top             =   900
            Width           =   345
         End
         Begin MSForms.CheckBox chkActive 
            Height          =   300
            Left            =   2592
            TabIndex        =   39
            Top             =   468
            Width           =   1308
            BackColor       =   16777183
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2307;529"
            Value           =   "1"
            Caption         =   "Active"
            FontName        =   "Myriad Web"
            FontHeight      =   156
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label LabelUserID 
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   72
            TabIndex        =   23
            Top             =   540
            Width           =   828
         End
         Begin VB.Label LabelUsersName 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   72
            TabIndex        =   22
            Top             =   936
            Width           =   828
         End
         Begin MSForms.TextBox txtUserID 
            Height          =   312
            Left            =   936
            TabIndex        =   21
            Top             =   468
            Width           =   1524
            VariousPropertyBits=   746604571
            Size            =   "2688;550"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtUserName 
            Height          =   312
            Left            =   936
            TabIndex        =   1
            Top             =   900
            Width           =   3000
            VariousPropertyBits=   746604571
            Size            =   "5292;550"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtPassword 
            Height          =   312
            Left            =   936
            TabIndex        =   2
            Top             =   1368
            Width           =   3000
            VariousPropertyBits=   746604571
            Size            =   "5292;550"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblPass 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   72
            TabIndex        =   20
            Top             =   1404
            Width           =   828
         End
         Begin MSForms.TextBox txtConPass 
            Height          =   312
            Left            =   5400
            TabIndex        =   3
            Top             =   468
            Width           =   3000
            VariousPropertyBits=   746604571
            Size            =   "5292;550"
            PasswordChar    =   42
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2lblConPass 
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4032
            TabIndex        =   19
            Top             =   540
            Width           =   1296
         End
         Begin MSForms.TextBox txtRole 
            Height          =   312
            Left            =   5400
            TabIndex        =   18
            Top             =   900
            Width           =   3000
            VariousPropertyBits=   746604569
            Size            =   "5292;550"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblRole 
            BackStyle       =   0  'Transparent
            Caption         =   "Role"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4572
            TabIndex        =   17
            Top             =   936
            Width           =   828
         End
         Begin MSForms.TextBox txtEmail 
            Height          =   312
            Left            =   5400
            TabIndex        =   5
            Top             =   1296
            Width           =   2964
            VariousPropertyBits=   746604571
            Size            =   "5228;550"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblEMail 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4572
            TabIndex        =   16
            Top             =   1332
            Width           =   828
         End
      End
      Begin VB.Frame FrameRolePermission 
         BackColor       =   &H00FFFFDF&
         Caption         =   "Company Access Details"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6636
         Left            =   -74964
         TabIndex        =   8
         Top             =   288
         Width           =   8652
         Begin VB.CheckBox chkSelectAllCompany 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFDF&
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   264
            Left            =   36
            TabIndex        =   62
            Top             =   5868
            Width           =   1452
         End
         Begin VB.PictureBox picComUserSearch 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   4116
            Left            =   0
            ScaleHeight     =   4080
            ScaleWidth      =   8550
            TabIndex        =   49
            Top             =   648
            Visible         =   0   'False
            Width           =   8580
            Begin VB.CommandButton cmdComSearchPicClose 
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
               Left            =   8316
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   -36
               Width           =   255
            End
            Begin VB.TextBox txtComStatus 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6576
               TabIndex        =   54
               Top             =   225
               Width           =   1980
            End
            Begin VB.TextBox txtComEmail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4905
               TabIndex        =   53
               Top             =   225
               Width           =   1632
            End
            Begin VB.TextBox txtComRole 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2700
               TabIndex        =   52
               Top             =   240
               Width           =   2190
            End
            Begin VB.TextBox txtComUserName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1110
               TabIndex        =   51
               Top             =   240
               Width           =   1560
            End
            Begin VB.TextBox txtComUserId 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   50
               TabIndex        =   50
               Top             =   240
               Width           =   1020
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxComUserDetails 
               Height          =   3516
               Left            =   36
               TabIndex        =   55
               Top             =   540
               Width           =   8520
               _ExtentX        =   15028
               _ExtentY        =   6191
               _Version        =   393216
               Cols            =   9
               FixedCols       =   0
               BackColorFixed  =   13553358
               ForeColorFixed  =   16777215
               BackColorSel    =   14737632
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               GridColor       =   -2147483638
               WordWrap        =   -1  'True
               HighLight       =   2
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
               _Band(0).Cols   =   9
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label lblSUserID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User ID"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   60
               TabIndex        =   60
               Top             =   36
               Width           =   480
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   1140
               TabIndex        =   59
               Top             =   36
               Width           =   396
            End
            Begin VB.Label lblSRole 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Role"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   2748
               TabIndex        =   58
               Top             =   36
               Width           =   288
            End
            Begin VB.Label lblSEmail 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   4896
               TabIndex        =   57
               Top             =   36
               Width           =   360
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Myriad Web"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   6660
               TabIndex        =   56
               Top             =   36
               Width           =   396
            End
            Begin VB.Shape Shape4 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00C0FFFF&
               FillStyle       =   0  'Solid
               Height          =   192
               Index           =   0
               Left            =   48
               Top             =   36
               Width           =   8472
            End
         End
         Begin VB.CommandButton cmdUsersSearch 
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
            Height          =   372
            Left            =   5148
            TabIndex        =   12
            Top             =   252
            Width           =   345
         End
         Begin VB.CommandButton cmdCompanyAccess 
            Caption         =   "Save Access"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   456
            Left            =   6876
            TabIndex        =   9
            Top             =   5976
            Width           =   1644
         End
         Begin MSComctlLib.ListView lvwavailableCompany 
            Height          =   4932
            Left            =   36
            TabIndex        =   13
            Top             =   900
            Width           =   8556
            _ExtentX        =   15081
            _ExtentY        =   8705
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CompanyID"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Company Name"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label lblAvlCom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available Companies"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   36
            TabIndex        =   14
            Top             =   648
            Width           =   1392
         End
         Begin MSForms.TextBox txtUserAccessName 
            Height          =   384
            Left            =   1764
            TabIndex        =   11
            Top             =   252
            Width           =   3732
            VariousPropertyBits=   746604569
            Size            =   "6583;677"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblUserName 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   792
            TabIndex        =   10
            Top             =   324
            Width           =   828
         End
      End
   End
End
Attribute VB_Name = "frmUsersNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 21
Dim NEWMODE_ As Boolean
'Dim CurrentItem As MSComctlLib.ListItem

Private Sub chkSelectAllCompany_Click()

For i = 1 To lvwavailableCompany.ListItems.Count
If chkSelectAllCompany.Value = 1 Then
                 lvwavailableCompany.ListItems(i).Checked = True
                 chkSelectAllCompany.Caption = "DeSelect All"
                 Else
                 lvwavailableCompany.ListItems(i).Checked = False
                 chkSelectAllCompany.Caption = "Select All"
                 End If
            Next
End Sub

'Private Sub cmdAddCompany_Click()
''If lvwSelectedCompany.SelectedItems.Count = 0 Then
''MsgBox "No company selected"
''Exit Sub
''End If
''If Not CurrentItem Is Nothing Then
''        MsgBox "Currently selected item: " & CurrentItem.text
''    Else
''        MsgBox "No items selected"
''    End If
''Dim i As Integer
''For i = 0 To lvwSelectedCompany
'Dim newItem As ListItem
'
'    Set newItem = lvwSelectedCompany.ListItems.Add(, , lvwavailableCompany.SelectedItem.text)
'    newItem.SubItems(1) = lvwavailableCompany.SelectedItem.SubItems(1)
''   lvwavailableCompany.ListItems.Remove lvwavailableCompany.SelectedItem.Index
'End Sub

Private Sub cmdCancel_Click()
''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 11 disable enable button
ComponentInFrameEnableMode Me, FrameUserInfo, DefaultMode
'   If NEWMODE_ Then SageCustomerAccCombo cboSageAccountNumber

   NEWMODE_ = False
'   SEARCHTenantMODE_ = True

   txtUserName.Enabled = True
   txtPassword.Enabled = True
   txtConPass.Enabled = True
   txtRole.Enabled = False
   txtEmail.Enabled = True
'''   cmdRoleId.Enabled = True
'   cmdRoleId.Visible = True
cmdRole.Enabled = False
'Code change by mahboob 26/08/2023
gridUserLookup.Enabled = True
   If txtUserID.text = "" Then Exit Sub


   ComponentInFrameEnableMode Me, FrameUserInfo, DefaultMode

'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth

   ComponentInFrameClearMode Me, FrameUserInfo, ClearOnlyTextBoxes
   txtUserID.Locked = True
cmdRole.Enabled = False
End Sub

Private Sub cmdClose_Click()
''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 12 close the form
Unload Me
End Sub

Private Sub cmdCompanyAccess_Click()

'''''Modified by Mahboob 03/04/2023 Change ID 11 work item 11
'Check any company is selected or not
Dim isChecked As Integer
isChecked = 0
For i = 1 To lvwavailableCompany.ListItems.Count
        If lvwavailableCompany.ListItems(i).Checked = True Then

                isChecked = 1
        End If
    Next
    If isChecked = 0 Then
    MsgBox "Please select a company."
    Exit Sub
    End If
    Dim aboCnn As New ADODB.Connection
    aboCnn.Open getConnectionStringUserAccess
    Dim psql As String
    aboCnn.Execute "DELETE FROM UserCompanyAccess where UserID=" & txtUserAccessName.Tag & ""
'    aboCnn.Execute "Update UserNames set  UserPassword='" & txtPassword.text & "' WHERE UserID=" & txtUserID.text & ""
For i = 1 To lvwavailableCompany.ListItems.Count
        If lvwavailableCompany.ListItems(i).Checked = True Then
            'INSERT
                psql = ""
                psql = "INSERT INTO UserCompanyAccess(UserID,CompanyID)"
                psql = psql & " VALUES(" & txtUserAccessName.Tag & "," & lvwavailableCompany.ListItems(i).Tag & ")"
                aboCnn.Execute psql
        End If
    Next
    
MsgBox "Company access successfully saved"
aboCnn.Close
    Set aboCnn = Nothing
    'Clear the check boxes
'    ClearCompnayCheck
End Sub
Function ClearCompnayCheck()
'''''Modified by Mahboob 03/04/2023 Change ID 11 work item 12 uncheck the list view
i = 1
For i = 1 To lvwavailableCompany.ListItems.Count

            lvwavailableCompany.ListItems(i).Checked = False

    Next
End Function
Private Sub cmdComSearchPicClose_Click()
picComUserSearch.Visible = False
End Sub

Private Sub cmdEdit_Click()
''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 10 disable enable button
If txtUserID.text = "" Then
      MsgBox "Please select a User to continue.", vbInformation, "Edit User"
      Exit Sub
   End If
   NEWMODE_ = False
'   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, FrameUserInfo, EditMode
'txtPassword.Enabled = False
'   txtUserName.SetFocus
'   cmdRoleId.Enabled = False
'Code change by mahboob 27/08/2023
gridUserLookup.Enabled = False
If txtUserID.text = "1" Then
chkActive.Enabled = False
Else
chkActive.Enabled = True
End If
cmdRole.Enabled = True
End Sub
Function LoadRoleList()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 13 load role in grid
Dim rRow As Integer
     Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess
    Dim sqlCom As String
   txtSearchRoletID.text = ""
   txtSearchRoleName.text = ""
   flxRole.RowHeight(0) = 0
   flxRole.Cols = 3
   flxRole.ColWidth(0) = 5
   flxRole.ColWidth(1) = 1500
   flxRole.ColWidth(2) = 5000
   flxRole.Clear
   flxRole.ColAlignment(0) = vbLeftJustify
   flxRole.ColAlignment(1) = vbLeftJustify
   flxRole.ColAlignment(2) = vbLeftJustify
   
    
  
   sqlCom = "SELECT * FROM Roles ORDER BY RoleID;"
   rst.Open sqlCom, adoConn, adOpenStatic, adLockReadOnly
rRow = 1
Dim rCount As Integer
rCount = rst.RecordCount
   flxRole.Rows = rst.RecordCount
     If Not rst.EOF Then
      While Not rst.EOF
         If rCount = 1 Then
      flxRole.row = 0
      rRow = 0
      Else
      flxRole.row = 1
      End If
           flxRole.RowSel = 0
           flxRole.ColSel = 0
           flxRole.TextMatrix(rRow, 0) = ""
           flxRole.TextMatrix(rRow, 1) = rst!roleID
           flxRole.TextMatrix(rRow, 2) = rst!RoleName
           flxRole.RowHeight(rRow) = 240
           If Not rst.EOF Then flxRole.AddItem ""
           rRow = rRow + 1
         rst.MoveNext
      Wend
   End If
'flxRole.Rows = flxRole.Rows - 1
' Remove extra rows
'Code added by Mahboob 13/07/2023
    Do While flxRole.Rows > rRow
        flxRole.RemoveItem flxRole.Rows - 1
    Loop
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub cmdNew_Click()
'''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 9 disable enable button
NEWMODE_ = True
'   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, FrameUserInfo, NewEntryMode
   txtUserID.Enabled = False
   chkActive.Value = True
   txtUserName.Enabled = True
   txtPassword.Enabled = True
   txtConPass.Enabled = True
   cmdRole.Enabled = True
   txtEmail.Enabled = True
   'Code Change by Mahboob 17/07/2023
   txtUserName.SetFocus
   'Code change by mahboob 27/08/2023
   gridUserLookup.Enabled = False
End Sub

Private Sub cmdPicCLose_Click()
picRole.Visible = False
End Sub

Private Sub cmdRole_Click()
'''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 22 load the picture box
    picRole.Left = 700
    picRole.Top = 800
    picRole.Visible = True
    LoadRoleList
    picRole.Enabled = True
    txtSearchRoletID.SetFocus
End Sub

Private Sub cmdSave_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 18 save user to db
'Check empty role ,username
If txtUserName.text = "" Then
    MsgBox "Please enter a User Name."
    txtUserName.SetFocus
    Exit Sub
End If
If txtRole.text = "" Then
    MsgBox "Please enter a Role Name."
    cmdRole.SetFocus
    Exit Sub
End If
If txtEmail.text = "" Then
    MsgBox "Please enter a valid email."
    txtEmail.SetFocus
    Exit Sub
End If
If txtPassword.text = "" Then
    MsgBox "Please enter a password."
    txtPassword.SetFocus
    Exit Sub
End If
'Check empty password  confirm password
'If NEWMODE_ Then
If txtPassword.text <> txtConPass.text Then
    MsgBox " The password entered and password confirmed do not match. Please try again."
    'code change by mahboob 27/08/2023
'    txtConPass.SetFocus
txtPassword.SetFocus
    Exit Sub
'End If
If txtPassword.text = "sysadmin#1" Then
MsgBox "The password entered is already in use."
    txtPassword.SetFocus
    Exit Sub
    End If
End If
'check duplicate role name
 If NEWMODE_ Then
    If checkDuplicateUserName Then
        MsgBox "The user name entered already exists. Please enter a different user name."
        'Code change by Mahboob 27/08/2023 dim the next line
'        txtUserName.SetFocus
        Exit Sub
    End If
 Else
    If txtUserName.text <> txtUserName.Tag Then
       If checkDuplicateUserName Then
           MsgBox "The user name entered already exists. Please enter a different user name."
           'Code change by Mahboob 27/08/2023 dim the next line
'           txtUserName.SetFocus
           Exit Sub
       End If
    End If
 End If
 'Save Role to DB
 If SaveUserInformation Then
      MsgBox "User information saved successfully."
      NEWMODE_ = False
      ComponentInFrameEnableMode Me, FrameUserInfo, DefaultMode
'      SEARCHTenantMODE_ = True
'      cmdRoleId.Enabled = True
'      cmdRoleId.Visible = True
'Code change by mahboob 27/08/2023
gridUserLookup.Enabled = True
cmdRole.Enabled = False
   Else
      txtUserID.SetFocus
   End If
   'Load user in grid
   LoadUserList
End Sub

Public Function checkDuplicateUserName() As Boolean
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 19
    Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    psql = "sELECT UserName FROM UserNames WHERE UserName='" & txtUserName.text & "'"
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        checkDuplicateUserName = True
    Else
        checkDuplicateUserName = False
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function

Function SaveUserInformation() As Boolean
''''''Modified by Mahboob 04/04/2023 Change ID 9 work item 20 save users to db
Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess

   Dim sSQLQuery As String
   Dim sSQLFilter As String

   If Not NEWMODE_ Then
    sSQLFilter = "WHERE UserID = " & txtUserID.text & ""
   Else
      sSQLFilter = ""
   End If
 sSQLQuery_ = "SELECT * FROM UserNames " & sSQLFilter

   rst.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
   If NEWMODE_ Then
   rst.AddNew
'    rst!UserPassword = txtPassword.text
   End If
   rst!UserName = txtUserName.text
   rst!UserPassword = txtPassword.text
   rst!roleID = txtRole.Tag
   rst!UserEmail = IIf(txtEmail.text = "", "", txtEmail.text)
   rst!IsActive = IIf(chkActive.Value = True, "Y", "N")
   rst.Update
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
   SaveUserInformation = True
End Function

Private Sub cmdUsersSearch_Click()
''''''''''''Modified by Mahboob 06/04/2023 Change ID 11 work item 1 load the picture box
'picComUserSearch.Left = 190
'    picComUserSearch.Top = 300
    picComUserSearch.Visible = True
    'Load the user in grid
    LoadSearchUserList
    
    picComUserSearch.Enabled = True
    txtComUserId.SetFocus
End Sub
Function LoadAviableCompany()
'''''''''''Modified by Mahboob 06/04/2023 Change ID 10 work item 9 load company from control database
Dim conConn As New ADODB.Connection, rstSQL As New ADODB.Recordset
Dim SQLStr As String
   SQLStr = "SELECT * FROM Databases ORDER BY ID;"
   If InStr(OS, "Server 2008") = 0 Then
      conConn.Open "DSN=PrestigeBMControlNS;UID=;PWD="
   Else
      conConn.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                        "Dbq=PBMControl.mdb;" & _
                        "DefaultDir=" & DB_PATH & ";" & _
                        "Uid=;Pwd=;"
   End If

   rstSQL.Open SQLStr, conConn, adOpenStatic, adLockReadOnly
lvwavailableCompany.ListItems.Clear

   If Not rstSQL.EOF Then
'   If Not .EOF Then
                Do While Not rstSQL.EOF
                        With lvwavailableCompany.ListItems.Add
                            .text = rstSQL!ID
                            .Tag = rstSQL!ID
                            .SubItems(1) = rstSQL!SCName
'                             .SubItems(2) = "" & pRS("Price")
'                            i = i + 1
                        End With
                        rstSQL.MoveNext
                Loop
'            End If
'      While Not rstSQL.EOF
'
''         cboShopCentre.AddItem rstSQL!Id & " / " & rstSQL!SCName
'''         If rstSQL!Id > latest Then latest = rstSQL!Id
''         rstSQL.MoveNext
'      Wend
'While Not rstSQL.EOF
'    Dim newItem As ListItem
'    Set newItem = lvwavailableCompany.ListItems.Add(, , rstSQL!Id)
'    newItem.SubItems(1) = rstSQL!SCName
'    rstSQL.MoveNext
'Wend
   End If
   rstSQL.Close
   Set rstSQL = Nothing
   conConn.Close
   Set conConn = Nothing
   
End Function
Function fnCheckedCompany()
''''''''Modified by Mahboob 04/04/2023 Change ID 11 work item 10 checked the list view which company have permission
Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   Dim psql As String
   adoConn.Open getConnectionStringUserAccess
   psql = ""
   psql = "Select * from  UserCompanyAccess where UserID=" & txtUserAccessName.Tag & ""
   rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
    Do While Not rst.EOF
            For i = 1 To lvwavailableCompany.ListItems.Count
                If lvwavailableCompany.ListItems(i).Tag = Val("" & rst.Fields("CompanyID")) Then lvwavailableCompany.ListItems(i).Checked = True
            Next
            rst.MoveNext
        Loop
    End If
End Function
Function LoadSearchUserList()
'''''''''''''''''''Modified by Mahboob 04/04/2023 Change ID 11 work item 2 function to load user in grid
Dim rRow As Integer
     Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess
    Dim sqlCom As String
   txtComUserId.text = ""
   txtComUserName.text = ""
   txtComRole.text = ""
   txtComEmail.text = ""
   txtComStatus.text = ""
   flxComUserDetails.RowHeight(0) = 0
   flxComUserDetails.Cols = 7
   flxComUserDetails.ColWidth(0) = 5
   flxComUserDetails.ColWidth(1) = 1020
   flxComUserDetails.ColWidth(2) = 1565
   flxComUserDetails.ColWidth(3) = 2195
   flxComUserDetails.ColWidth(4) = 1640
   flxComUserDetails.ColWidth(5) = 1925
   flxComUserDetails.ColWidth(6) = 5
   flxComUserDetails.Clear
   flxComUserDetails.ColAlignment(0) = vbLeftJustify
   flxComUserDetails.ColAlignment(1) = vbLeftJustify
   flxComUserDetails.ColAlignment(2) = vbLeftJustify
   flxComUserDetails.ColAlignment(3) = vbLeftJustify
   flxComUserDetails.ColAlignment(4) = vbLeftJustify
   flxComUserDetails.ColAlignment(5) = vbLeftJustify
   
    
    sqlCom = ""
    sqlCom = sqlCom & " SELECT UserNames.RoleID,UserNames.UserID, UserNames.UserName, Roles.RoleName, UserNames.UserEmail, UserNames.IsActive "
    sqlCom = sqlCom & " FROM Roles INNER JOIN UserNames ON Roles.RoleID = UserNames.RoleID ;"
   rst.Open sqlCom, adoConn, adOpenStatic, adLockReadOnly
rRow = 1
Dim rCount As Integer
rCount = rst.RecordCount
   flxComUserDetails.Rows = rst.RecordCount
     If Not rst.EOF Then
      While Not rst.EOF
         If rCount = 1 Then
      flxComUserDetails.row = 0
      rRow = 0
      Else
      flxComUserDetails.row = 1
      End If
           flxComUserDetails.RowSel = 0
           flxComUserDetails.ColSel = 0
           flxComUserDetails.TextMatrix(rRow, 0) = ""
           flxComUserDetails.TextMatrix(rRow, 1) = rst!UserId
           flxComUserDetails.TextMatrix(rRow, 2) = rst!UserName
           flxComUserDetails.TextMatrix(rRow, 3) = rst!RoleName
           flxComUserDetails.TextMatrix(rRow, 4) = IIf(IsNull(rst!UserEmail), "", rst!UserEmail)
'           flxComUserDetails.TextMatrix(rRow, 5) = rst!IsActive
flxComUserDetails.TextMatrix(rRow, 5) = IIf(rst!IsActive = "Y", "Active", "Disabled")
           flxComUserDetails.TextMatrix(rRow, 6) = rst!roleID
           flxComUserDetails.RowHeight(rRow) = 240
           If Not rst.EOF Then flxComUserDetails.AddItem ""
           rRow = rRow + 1
         rst.MoveNext
      Wend
   End If
'gridUserLookup.Rows = gridUserLookup.Rows - 1
'Code added by mahboob 13/07/2023
Do While flxComUserDetails.Rows > rRow
        flxComUserDetails.RemoveItem flxComUserDetails.Rows - 1
    Loop
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub flxComUserDetails_Click()
''''''''''''''''Modified by Mahboob 06/04/2023 Change ID 11 work item 8
'txtUserAccessName.Tag = flxComUserDetails.TextMatrix(flxComUserDetails.row, 1)
'    txtUserAccessName.text = flxComUserDetails.TextMatrix(flxComUserDetails.row, 2)
'    If txtUserAccessName.Tag = "1" Then
'    MsgBox "The Admin user is a system user and therefore does not require company access"
'    txtUserAccessName.Tag = ""
'    txtUserAccessName.text = ""
'    Exit Sub
'    End If
'    picComUserSearch.Visible = False
'    'Load the company in list view
'    LoadAviableCompany
'    'Check the saved company
'    fnCheckedCompany
'    cmdCompanyAccess.Enabled = True
'''''''''''''''Modified by Mahboob 17/07/2023
'If flxComUserDetails.TextMatrix(flxComUserDetails.row, 2) = "" Then Exit Sub
'If txtUserAccessName.Tag = "1" Then
If flxComUserDetails.TextMatrix(flxComUserDetails.row, 1) = "1" Then
    MsgBox "The Admin user is a system user and therefore does not require company access"
'    txtUserAccessName.Tag = ""
'    txtUserAccessName.text = ""
    Exit Sub
    End If
txtUserAccessName.Tag = flxComUserDetails.TextMatrix(flxComUserDetails.row, 1)
    txtUserAccessName.text = flxComUserDetails.TextMatrix(flxComUserDetails.row, 2)
    
    picComUserSearch.Visible = False
    'Load the company in list view
    LoadAviableCompany
    'Check the saved company
    fnCheckedCompany
    cmdCompanyAccess.Enabled = True
End Sub

Private Sub flxRole_Click()
'''''''''''''''''Modified by Mahboob 04/04/2023 Change ID 9 work item 14  lod role from grid to text box
'    If txtRole.text = flxRole.TextMatrix(flxRole.row, 2) = "" Then Exit Sub
    If flxRole.TextMatrix(flxRole.row, 1) = "1" Then
    MsgBox "The Administrator role cannot be assigned."
'     txtRole.Tag = ""
'     txtRole.text = ""
    Exit Sub
    End If
    txtRole.Tag = flxRole.TextMatrix(flxRole.row, 1)
    txtRole.text = flxRole.TextMatrix(flxRole.row, 2)
     
    If RoleAssain Then
    MsgBox "This role has not been assigned any permissions. Please assign permissions to this role."
    txtRole.Tag = ""
    txtRole.text = ""
    Exit Sub
    End If
    picRole.Visible = False
End Sub
Function RoleAssain() As Boolean
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 7 check the role have permission saved in database
'If txtRolePermissionName.Tag = "" Then Exit Function
Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    RoleAssain = False
    psql = ""
    psql = "SELECT * FROM RolePermissions where RoleID=" & txtRole.Tag & ""
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        'fnLogInCheck = True
    Else
       ' MsgBox "This role is not assian"
        RoleAssain = True
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub Form_Load()
Me.BackColor = MODULEBACKCOLOR
FrameUserInfo.BackColor = MODULEBACKCOLOR
FrameRolePermission.BackColor = MODULEBACKCOLOR
'''''''''''''''''''Modified by Mahboob 04/04/2023 Change ID 9 work item 1 disable enable button and function to load user in grid
SSTabUsers.Tab = 0
ComponentInFrameEnableMode Me, FrameUserInfo, DefaultMode
'   cmdRoleId.Enabled = True
   NEWMODE_ = False
'   SEARCHTenantMODE_ = True
'Load the user in grid view
    LoadUserList
    
End Sub
Function LoadUserList()
'''''''''''''''''''Modified by Mahboob 04/04/2023 Change ID 9 work item 2 function to load user in grid
Dim rRow As Integer
     Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess
    Dim sqlCom As String
   txtSearchUser.text = ""
   txtSearchName.text = ""
   txtSearchRole.text = ""
   txtSearchEmail.text = ""
   txtSearchActive.text = ""
   gridUserLookup.RowHeight(0) = 0
   gridUserLookup.Cols = 7
   gridUserLookup.ColWidth(0) = 5
   gridUserLookup.ColWidth(1) = 1020
   gridUserLookup.ColWidth(2) = 1565
   gridUserLookup.ColWidth(3) = 2195
   gridUserLookup.ColWidth(4) = 1640
   gridUserLookup.ColWidth(5) = 1925
   gridUserLookup.ColWidth(6) = 5
   gridUserLookup.Clear
   gridUserLookup.ColAlignment(0) = vbLeftJustify
   gridUserLookup.ColAlignment(1) = vbLeftJustify
   gridUserLookup.ColAlignment(2) = vbLeftJustify
   gridUserLookup.ColAlignment(3) = vbLeftJustify
   gridUserLookup.ColAlignment(4) = vbLeftJustify
   gridUserLookup.ColAlignment(5) = vbLeftJustify
   
    
    sqlCom = ""
    sqlCom = sqlCom & " SELECT UserNames.RoleID,UserNames.UserID, UserNames.UserName, Roles.RoleName, UserNames.UserEmail, UserNames.IsActive "
    sqlCom = sqlCom & " FROM Roles INNER JOIN UserNames ON Roles.RoleID = UserNames.RoleID;"
   rst.Open sqlCom, adoConn, adOpenStatic, adLockReadOnly
rRow = 1
Dim rCount As Integer
rCount = rst.RecordCount
   gridUserLookup.Rows = rst.RecordCount
     If Not rst.EOF Then
      While Not rst.EOF
         If rCount = 1 Then
      gridUserLookup.row = 0
      rRow = 0
      Else
      gridUserLookup.row = 1
      End If
           gridUserLookup.RowSel = 0
           gridUserLookup.ColSel = 0
           gridUserLookup.TextMatrix(rRow, 0) = ""
           gridUserLookup.TextMatrix(rRow, 1) = rst!UserId
           gridUserLookup.TextMatrix(rRow, 2) = rst!UserName
           gridUserLookup.TextMatrix(rRow, 3) = rst!RoleName
           gridUserLookup.TextMatrix(rRow, 4) = IIf(IsNull(rst!UserEmail), "", rst!UserEmail)
           gridUserLookup.TextMatrix(rRow, 5) = IIf(rst!IsActive = "Y", "Active", "Disabled")
           gridUserLookup.TextMatrix(rRow, 6) = rst!roleID
           gridUserLookup.RowHeight(rRow) = 240
           If Not rst.EOF Then gridUserLookup.AddItem ""
           rRow = rRow + 1
         rst.MoveNext
      Wend
   End If
'gridUserLookup.Rows = gridUserLookup.Rows - 1
' Remove extra rows
'Code added by Mahboob 13/07/2023
    Do While gridUserLookup.Rows > rRow
        gridUserLookup.RemoveItem gridUserLookup.Rows - 1
    Loop
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub gridUserLookup_Click()
'''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 8 load user details in text boxes
 txtUserID.text = gridUserLookup.TextMatrix(gridUserLookup.row, 1)
    txtUserName.text = gridUserLookup.TextMatrix(gridUserLookup.row, 2)
    txtUserName.Tag = gridUserLookup.TextMatrix(gridUserLookup.row, 2)
    txtRole.text = gridUserLookup.TextMatrix(gridUserLookup.row, 3)
    txtEmail.text = gridUserLookup.TextMatrix(gridUserLookup.row, 4)
    chkActive.Value = IIf(gridUserLookup.TextMatrix(gridUserLookup.row, 5) = "Active", True, False)
    txtRole.Tag = gridUserLookup.TextMatrix(gridUserLookup.row, 6)
    txtUserName.Enabled = False
    If txtUserID.text = "1" Then
'    txtUserName.Enabled = False
    cmdRole.Enabled = False
'    txtPassword.Enabled = False
'    txtConPass.Enabled = False
    chkActive.Enabled = False
    Else
'    txtUserName.Enabled = True
    cmdRole.Enabled = False
    txtPassword.Enabled = True
    txtConPass.Enabled = True
    chkActive.Enabled = True
    End If
    Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    psql = "sELECT * FROM UserNames WHERE UserID=" & txtUserID.text & ""
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        txtPassword.text = rst!UserPassword
        txtConPass.text = txtPassword.text
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Sub

Private Sub lvwavailableCompany_ItemClick(ByVal Item As MSComctlLib.ListItem)
'Set CurrentItem = Item
End Sub

Private Sub SSTabUsers_Click(PreviousTab As Integer)
''Modified by Mahboob 04/06/2023 Change ID 11 work item 13 clear company selection
If SSTabUsers.Tab = 1 Then
txtUserAccessName.text = ""
txtUserAccessName.Tag = ""
lvwavailableCompany.ListItems.Clear
cmdCompanyAccess.Enabled = False
End If
End Sub

Private Sub txtComEmail_Change()
''''Modified by Mahboob 06/04/2023 Change ID 11 work item 6 search by Email
Dim i As Integer
   If Len(txtComEmail.text) > 0 Then
        txtComUserId.text = ""
        txtComUserName.text = ""
        txtComRole.text = ""
        txtComStatus.text = ""
   End If
   For i = flxComUserDetails.Rows - 1 To 1 Step -1
   flxComUserDetails.RowHeight(i) = 240
        If InStr(1, UCase(flxComUserDetails.TextMatrix(i, 4)), UCase(txtComEmail.text), vbTextCompare) = 0 Then
              flxComUserDetails.RowHeight(i) = 0
        End If
        If flxComUserDetails.RowHeight(i) = 240 Then
              flxComUserDetails.row = i
        End If
   Next i
End Sub

Private Sub txtComRole_Change()
''''''Modified by Mahboob 06/04/2023 Change ID 11 work item 5 search by Role
Dim i As Integer
   If Len(txtComRole.text) > 0 Then
        txtComUserId.text = ""
        txtComUserName.text = ""
        txtComEmail.text = ""
        txtComStatus.text = ""
   End If
   For i = flxComUserDetails.Rows - 1 To 1 Step -1
   flxComUserDetails.RowHeight(i) = 240
        If InStr(1, UCase(flxComUserDetails.TextMatrix(i, 3)), UCase(txtComRole.text), vbTextCompare) = 0 Then
              flxComUserDetails.RowHeight(i) = 0
        End If
        If flxComUserDetails.RowHeight(i) = 240 Then
              flxComUserDetails.row = i
        End If
   Next i
End Sub

Private Sub txtComStatus_Change()
'''Modified by Mahboob 06/04/2023 Change ID 11 work item 7 search by status
Dim i As Integer
   If Len(txtComStatus.text) > 0 Then
        txtComUserId.text = ""
'        txtComUserName.text
        txtComRole.text = ""
        txtComEmail.text = ""
        txtComUserName.text = ""
   End If
   For i = flxComUserDetails.Rows - 1 To 1 Step -1
   flxComUserDetails.RowHeight(i) = 240
        If InStr(1, UCase(flxComUserDetails.TextMatrix(i, 5)), UCase(txtComStatus.text), vbTextCompare) = 0 Then
              flxComUserDetails.RowHeight(i) = 0
        End If
        If flxComUserDetails.RowHeight(i) = 240 Then
              flxComUserDetails.row = i
        End If
   Next i
End Sub

Private Sub txtComUserId_Change()
'''''''''Modified by Mahboob 06/04/2023 Change ID 11 work item 3 search by User Id
Dim i As Integer
   If Len(txtComUserId.text) > 0 Then
        txtComUserName.text = ""
        txtComRole.text = ""
        txtComEmail.text = ""
        txtComStatus.text = ""
   End If
   For i = flxComUserDetails.Rows - 1 To 1 Step -1
   flxComUserDetails.RowHeight(i) = 240
        If InStr(1, UCase(flxComUserDetails.TextMatrix(i, 1)), UCase(txtComUserId.text), vbTextCompare) = 0 Then
              flxComUserDetails.RowHeight(i) = 0
        End If
        If flxComUserDetails.RowHeight(i) = 240 Then
              flxComUserDetails.row = i
        End If
   Next i
End Sub

Private Sub txtComUserName_Change()
''''''Modified by Mahboob 06/04/2023 Change ID 11 work item 4 search by User name
Dim i As Integer
   If Len(txtComUserName.text) > 0 Then
        txtComUserId.text = ""
        txtComRole.text = ""
        txtComEmail.text = ""
        txtComStatus.text = ""
   End If
   For i = flxComUserDetails.Rows - 1 To 1 Step -1
   flxComUserDetails.RowHeight(i) = 240
        If InStr(1, UCase(flxComUserDetails.TextMatrix(i, 2)), UCase(txtComUserName.text), vbTextCompare) = 0 Then
              flxComUserDetails.RowHeight(i) = 0
        End If
        If flxComUserDetails.RowHeight(i) = 240 Then
              flxComUserDetails.row = i
        End If
   Next i
End Sub

Private Sub txtConPass_LostFocus()
If txtPassword.text <> "" Then
If txtPassword.text <> txtConPass.text Then
    MsgBox " The password entered and password confirmed do not match. Please try again"
    'code change by mahboob 27/08/2023 dim the next line
'    txtConPass.SetFocus
    Exit Sub
    End If
    End If
End Sub

Private Sub txtSearchActive_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 7 search by Status
Dim i As Integer
   If Len(txtSearchActive.text) > 0 Then
   txtSearchName.text = ""
        txtSearchUser.text = ""
txtSearchRole.text = ""
        txtSearchEmail.text = ""

   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 5)), UCase(txtSearchActive.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchEmail_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 6 search by Email
Dim i As Integer
   If Len(txtSearchEmail.text) > 0 Then
        txtSearchUser.text = ""
        txtSearchRole.text = ""
        txtSearchName.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 4)), UCase(txtSearchEmail.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
   End Sub

Private Sub txtSearchName_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 4 search by user Name
Dim i As Integer
   If Len(txtSearchName.text) > 0 Then
        txtSearchUser.text = ""
        txtSearchRole.text = ""
        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 2)), UCase(txtSearchName.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchRole_Change()
'''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 5 search by Role Name
Dim i As Integer
   If Len(txtSearchRole.text) > 0 Then
   txtSearchName.text = ""
        txtSearchUser.text = ""

        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 3)), UCase(txtSearchRole.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

Private Sub txtSearchRoleName_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 16 search by role name
Dim i As Integer
   If Len(txtSearchRoleName.text) > 0 Then
        txtSearchRoletID.text = ""
   End If
   For i = flxRole.Rows - 1 To 1 Step -1
   flxRole.RowHeight(i) = 240
        If InStr(1, UCase(flxRole.TextMatrix(i, 2)), UCase(txtSearchRoleName.text), vbTextCompare) = 0 Then
              flxRole.RowHeight(i) = 0
        End If
        If flxRole.RowHeight(i) = 240 Then
              flxRole.row = i
        End If
   Next i
End Sub

Private Sub txtSearchRoletID_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 15 search by role id
Dim i As Integer
   If Len(txtSearchRoletID.text) > 0 Then
        txtSearchRoleName.text = ""
   End If
   For i = flxRole.Rows - 1 To 1 Step -1
   flxRole.RowHeight(i) = 240
        If InStr(1, UCase(flxRole.TextMatrix(i, 1)), UCase(txtSearchRoletID.text), vbTextCompare) = 0 Then
              flxRole.RowHeight(i) = 0
        End If
        If flxRole.RowHeight(i) = 240 Then
              flxRole.row = i
        End If
   Next i
End Sub

Private Sub txtSearchRoletID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 17 Move focus from id to name
If KeyCode = vbKeyDown Then
        flxRole.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchRoleName.SetFocus
    End If
End Sub

Private Sub txtSearchUser_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 9 work item 3 search by user ID
Dim i As Integer
   If Len(txtSearchUser.text) > 0 Then
        txtSearchName.text = ""
        txtSearchRole.text = ""
        txtSearchEmail.text = ""
        txtSearchActive.text = ""
   End If
   For i = gridUserLookup.Rows - 1 To 1 Step -1
   gridUserLookup.RowHeight(i) = 240
        If InStr(1, UCase(gridUserLookup.TextMatrix(i, 1)), UCase(txtSearchUser.text), vbTextCompare) = 0 Then
              gridUserLookup.RowHeight(i) = 0
        End If
        If gridUserLookup.RowHeight(i) = 240 Then
              gridUserLookup.row = i
        End If
   Next i
End Sub

