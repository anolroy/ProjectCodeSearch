VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoles 
   BackColor       =   &H00FFFFDF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roles and Permissions"
   ClientHeight    =   7485
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRoles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabRole 
      Height          =   7572
      Left            =   -36
      TabIndex        =   0
      Top             =   36
      Width           =   8256
      _ExtentX        =   14552
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Roles"
      TabPicture(0)   =   "frmRoles.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameRoleInfo"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Role Permissions"
      TabPicture(1)   =   "frmRoles.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameRolePermission"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdRolePermissionSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "picRolePermission"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picRolePermission 
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
         Height          =   4452
         Left            =   756
         ScaleHeight     =   4425
         ScaleWidth      =   6975
         TabIndex        =   24
         Top             =   1476
         Visible         =   0   'False
         Width           =   7008
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
            Left            =   6696
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   36
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRoleSearch 
            Height          =   3732
            Left            =   48
            TabIndex        =   25
            Top             =   672
            Width           =   6888
            _ExtentX        =   12144
            _ExtentY        =   6588
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
         Begin MSForms.TextBox txtTvwRoleName 
            Height          =   252
            Left            =   1620
            TabIndex        =   31
            Top             =   372
            Width           =   5340
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "9419;444"
            SpecialEffect   =   0
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtTvwRoleId 
            Height          =   255
            Left            =   45
            TabIndex        =   30
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
         Begin MSForms.Label Label2 
            Height          =   195
            Left            =   120
            TabIndex        =   29
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
            Left            =   1515
            TabIndex        =   28
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblFlxPayee 
            Caption         =   "EMPTY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2115
            TabIndex        =   27
            Top             =   1200
            Width           =   1095
         End
         Begin MSForms.Label Label1 
            Height          =   195
            Left            =   1665
            TabIndex        =   26
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
            Index           =   0
            Left            =   48
            Top             =   900
            Width           =   5856
         End
      End
      Begin VB.CommandButton cmdRolePermissionSearch 
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
         Height          =   336
         Left            =   5076
         TabIndex        =   7
         Top             =   540
         Width           =   384
      End
      Begin VB.Frame FrameRolePermission 
         BackColor       =   &H00FFFFDF&
         Caption         =   "Role Permission Details"
         Height          =   7176
         Left            =   -36
         TabIndex        =   8
         Top             =   252
         Width           =   8256
         Begin VB.OptionButton optCollapse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Collapse All"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6930
            TabIndex        =   35
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optExpand 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Expand All"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   5625
            TabIndex        =   34
            Top             =   315
            Width           =   1140
         End
         Begin VB.CommandButton cmdRolePermission 
            Caption         =   "Save Permissions"
            Height          =   456
            Left            =   6408
            TabIndex        =   11
            Top             =   6480
            Width           =   1644
         End
         Begin MSComctlLib.TreeView tvwRolePremission 
            Height          =   5628
            Left            =   72
            TabIndex        =   32
            Top             =   756
            Width           =   7992
            _ExtentX        =   14102
            _ExtentY        =   9922
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            ImageList       =   "imgFormIcon"
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
         Begin MSComctlLib.ImageList imgFormIcon 
            Left            =   180
            Top             =   144
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   14
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":0902
                  Key             =   ""
                  Object.Tag             =   "Client"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":11DC
                  Key             =   ""
                  Object.Tag             =   "Property"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":1AB6
                  Key             =   ""
                  Object.Tag             =   "Unit"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":2390
                  Key             =   ""
                  Object.Tag             =   "Lessee"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":31E2
                  Key             =   ""
                  Object.Tag             =   "Tenant"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":34FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":3816
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":3C68
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":4542
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":47D5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":4D83
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":51B3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":57D1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRoles.frx":5BD7
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSForms.TextBox txtRolePermissionName 
            Height          =   384
            Left            =   1764
            TabIndex        =   10
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
         Begin VB.Label lblRoleName 
            BackStyle       =   0  'Transparent
            Caption         =   "Role Name"
            Height          =   240
            Left            =   792
            TabIndex        =   9
            Top             =   396
            Width           =   828
         End
      End
      Begin VB.Frame FrameRoleInfo 
         BackColor       =   &H00FFFFDF&
         Caption         =   "Role Details"
         Height          =   7464
         Left            =   -74964
         TabIndex        =   3
         Top             =   252
         Width           =   8580
         Begin VB.PictureBox picRole 
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
            Height          =   5496
            Left            =   0
            ScaleHeight     =   5460
            ScaleWidth      =   8055
            TabIndex        =   16
            Top             =   1764
            Width           =   8088
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRole 
               Height          =   4776
               Left            =   12
               TabIndex        =   17
               Top             =   672
               Width           =   8004
               _ExtentX        =   14129
               _ExtentY        =   8414
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
            Begin VB.Shape Shape4 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00E0FFFF&
               FillStyle       =   0  'Solid
               Height          =   240
               Index           =   15
               Left            =   48
               Top             =   900
               Width           =   5856
            End
            Begin MSForms.Label Label3 
               Height          =   195
               Left            =   1665
               TabIndex        =   23
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
            Begin VB.Label lblFlxPayee 
               Caption         =   "EMPTY"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   2115
               TabIndex        =   22
               Top             =   1200
               Width           =   1095
            End
            Begin VB.Label lblPayeeFlxConfigured 
               Caption         =   "NOT"
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
               Index           =   4
               Left            =   1515
               TabIndex        =   21
               Top             =   1800
               Width           =   1095
            End
            Begin MSForms.Label lblRoleID 
               Height          =   195
               Left            =   120
               TabIndex        =   20
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
            Begin MSForms.TextBox txtSearchRoletID 
               Height          =   252
               Left            =   12
               TabIndex        =   19
               Top             =   372
               Width           =   1536
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "2699;450"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox txtSearchRoleName 
               Height          =   252
               Left            =   1584
               TabIndex        =   18
               Top             =   372
               Width           =   6420
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "11324;444"
               SpecialEffect   =   0
               FontName        =   "Myriad Web"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   345
            Left            =   6480
            TabIndex        =   15
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4932
            TabIndex        =   14
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   345
            Left            =   3492
            TabIndex        =   2
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2016
            TabIndex        =   13
            Top             =   1260
            Width           =   1215
         End
         Begin VB.CommandButton cmdNew 
            Caption         =   "New"
            Height          =   345
            Left            =   540
            TabIndex        =   12
            Top             =   1260
            Width           =   1215
         End
         Begin MSForms.TextBox txtRoleName 
            Height          =   312
            Left            =   2664
            TabIndex        =   1
            Top             =   792
            Width           =   2436
            VariousPropertyBits=   746604571
            Size            =   "4286;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtRoleId 
            Height          =   312
            Left            =   2664
            TabIndex        =   6
            Top             =   216
            Width           =   2436
            VariousPropertyBits=   746604569
            Size            =   "4286;556"
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label LabelRoleName 
            BackStyle       =   0  'Transparent
            Caption         =   "Role Name"
            Height          =   240
            Left            =   1836
            TabIndex        =   5
            Top             =   828
            Width           =   828
         End
         Begin VB.Label LabelROleID 
            BackStyle       =   0  'Transparent
            Caption         =   "Role ID"
            Height          =   240
            Left            =   1836
            TabIndex        =   4
            Top             =   288
            Width           =   828
         End
      End
   End
End
Attribute VB_Name = "frmRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 14
Dim NEWMODE_ As Boolean


'Private Sub chkExpandAll_Click()
''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 17 work item 1 expand de expand tree node
''Expand the nodes
'Dim n As Node
'    For Each n In tvwRolePremission.Nodes
'        If n.Expanded = False Then
'        n.Expanded = True
'        chkExpandAll.Caption = "Collapse All"
'        Else
'        n.Expanded = False
'        chkExpandAll.Caption = "Expand All"
'        End If
'    Next
''    chkExpandAll.Value = 0
'End Sub

Private Sub cmdCancel_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 4 enable and disable button logical ways
ComponentInFrameEnableMode Me, FrameRoleInfo, DefaultMode
'   If NEWMODE_ Then SageCustomerAccCombo cboSageAccountNumber

   NEWMODE_ = False
'   SEARCHTenantMODE_ = True

   txtRoleName.Enabled = True
'''   cmdRoleId.Enabled = True
'   cmdRoleId.Visible = True
'Code change by mahboob 27/08/2023
   flxRole.Enabled = True
   If txtRoleId.text = "" Then Exit Sub


   ComponentInFrameEnableMode Me, FrameRoleInfo, DefaultMode

'   ComponentInFrameClearMode Me, fmeBankPaymentDetails, ClearBoth

   ComponentInFrameClearMode Me, FrameRoleInfo, ClearOnlyTextBoxes
   txtRoleId.Locked = True
'   cmdRoleId.Visible = True
End Sub

Private Sub cmdClose_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 5
    Unload Me
End Sub

Private Sub cmdEdit_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 3 enable and disable button logical ways
 If txtRoleId.text = "" Then
      MsgBox "Please select a Role to continue.", vbInformation, "Edit Role"
      Exit Sub
   End If
   If txtRoleId = 1 Then
    MsgBox "The Administrator role cannot be edited"
    Exit Sub
End If
   NEWMODE_ = False
'   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, FrameRoleInfo, EditMode
 'Code change by mahboob 27/08/2023
   flxRole.Enabled = False
   txtRoleName.SetFocus
'   cmdRoleId.Enabled = False
End Sub

Private Sub cmdNew_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 2 enable and disable button logical ways
   NEWMODE_ = True
'   SEARCHTenantMODE_ = False
   ComponentInFrameEnableMode Me, FrameRoleInfo, NewEntryMode
   txtRoleId.Enabled = False
   'Code Change by Mahboob 17/07/2023
   txtRoleName.SetFocus
   'Code change by mahboob 27/08/2023
   flxRole.Enabled = False
End Sub

'Private Sub cmdPicCLose_Click()
'picRole.Visible = False
'End Sub

'Private Sub cmdRoleId_Click()
'    picRole.Left = 700
'    picRole.Top = 800
'    picRole.Visible = True
'    LoadRoleList
'    picRole.Enabled = True
'    txtSearchRoletID.SetFocus
'End Sub
Function LoadRoleList()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 7 load role in grid
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
'flxRole.Rows = flxRole.Rows - 1
'Code change mahboob 13/07/2023
 Do While flxRole.Rows > rRow
        flxRole.RemoveItem flxRole.Rows - 1
    Loop
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function
Function LoadSearchRoleList()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 2 load role in grid
Dim rRow As Integer
     Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess
    Dim sqlCom As String
   txtTvwRoleId.text = ""
   txtTvwRoleName.text = ""
   flxRoleSearch.RowHeight(0) = 0
   flxRoleSearch.Cols = 3
   flxRoleSearch.ColWidth(0) = 5
   flxRoleSearch.ColWidth(1) = 1500
   flxRoleSearch.ColWidth(2) = 5000
   flxRoleSearch.Clear
   flxRoleSearch.ColAlignment(0) = vbLeftJustify
   flxRoleSearch.ColAlignment(1) = vbLeftJustify
   flxRoleSearch.ColAlignment(2) = vbLeftJustify
   
    
  
   sqlCom = "SELECT * FROM Roles ORDER BY RoleID;"
   rst.Open sqlCom, adoConn, adOpenStatic, adLockReadOnly
rRow = 1
Dim rCount As Integer
rCount = rst.RecordCount
   flxRoleSearch.Rows = rst.RecordCount
     If Not rst.EOF Then
      While Not rst.EOF
         If rCount = 1 Then
      flxRoleSearch.row = 0
      rRow = 0
      Else
      flxRoleSearch.row = 1
      End If
           flxRoleSearch.RowSel = 0
           flxRoleSearch.ColSel = 0
           flxRoleSearch.TextMatrix(rRow, 0) = ""
           flxRoleSearch.TextMatrix(rRow, 1) = rst!roleID
           flxRoleSearch.TextMatrix(rRow, 2) = rst!RoleName
           flxRoleSearch.RowHeight(rRow) = 240
           If Not rst.EOF Then flxRoleSearch.AddItem ""
           rRow = rRow + 1
         rst.MoveNext
      Wend
   End If
'flxRoleSearch.Rows = flxRoleSearch.Rows - 1
'Code change mahboob 13/07/2023
 Do While flxRoleSearch.Rows > rRow
        flxRoleSearch.RemoveItem flxRoleSearch.Rows - 1
    Loop
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
End Function
Public Function checkDuplicateRole() As Boolean
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 8
    Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    psql = ""
    psql = "sELECT RoleName FROM Roles WHERE RoleName='" & txtRoleName.text & "'"
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        checkDuplicateRole = True
    Else
        checkDuplicateRole = False
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Function fnCheckedNode()
'''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 8 checked the node
Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim SubNode As Node
    Dim itemID As String
    Set cnn = New ADODB.Connection
'    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;"
    cnn.Open getConnectionStringUserAccess
'Unchecked all Item node
    For Each SubNode In tvwRolePremission.Nodes
            If (SubNode.Children = 0) Then
            SubNode.Checked = False
            End If
            Next SubNode
    strSQL = "SELECT * FROM RolePermissions where RoleID=" & txtRolePermissionName.Tag & ""
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
rs.MoveFirst
    Do Until rs.EOF
    
    itemID = rs.Fields("ItemID").Value
    
'    For Each PNode In tvwRolePremission.Nodes
'        If PNode.Children > 0 Then
'            Set currentNode = FindNodeByKey(tvwRolePremission.nodes, "S" & itemID)
'            If Not currentNode Is Nothing Then
'        currentNode.Checked = True
'    End If
            For Each SubNode In tvwRolePremission.Nodes
'            If (SubNode.Children = 0) Then
'            SubNode.Checked = False
'            End If
''itemID = ""
'SubNode.Checked = False
                  If itemID = Right(SubNode.key, Len(SubNode.key) - 1) Then
'                  If SubNode.Checked = False Then
                  SubNode.Checked = True
                  itemID = ""
'                  Exit For
                  End If
'
''Else
''SubNode.Checked = False
''Exit For
'                  End If
'                  If itemID = Right(SubNode.key, Len(SubNode.key) - 1) Then
'                  SubNode.Checked = False
''                  Exit For
'                  End If
'                  If (SubNode.Children <> 0) Then
'                SubNode.Checked = True
'                End If
''strSQL = "INSERT INTO RolePermissions (RoleID,ItemID) VALUES (" & txtRolePermissionName.Tag & ",'" & itemID & "')"
''                cnn.Execute strSQL
'
            Next SubNode
rs.MoveNext
Loop
'    End If
End Function

Private Sub cmdPicCLose_Click()
''''Modified by Mahboob 03/04/2023 Change ID 13 work item 12 unload the picture box
picRolePermission.Visible = False
End Sub

'Private Function FindNodeByKey(nodes As nodes, key As String) As Node
'    Dim currentNode As Node
'    For Each currentNode In nodes
'        If currentNode.key = key Then
'            'The node was found
'            Set FindNodeByKey = currentNode
'            Exit Function
'        Else
'            'The node was not found, search its children recursively
'            If currentNode.Children.Count > 0 Then
'                Set FindNodeByKey = FindNodeByKey(currentNode.Children, key)
'                If Not FindNodeByKey Is Nothing Then Exit Function
'            End If
'        End If
'    Next currentNode
'
'    'The node was not found
'    Set FindNodeByKey = Nothing
'End Function
Private Sub cmdRolePermission_Click()
''''Modified by Mahboob 03/04/2023 Change ID 13 work item 9 saved to database
'Code added by mahboob 13/07/2023
'If txtRolePermissionName.text = "" Then
'MsgBox "Please Select a role for assigned"
'Exit Sub
'End If
'Code added by mahboob 17/07/2023
If txtRolePermissionName.text = "" Then
MsgBox "Please select a role to assign"
Exit Sub
End If
Dim cnn As ADODB.Connection
'    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim isChecked As Integer
    Set cnn = New ADODB.Connection
'    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;"
    cnn.Open getConnectionStringUserAccess
    'Check on node selected
     isChecked = 0
     Dim cNode As Node
     For Each cNode In tvwRolePremission.Nodes
            If (cNode.Children = 0) And cNode.Checked = True Then
            isChecked = 1
            End If
            Next cNode
            If isChecked = 0 Then
            MsgBox "No items have been selected. Please select an item"
            Exit Sub
            End If
    strSQL = "DELETE FROM RolePermissions where RoleID=" & txtRolePermissionName.Tag & ""
    cnn.Execute strSQL
    
'    strSQL = "INSERT INTO CheckedItems (ItemID) VALUES (?)"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    
'    Dim PNode As Node
    Dim SubNode As Node
'    For Each PNode In tvwRolePremission.Nodes
'        If PNode.Children > 0 Then
            
            For Each SubNode In tvwRolePremission.Nodes
                If SubNode.Children = 0 And SubNode.Checked Then
                    Dim itemID As String
                    itemID = Right(SubNode.key, Len(SubNode.key) - 1)
strSQL = "INSERT INTO RolePermissions (RoleID,ItemID) VALUES (" & txtRolePermissionName.Tag & ",'" & itemID & "')"
                cnn.Execute strSQL
                End If
            Next
'        End If
'    Next
    
'    rs.Close
    cnn.Close
    
'    Set rs = Nothing
    Set cnn = Nothing
MsgBox "Role Permissions successfully Saved"
End Sub

Private Sub cmdRolePermissionSearch_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 1 load the picture box
'picRole.Left = 700
'    picRole.Top = 800
    picRolePermission.Visible = True
    'Loading the Role in grid
    LoadSearchRoleList
    picRolePermission.Enabled = True
    txtTvwRoleId.SetFocus
    'Code added by mahboob 13/07/2023
    optExpand.Value = 0
    optCollapse.Value = 0
End Sub

Private Sub cmdSave_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 6 save role in db
'Adminstrator Id is 1 which cannot be edited
'If txtRoleId = 1 Then
'    MsgBox "The Administrator role cannot be edited"
'    Exit Sub
'End If
'Check empty role name
If txtRoleName.text = "" Then
    MsgBox "Please enter a Role Name"
    txtRoleName.SetFocus
    Exit Sub
End If
'check duplicate role name
 If NEWMODE_ Then
    If checkDuplicateRole Then
        MsgBox "The role name entered already exists. Please enter a different role name"
        'Code change by mahboob 27/08/2023 dim the next line
'        txtRoleName.SetFocus
        Exit Sub
    End If
 Else
    If txtRoleName.text <> txtRoleName.Tag Then
       If checkDuplicateRole Then
           MsgBox "The role name entered already exists. Please enter a different role name"
           'Code change by mahboob 27/08/2023 dim the next line
'           txtRoleName.SetFocus
           Exit Sub
       End If
    End If
 End If
 'Save Role to DB
 If SaveRoleInformation Then
      MsgBox "Role information successfully saved"
      NEWMODE_ = False
      ComponentInFrameEnableMode Me, FrameRoleInfo, DefaultMode
'      SEARCHTenantMODE_ = True
''      cmdRoleId.Enabled = True
'      cmdRoleId.Visible = True
 'Code change by mahboob 27/08/2023
   flxRole.Enabled = True
   Else
      txtRoleId.SetFocus
   End If
   'Loading the Role in grid after save
   LoadRoleList
End Sub
Public Function SaveRoleInformation() As Boolean
   '''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 9
   Dim adoConn As New ADODB.Connection
   Dim rst As New ADODB.Recordset
   adoConn.Open getConnectionStringUserAccess

   Dim sSQLQuery As String
   Dim sSQLFilter As String

   If Not NEWMODE_ Then
    sSQLFilter = "WHERE RoleID = " & txtRoleId.text & ""
   Else
      sSQLFilter = ""
   End If
 sSQLQuery_ = "SELECT * FROM Roles " & sSQLFilter

   rst.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic
   If NEWMODE_ Then
   rst.AddNew
   End If
   rst!RoleName = txtRoleName.text
   rst.Update
   rst.Close
   Set rst = Nothing
   adoConn.Close
   Set adoConn = Nothing
   SaveRoleInformation = True
End Function

Private Sub flxRole_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 10
    txtRoleId.text = flxRole.TextMatrix(flxRole.row, 1)
    txtRoleName.text = flxRole.TextMatrix(flxRole.row, 2)
    txtRoleName.Tag = flxRole.TextMatrix(flxRole.row, 2)
    'picRole.Visible = False
End Sub

Private Sub flxRoleSearch_Click()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 5 load in text box and in tree
'If flxRoleSearch.TextMatrix(flxRoleSearch.row, 2) = "" Then Exit Sub
If flxRoleSearch.TextMatrix(flxRoleSearch.row, 1) = "1" Then
    MsgBox "The Administrator role cannot be assigned role permissions."
'    txtRolePermissionName.Tag = ""
'    txtRolePermissionName.text = ""
    Exit Sub
    End If
txtRolePermissionName.text = flxRoleSearch.TextMatrix(flxRoleSearch.row, 2)
    txtRolePermissionName.Tag = flxRoleSearch.TextMatrix(flxRoleSearch.row, 1)
'    If txtRolePermissionName.Tag = "1" Then
'    MsgBox "The Administrator role cannot be assigned role permissions"
'    txtRolePermissionName.Tag = ""
'    txtRolePermissionName.text = ""
'    Exit Sub
'    End If
    picRolePermission.Visible = False
    'Loading the Tab, Group,Items in Tree view
    LoadTrvRole
    'code change by mahboob 27/08/2023
    cmdRolePermission.Enabled = True
    'Check current Role has been assain or not.
    If RoleAssain Then
    MsgBox "This role has not been assigned any permissions. Please assign permissions to this role."
    Exit Sub
    End If
    'Checked tree node, which item has been save to DB
    fnCheckedNode
End Sub
Function RoleAssain() As Boolean
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 7 check the role have permission saved in database
Dim adoConn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    adoConn.Open getConnectionStringUserAccess
    Dim psql As String
    RoleAssain = False
    psql = ""
    psql = "SELECT * FROM RolePermissions where RoleID=" & txtRolePermissionName.Tag & ""
    rst.Open psql, adoConn, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        'fnLogInCheck = True
    Else
       ' MsgBox "This role is not assigned"
        RoleAssain = True
    End If
    rst.Close
    Set rst = Nothing
    adoConn.Close
    Set adoConn = Nothing
End Function
Private Sub Form_Load()
Me.BackColor = MODULEBACKCOLOR
FrameRoleInfo.BackColor = MODULEBACKCOLOR
FrameRolePermission.BackColor = MODULEBACKCOLOR
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 1 enable and disable button logical ways and load role in grid
   SSTabRole.Tab = 0
   ComponentInFrameEnableMode Me, FrameRoleInfo, DefaultMode
'   cmdRoleId.Enabled = True
   NEWMODE_ = False
'   SEARCHTenantMODE_ = True

'Loading the Role in grid
    LoadRoleList
End Sub

Private Sub optCollapse_Click()
'''''Modified by Mahboob 03/04/2023 Change ID 17 work item  de expand treeview
Dim n As Node
    For Each n In tvwRolePremission.Nodes
        If n.Expanded = True Then
        n.Expanded = False
'        chkExpandAll.Caption = "De Expand All"
'        Else
'        n.Expanded = False
'        chkExpandAll.Caption = "Expand All"
        End If
    Next
'    '    tvwRolePremission.Nodes(1).Selected = True
'If tvwRolePremission.Nodes.Count > 0 Then
'        ' Select the first node
'        tvwRolePremission.Nodes(1).Selected = True
'    End If
'tvwRolePremission.Nodes(1).Selected = False
Set tvwRolePremission.SelectedItem = Nothing
End Sub

Private Sub optExpand_Click()
''''''''''Modified by Mahboob 03/04/2023 Change ID 17 work item 1 expand treeview
'tvwRolePremission.Nodes(1).Selected = Nothing
'tvwRolePremission.se
Dim n As Node
    For Each n In tvwRolePremission.Nodes
        If n.Expanded = False Then
        n.Expanded = True
'        chkExpandAll.Caption = "De Expand All"
'        Else
'        n.Expanded = False
'        chkExpandAll.Caption = "Expand All"
        End If
    Next
'    tvwRolePremission.Nodes(1).Selected = True
'If tvwRolePremission.Nodes.Count > 0 Then
'        ' Select the first node
'        tvwRolePremission.Nodes(1).Selected = True
'    End If
'tvwRolePremission.Nodes(1).Selected = True
'tvwRolePremission.Nodes(1).Selected = False
If tvwRolePremission.Nodes.Count > 0 Then
        tvwRolePremission.Nodes(1).Selected = True
    End If
End Sub

Private Sub SSTabRole_Click(PreviousTab As Integer)
'Code change by mahboob 27/08/2023 clearing role permission tab
If SSTabRole.Tab = 1 Then
txtRolePermissionName.text = ""
txtRolePermissionName.Tag = ""
ClearTreeView tvwRolePremission
cmdRolePermission.Enabled = False
End If
End Sub
Private Sub ClearTreeView(tv As TreeView)
    'Code change by mahboob 27/08/2023 clearing node
    Dim i As Integer
    
    ' Remove nodes in reverse order to avoid index issues
    For i = tv.Nodes.Count To 1 Step -1
        tv.Nodes.Remove i
    Next i
End Sub

Private Sub txtSearchRoleName_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 12
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
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 11
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
Function LoadTrvRole()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 6
    ' Load data from database and populate TreeView control
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set cnn = New ADODB.Connection
'    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;"
    cnn.Open getConnectionStringUserAccess
    'Some item all ready checked, so all items display set to true expect those items
'    cnn.Execute "UPDATE Items SET Display =True where ID not in ('" & "I107" & "','" & "I112" & "','" & "I19" & "','" & "I22" & "','" & "I26" & "','" & "I48" & "','" & "I52" & "','" & "I69" & "','" & "I73" & "','" & "I85" & "','" & "I86" & "','" & "I94" & "','" & "I95" & "','" & "I96" & "','" & "I97" & "','" & "I98" & "')"
     cnn.Execute "UPDATE Items SET Display =True where ID not in (Select ID from TempItemLoad)"
     strSQL = ""
     strSQL = "SELECT Tabs.ID AS TabID, Tabs.TabName, Tabs.Display, Groups.ID AS GID, Groups.GroupName, Groups.Display, Items.ItemNameL1 & ' ' & Items.ItemNameL2 AS ItemName, Items.Display, Items.ID as ItemID"
strSQL = strSQL & " FROM Tabs LEFT JOIN (Groups LEFT JOIN Items ON Groups.ID = Items.G_ID) ON Tabs.ID = Groups.T_ID"
strSQL = strSQL & " WHERE (((Tabs.Display)=True) AND ((Groups.Display)=True) AND ((Items.Display)=True));"
    
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
    tvwRolePremission.Nodes.Clear
    If Not rs.EOF Then
        rs.MoveFirst
        Dim CurrentTabID As String
        Dim CurrentGroupID As String
        Dim RootNode As Node
        Dim CurrentTabNode As Node
        Dim CurrentGroupNode As Node

'      Set RootNode = tvwRolePremission.Nodes.Add(, , "Root", "Tabs")
        Do While Not rs.EOF
            If rs!TabID <> CurrentTabID Then
                Set CurrentTabNode = tvwRolePremission.Nodes.Add(, , "C" & rs!TabID, rs!TabName)
                CurrentTabID = rs!TabID
                CurrentGroupID = 0
                CurrentTabNode.Checked = True
            End If
            If rs!GID <> CurrentGroupID Then
                Set CurrentGroupNode = tvwRolePremission.Nodes.Add("C" & rs!TabID, tvwChild, "S" & rs!GID, rs!GroupName)
                CurrentGroupID = rs!GID
                CurrentGroupNode.Checked = True
            End If
            If Not IsNull(rs!itemID) Then
                ' Add checkbox only to grandchild nodes
                Dim ItemNode As Node

                Set ItemNode = tvwRolePremission.Nodes.Add("S" & rs.Fields("GID"), tvwChild, "S" & rs.Fields("ItemID"), rs.Fields("ItemName"))
'                Set ItemNode = tvwRolePremission.Nodes.Add("S" & rs.Fields("GID"), tvwChild, "S" & rs.Fields("ItemID"), rs.Fields("ItemName"), 1, 2)

ItemNode.Checked = True
'                tvwRolePremission.Nodes.Add "S" & rs.Fields("GID"), tvwChild, "S" & rs.Fields("ItemID"), rs.Fields("ItemName"), 1, 2
' tvwReports.Nodes.Add , , adoRst.Fields.Item(0).Value, adoRst.Fields.Item(1).Value, 3, 2
'                tvwRolePremission.Nodes.Add , , "S" & rs!GID, tvwChild, , rs!ItemName, , "I" & rs!ItemID
'                ItemNode.CheckBoxes = True
            End If
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    cnn.Close
    Set CurrentTabNode = Nothing
    Set CurrentGroupNode = Nothing
    Set ItemNode = Nothing
    Set rs = Nothing
    Set cnn = Nothing
    'Expand the nodes
'Dim n As Node
'    For Each n In tvwRolePremission.Nodes
'        If n.Expanded = False Then n.Expanded = True
'    Next
End Function
Private Sub txtSearchRoletID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 8 work item 13
If KeyCode = vbKeyDown Then
        flxRole.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchRoleName.SetFocus
    End If
End Sub

Private Sub txtTvwRoleId_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 3
If Len(txtTvwRoleId.text) > 0 Then
        txtTvwRoleName.text = ""
   End If
   For i = flxRoleSearch.Rows - 1 To 1 Step -1
   flxRoleSearch.RowHeight(i) = 240
        If InStr(1, UCase(flxRoleSearch.TextMatrix(i, 1)), UCase(txtTvwRoleId.text), vbTextCompare) = 0 Then
              flxRoleSearch.RowHeight(i) = 0
        End If
        If flxRoleSearch.RowHeight(i) = 240 Then
              flxRoleSearch.row = i
        End If
   Next i
End Sub

Private Sub txtTvwRoleName_Change()
'''''''''''''''''''Modified by Mahboob 03/04/2023 Change ID 13 work item 4
If Len(txtTvwRoleName.text) > 0 Then
        txtTvwRoleId.text = ""
   End If
   For i = flxRoleSearch.Rows - 1 To 1 Step -1
   flxRoleSearch.RowHeight(i) = 240
        If InStr(1, UCase(flxRoleSearch.TextMatrix(i, 2)), UCase(txtTvwRoleName.text), vbTextCompare) = 0 Then
              flxRoleSearch.RowHeight(i) = 0
        End If
        If flxRoleSearch.RowHeight(i) = 240 Then
              flxRoleSearch.row = i
        End If
        Next i
End Sub
