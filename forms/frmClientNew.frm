VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   Icon            =   "frmClientNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12060
   Begin TabDlg.SSTab tabMain 
      Height          =   5295
      Left            =   120
      TabIndex        =   160
      Top             =   2600
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   5
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmClientNew.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "cmdClientDetailsSave"
      Tab(0).Control(3)=   "cmdClientDetailsEdit"
      Tab(0).Control(4)=   "cmdClientDetailsCancel"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Property"
      TabPicture(1)   =   "frmClientNew.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOccupied"
      Tab(1).Control(1)=   "fraType"
      Tab(1).Control(2)=   "cmdImgDelete"
      Tab(1).Control(3)=   "cmdUploadImageAdd"
      Tab(1).Control(4)=   "tvwLandLord"
      Tab(1).Control(5)=   "imgList"
      Tab(1).Control(6)=   "lblImageName"
      Tab(1).Control(7)=   "cmdImgLeftMove"
      Tab(1).Control(8)=   "imgPremises"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Agreement"
      TabPicture(2)   =   "frmClientNew.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCommission"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(3)=   "fraAgreement"
      Tab(2).Control(4)=   "Frame7"
      Tab(2).Control(5)=   "cmdAgmntEdit"
      Tab(2).Control(6)=   "cmdAgmntDelete"
      Tab(2).Control(7)=   "cmdAgmntSave"
      Tab(2).Control(8)=   "cboProperty"
      Tab(2).Control(9)=   "dtDate"
      Tab(2).Control(10)=   "Label15"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Bank/Payment Details"
      TabPicture(3)   =   "frmClientNew.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraBank(0)"
      Tab(3).Control(1)=   "fraBank(1)"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Account History"
      TabPicture(4)   =   "frmClientNew.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture2"
      Tab(4).Control(1)=   "MSHFlexGrid1"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Global Settings"
      TabPicture(5)   =   "frmClientNew.frx":0956
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "tabFee"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdGSSave"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdGSEdit"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdGSCancel"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Memo/Attachemnt"
      TabPicture(6)   =   "frmClientNew.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame17"
      Tab(6).Control(1)=   "cmdUnitMemoCancel"
      Tab(6).Control(2)=   "cmdUnitMemoSave"
      Tab(6).Control(3)=   "cmdUnitMemoEdit"
      Tab(6).Control(4)=   "txtNote"
      Tab(6).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Landlord Address:"
         Height          =   4575
         Left            =   -74640
         TabIndex        =   293
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtClientAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtClientPostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtClientAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtClientAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   24
            Top             =   3960
            Width           =   2655
         End
         Begin VB.TextBox txtClientPersonalEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   23
            Top             =   3480
            Width           =   2655
         End
         Begin VB.TextBox txtClientMobile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Top             =   3000
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Top             =   2565
            Width           =   2655
         End
         Begin VB.TextBox txtClientHomeTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   20
            Top             =   2160
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   300
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   299
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Tel:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   298
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Email:"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   297
            Top             =   3480
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile:"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   296
            Top             =   3000
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Email:"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   295
            Top             =   3960
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Tel:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   294
            Top             =   2520
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Landlord Statement Address:"
         Height          =   4095
         Left            =   -68760
         TabIndex        =   290
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtClientOfficeAddressLine1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficeAddressLine3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtClientOfficePostCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtClientOfficeAddressLine2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Index           =   16
            Left            =   360
            TabIndex        =   292
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   15
            Left            =   360
            TabIndex        =   291
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame fraOccupied 
         BackColor       =   &H80000016&
         Caption         =   "Tenancy Details:"
         Height          =   2535
         Left            =   -70320
         TabIndex        =   280
         Top             =   2640
         Width           =   3735
         Begin VB.TextBox txtPreRentRvw 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtPreTenancyType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtPreOccupiedFr 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label lblTenantNameLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantName"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            MousePointer    =   99  'Custom
            TabIndex        =   289
            Top             =   450
            Width           =   1095
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   288
            Top             =   450
            Width           =   1020
         End
         Begin VB.Label lblTenantIDLink 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TenantID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            MouseIcon       =   "frmClientNew.frx":098E
            MousePointer    =   99  'Custom
            TabIndex        =   287
            Top             =   200
            Width           =   810
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenant ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   286
            Top             =   200
            Width           =   765
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   285
            Top             =   1800
            Width           =   1365
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tenancy Type:"
            Height          =   195
            Left            =   120
            TabIndex        =   284
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   283
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   282
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Review Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   281
            Top             =   2160
            Width           =   1365
         End
      End
      Begin VB.Frame fraType 
         BackColor       =   &H80000016&
         Caption         =   "LANDLORD"
         Height          =   2175
         Left            =   -70305
         TabIndex        =   277
         Top             =   360
         Width           =   3720
         Begin VB.TextBox txtTVInfoName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   255
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Index           =   0
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Index           =   2
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   1380
            Width           =   2655
         End
         Begin VB.TextBox txtTVInfoPC 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1770
            Width           =   1455
         End
         Begin VB.TextBox txtTVInfoAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFEA&
            Height          =   285
            Index           =   1
            Left            =   740
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   990
            Width           =   2655
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Left            =   80
            TabIndex        =   279
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   75
            TabIndex        =   278
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame fraCommission 
         Caption         =   "Commission/Fees:"
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   272
         Top             =   720
         Width           =   4935
         Begin VB.Frame Frame5 
            Caption         =   "Letting Fee"
            Height          =   1095
            Left            =   120
            TabIndex        =   275
            Top             =   240
            Width           =   4575
            Begin VB.TextBox txtLETTING_FEES_VALUE 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   240
               Width           =   2000
            End
            Begin VB.TextBox txtLETTING_FEES_VALUE 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   600
               Width           =   2000
            End
            Begin VB.Frame fraLettingFees 
               BorderStyle     =   0  'None
               Caption         =   "Frame15"
               Enabled         =   0   'False
               Height          =   615
               Left            =   120
               TabIndex        =   276
               Top             =   240
               Width           =   2055
               Begin VB.OptionButton optLetting_Fees 
                  Caption         =   "Fixed Amount "
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   45
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.OptionButton optLetting_Fees 
                  Caption         =   "% of Rent Receivable"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   43
                  Top             =   0
                  Width           =   1935
               End
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Management Fee"
            Height          =   975
            Left            =   120
            TabIndex        =   273
            Top             =   1320
            Width           =   4575
            Begin VB.TextBox txtMGT_FEES_VALUE 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   2000
            End
            Begin VB.TextBox txtMGT_FEES_VALUE 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   600
               Width           =   2000
            End
            Begin VB.Frame fraManagementFees 
               BorderStyle     =   0  'None
               Caption         =   "Frame15"
               Enabled         =   0   'False
               Height          =   615
               Left            =   120
               TabIndex        =   274
               Top             =   240
               Width           =   1815
               Begin VB.OptionButton optManagement_Fees 
                  Caption         =   "% of Rent Received"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   47
                  Top             =   0
                  Width           =   2055
               End
               Begin VB.OptionButton optManagement_Fees 
                  Caption         =   "Fixed Fee"
                  Height          =   375
                  Index           =   1
                  Left            =   0
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1095
               End
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Recharges"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   270
         Top             =   3240
         Width           =   4935
         Begin VB.Frame fraRecharges 
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            Enabled         =   0   'False
            Height          =   975
            Left            =   200
            TabIndex        =   271
            Top             =   240
            Width           =   3375
            Begin VB.OptionButton optRecharges 
               Caption         =   "Landlord pays repair invoices"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   53
               Top             =   720
               Width           =   2415
            End
            Begin VB.OptionButton optRecharges 
               Caption         =   "Landlord does own repairs"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   52
               Top             =   360
               Width           =   2295
            End
            Begin VB.OptionButton optRecharges 
               Caption         =   "Agent manages and pays repairs invoices"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   3375
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Landlord Rent Receivable/Payable:"
         Height          =   855
         Left            =   -67800
         TabIndex        =   267
         Top             =   2760
         Width           =   4455
         Begin VB.TextBox txtRRPA 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   320
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rent Receivable (p.a.):"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   269
            Top             =   320
            Width           =   1650
         End
         Begin MSForms.Label Label2 
            Height          =   200
            Left            =   120
            TabIndex        =   268
            Top             =   600
            Width           =   4000
            ForeColor       =   128
            Caption         =   "Rent payable is based on rent received from tenants."
            Size            =   "7056;353"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraAgreement 
         Caption         =   "Dates:"
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   -67800
         TabIndex        =   262
         Top             =   720
         Width           =   4455
         Begin VB.TextBox txtSTART_DATE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   360
            Width           =   2000
         End
         Begin VB.TextBox txtEND_DATE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   720
            Width           =   2000
         End
         Begin VB.TextBox txtREVIEW_DATE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1080
            Width           =   2000
         End
         Begin VB.TextBox txtNOTICE_DAYS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1440
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Notice Period (Days):"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   266
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Next Review Date:"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   265
            Top             =   1080
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agreement End Date:"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   264
            Top             =   720
            Width           =   1530
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agreement Start Date:"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   263
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ownership Percentage"
         Height          =   855
         Left            =   -67800
         TabIndex        =   260
         Top             =   3720
         Width           =   4455
         Begin VB.TextBox txtOWNERSHIP_PERCENT 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   360
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ownership Percentage:"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   261
            Top             =   360
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdAgmntEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -67800
         TabIndex        =   60
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgmntDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64560
         TabIndex        =   62
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgmntSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66120
         TabIndex        =   61
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Bank Details:"
         Height          =   2295
         Index           =   0
         Left            =   -74760
         TabIndex        =   254
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtBANK_NAME 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   600
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   960
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1260
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_ADDRESS3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1560
            Width           =   3195
         End
         Begin VB.TextBox txtBANK_POST_CODE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   1920
            Width           =   1395
         End
         Begin VB.TextBox txtBank_ID_ 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   255
            Top             =   1920
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdNewBank 
            Caption         =   "New"
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
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSAdodcLib.Adodc adoBank 
            Height          =   330
            Left            =   3240
            Top             =   1920
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
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
            Caption         =   "Adodc1"
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   259
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Left            =   120
            TabIndex        =   258
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   257
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   256
            Top             =   240
            Width           =   630
         End
         Begin MSForms.ComboBox cboBank_ID 
            Height          =   285
            Left            =   1200
            TabIndex        =   63
            Top             =   240
            Width           =   3195
            VariousPropertyBits=   1820346399
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5636;503"
            TextColumn      =   1
            ColumnCount     =   6
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame fraBank 
         Caption         =   "Default Account Details:"
         Height          =   2295
         Index           =   1
         Left            =   -67800
         TabIndex        =   248
         Top             =   480
         Width           =   4575
         Begin VB.TextBox txtBank_AC_Name 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   720
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_SC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   72
            Top             =   1080
            Width           =   2800
         End
         Begin VB.TextBox txtBANK_AC_NUM 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   1440
            Width           =   2800
         End
         Begin VB.TextBox txtBacsRef 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   74
            Top             =   1800
            Width           =   2800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   195
            Index           =   57
            Left            =   120
            TabIndex        =   253
            Top             =   720
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sort Code:"
            Height          =   195
            Index           =   58
            Left            =   120
            TabIndex        =   252
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number:"
            Height          =   195
            Index           =   59
            Left            =   120
            TabIndex        =   251
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "BACS REF:"
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   250
            Top             =   1800
            Width           =   825
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method:"
            Height          =   195
            Index           =   56
            Left            =   120
            TabIndex        =   249
            Top             =   240
            Width           =   1335
         End
         Begin MSForms.ComboBox cboPaymentMethod 
            Height          =   285
            Left            =   1560
            TabIndex        =   70
            Top             =   240
            Width           =   2800
            VariousPropertyBits=   746604575
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4939;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -67440
         ScaleHeight     =   1305
         ScaleWidth      =   4185
         TabIndex        =   244
         Top             =   480
         Width           =   4215
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   82
            Top             =   840
            Width           =   2000
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   81
            Top             =   480
            Width           =   2000
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   80
            Top             =   120
            Width           =   2000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Receivable (YTD):"
            Height          =   195
            Index           =   55
            Left            =   120
            TabIndex        =   247
            Top             =   840
            Width           =   1710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Received (YTD):"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   246
            Top             =   480
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFDFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance:"
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   245
            Top             =   120
            Width           =   630
         End
      End
      Begin VB.TextBox txtNote 
         Height          =   3135
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   243
         Top             =   480
         Width           =   11595
      End
      Begin VB.CommandButton cmdUnitMemoEdit 
         Caption         =   "&Edit Memo"
         Height          =   435
         Left            =   -67920
         TabIndex        =   242
         Top             =   3780
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoSave 
         Caption         =   "&Save Memo"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -66300
         TabIndex        =   241
         Top             =   3780
         Width           =   1350
      End
      Begin VB.CommandButton cmdUnitMemoCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -64680
         TabIndex        =   240
         Top             =   3780
         Width           =   1350
      End
      Begin VB.ComboBox cboProperty 
         Height          =   315
         Left            =   -70560
         TabIndex        =   42
         Top             =   400
         Width           =   3495
      End
      Begin VB.Frame Frame14 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   168
         Top             =   2760
         Width           =   11535
         Begin VB.CommandButton cmdAddNewBank 
            Caption         =   "&Add New"
            Height          =   360
            Left            =   3720
            TabIndex        =   75
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   360
            Left            =   6960
            TabIndex        =   77
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteBank 
            Caption         =   "&Delete"
            Height          =   360
            Left            =   10080
            TabIndex        =   79
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditBank 
            Caption         =   "&Edit"
            Height          =   360
            Left            =   5280
            TabIndex        =   76
            Top             =   2020
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "Canc&el"
            Enabled         =   0   'False
            Height          =   360
            Left            =   8520
            TabIndex        =   78
            Top             =   2020
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOtherBankDetails 
            Height          =   1785
            Left            =   120
            TabIndex        =   169
            Top             =   195
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   3149
            _Version        =   393216
            BackColorFixed  =   12632256
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.CommandButton cmdGSCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   10080
         TabIndex        =   143
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   6840
         TabIndex        =   141
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdGSSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   8460
         TabIndex        =   142
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientDetailsSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -66360
         TabIndex        =   30
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientDetailsEdit 
         Caption         =   "&Edit"
         Height          =   360
         Left            =   -68040
         TabIndex        =   29
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClientDetailsCancel 
         Caption         =   "Canc&el"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -64800
         TabIndex        =   31
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame Frame17 
         Caption         =   "Attactment Files:"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -74760
         TabIndex        =   163
         Top             =   4200
         Width           =   11535
         Begin VB.CommandButton cmdOpenFile 
            Caption         =   "&Open File"
            Height          =   435
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdClinetAddAtch 
            Caption         =   "&Add New"
            Height          =   435
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   165
            Top             =   240
            Width           =   1350
         End
         Begin VB.CommandButton cmdDeleteFile 
            Caption         =   "&Delete File"
            Height          =   435
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   240
            Width           =   1350
         End
         Begin MSForms.ComboBox cmbFiles 
            Height          =   285
            Left            =   120
            TabIndex        =   167
            Top             =   360
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
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763;4233"
         End
      End
      Begin VB.CommandButton cmdImgDelete 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -63840
         TabIndex        =   162
         Top             =   3960
         Width           =   555
      End
      Begin VB.CommandButton cmdUploadImageAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -64440
         TabIndex        =   161
         Top             =   3960
         Width           =   555
      End
      Begin TabDlg.SSTab tabFee 
         Height          =   4095
         Left            =   600
         TabIndex        =   170
         Top             =   480
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Letting Fee"
         TabPicture(0)   =   "frmClientNew.frx":0C98
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkLettingFee"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraFee(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Management Fee"
         TabPicture(1)   =   "frmClientNew.frx":0CB4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkMngFee"
         Tab(1).Control(1)=   "fraFee(1)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Rent Payable"
         TabPicture(2)   =   "frmClientNew.frx":0CD0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkRentPayble"
         Tab(2).Control(1)=   "fraFee(2)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Payment Dates"
         TabPicture(3)   =   "frmClientNew.frx":0CEC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label83(13)"
         Tab(3).Control(1)=   "Label83(0)"
         Tab(3).Control(2)=   "tabDates"
         Tab(3).Control(3)=   "txtPayIsuDays"
         Tab(3).Control(4)=   "txtFeeIsuDays"
         Tab(3).ControlCount=   5
         Begin VB.Frame fraFee 
            BackColor       =   &H00F1F1F1&
            Caption         =   "Letting Fee:"
            Enabled         =   0   'False
            Height          =   2655
            Index           =   0
            Left            =   720
            TabIndex        =   183
            Top             =   780
            Width           =   9255
            Begin VB.ComboBox cboLettingAM 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0D08
               Left            =   2160
               List            =   "frmClientNew.frx":0D12
               TabIndex        =   84
               Top             =   480
               Width           =   2655
            End
            Begin VB.ComboBox cboLettingFreq 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0D29
               Left            =   6240
               List            =   "frmClientNew.frx":0D39
               TabIndex        =   87
               Top             =   480
               Width           =   2655
            End
            Begin VB.TextBox txtLettingNtDueDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2160
               TabIndex        =   85
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtLettingStDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6240
               TabIndex        =   88
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Letting Fee:"
               Height          =   195
               Index           =   23
               Left            =   840
               TabIndex        =   188
               Top             =   480
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Frequency:"
               Height          =   195
               Index           =   25
               Left            =   5040
               TabIndex        =   187
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Next Due Date:"
               Height          =   195
               Index           =   24
               Left            =   840
               TabIndex        =   186
               Top             =   1320
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fee Start Date:"
               Height          =   195
               Index           =   26
               Left            =   5040
               TabIndex        =   185
               Top             =   1320
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fee Charge Type:"
               Height          =   195
               Index           =   64
               Left            =   840
               TabIndex        =   184
               Top             =   2040
               Width           =   1275
            End
            Begin MSForms.ComboBox cboLettingChrgType 
               Height          =   315
               Left            =   2160
               TabIndex        =   86
               Top             =   2040
               Width           =   2655
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4683;556"
               BoundColumn     =   2
               ColumnCount     =   2
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1058;5000"
            End
         End
         Begin VB.Frame fraFee 
            BackColor       =   &H00F1F1F1&
            Caption         =   "Management Fee:"
            Enabled         =   0   'False
            Height          =   2655
            Index           =   1
            Left            =   -74280
            TabIndex        =   177
            Top             =   780
            Width           =   9255
            Begin VB.TextBox txtMngStDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6240
               TabIndex        =   94
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtMngNtDueDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2160
               TabIndex        =   91
               Top             =   1320
               Width           =   2655
            End
            Begin VB.ComboBox cboMngFreq 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0D66
               Left            =   6240
               List            =   "frmClientNew.frx":0D76
               TabIndex        =   93
               Top             =   480
               Width           =   2655
            End
            Begin VB.ComboBox cboMngAM 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0DA3
               Left            =   2160
               List            =   "frmClientNew.frx":0DAD
               TabIndex        =   90
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fee Start Date:"
               Height          =   195
               Index           =   30
               Left            =   5040
               TabIndex        =   182
               Top             =   1320
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Next Due Date:"
               Height          =   195
               Index           =   28
               Left            =   840
               TabIndex        =   181
               Top             =   1320
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Frequency:"
               Height          =   195
               Index           =   29
               Left            =   5040
               TabIndex        =   180
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Management Fee:"
               Height          =   195
               Index           =   27
               Left            =   840
               TabIndex        =   179
               Top             =   480
               Width           =   1290
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fee Charge Type:"
               Height          =   195
               Index           =   65
               Left            =   840
               TabIndex        =   178
               Top             =   2040
               Width           =   1275
            End
            Begin MSForms.ComboBox cboMngChrgType 
               Height          =   285
               Left            =   2160
               TabIndex        =   92
               Top             =   2040
               Width           =   2655
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4683;503"
               ColumnCount     =   2
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1058;5000"
            End
         End
         Begin VB.Frame fraFee 
            BackColor       =   &H00F1F1F1&
            Caption         =   "Rent Payable:"
            Enabled         =   0   'False
            Height          =   2655
            Index           =   2
            Left            =   -74280
            TabIndex        =   171
            Top             =   780
            Width           =   9255
            Begin VB.TextBox txtRentStDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6240
               TabIndex        =   100
               Top             =   1320
               Width           =   2655
            End
            Begin VB.TextBox txtRentNtDueDt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2160
               TabIndex        =   97
               Top             =   1320
               Width           =   2655
            End
            Begin VB.ComboBox cboRentFreq 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0DC4
               Left            =   6240
               List            =   "frmClientNew.frx":0DD4
               TabIndex        =   99
               Top             =   480
               Width           =   2655
            End
            Begin VB.ComboBox cboRentAM 
               Height          =   315
               ItemData        =   "frmClientNew.frx":0E01
               Left            =   2160
               List            =   "frmClientNew.frx":0E0B
               TabIndex        =   96
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Start Date:"
               Height          =   195
               Index           =   34
               Left            =   5040
               TabIndex        =   176
               Top             =   1320
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Next Due Date:"
               Height          =   195
               Index           =   33
               Left            =   840
               TabIndex        =   175
               Top             =   1320
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Frequency:"
               Height          =   195
               Index           =   32
               Left            =   5040
               TabIndex        =   174
               Top             =   480
               Width           =   795
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Rent Payable:"
               Height          =   195
               Index           =   31
               Left            =   840
               TabIndex        =   173
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFDFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fee Charge Type:"
               Height          =   195
               Index           =   66
               Left            =   840
               TabIndex        =   172
               Top             =   2040
               Width           =   1275
            End
            Begin MSForms.ComboBox cboRentChrgType 
               Height          =   285
               Left            =   2160
               TabIndex        =   98
               Top             =   2040
               Width           =   2655
               VariousPropertyBits=   746604571
               BorderStyle     =   1
               DisplayStyle    =   3
               Size            =   "4683;503"
               ColumnCount     =   2
               cColumnInfo     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               SpecialEffect   =   0
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               Object.Width           =   "1058;5000"
            End
         End
         Begin VB.TextBox txtFeeIsuDays 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -72960
            TabIndex        =   101
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtPayIsuDays 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -66720
            TabIndex        =   102
            Top             =   480
            Width           =   915
         End
         Begin TabDlg.SSTab tabDates 
            Height          =   3075
            Left            =   -74520
            TabIndex        =   189
            Top             =   840
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   5424
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Monthly Payment Dates"
            TabPicture(0)   =   "frmClientNew.frx":0E22
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fraPaymentDate(2)"
            Tab(0).Control(1)=   "fraPaymentDate(1)"
            Tab(0).Control(2)=   "fraPaymentDate(0)"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Quarterly Payment Dates"
            TabPicture(1)   =   "frmClientNew.frx":0E3E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraPaymentDate(3)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Half Yearly payments"
            TabPicture(2)   =   "frmClientNew.frx":0E5A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fraPaymentDate(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Yearly payments"
            TabPicture(3)   =   "frmClientNew.frx":0E76
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "fraPaymentDate(5)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            Begin VB.Frame Frame8 
               Caption         =   "Yearly Payment Date"
               Height          =   975
               Left            =   -71317
               TabIndex        =   235
               Top             =   1440
               Width           =   2535
               Begin VB.ComboBox cboM7 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   960
                  TabIndex        =   237
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboD7 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   240
                  TabIndex        =   236
                  Top             =   360
                  Width           =   615
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Half Yearly Payment Dates"
               Height          =   1575
               Left            =   -71580
               TabIndex        =   228
               Top             =   1260
               Width           =   3135
               Begin VB.ComboBox cboD5 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   232
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboM5 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   231
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboD6 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   230
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboM6 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   229
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Second"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   234
                  Top             =   840
                  Width           =   615
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   "First"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   233
                  Top             =   360
                  Width           =   495
               End
            End
            Begin VB.Frame fraPaymentDate 
               Enabled         =   0   'False
               Height          =   2295
               Index           =   2
               Left            =   -68700
               TabIndex        =   223
               Top             =   420
               Width           =   3015
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   8
                  Left            =   720
                  TabIndex        =   119
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   8
                  Left            =   1440
                  TabIndex        =   120
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   9
                  Left            =   720
                  TabIndex        =   121
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   9
                  Left            =   1440
                  TabIndex        =   122
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   10
                  Left            =   720
                  TabIndex        =   123
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   10
                  Left            =   1440
                  TabIndex        =   124
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   11
                  Left            =   720
                  TabIndex        =   125
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   11
                  Left            =   1440
                  TabIndex        =   126
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "9th"
                  Height          =   195
                  Index           =   43
                  Left            =   390
                  TabIndex        =   227
                  Top             =   420
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "10th"
                  Height          =   195
                  Index           =   44
                  Left            =   300
                  TabIndex        =   226
                  Top             =   900
                  Width           =   315
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "11th"
                  Height          =   195
                  Index           =   45
                  Left            =   300
                  TabIndex        =   225
                  Top             =   1380
                  Width           =   315
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "12th"
                  Height          =   195
                  Index           =   46
                  Left            =   300
                  TabIndex        =   224
                  Top             =   1860
                  Width           =   315
               End
            End
            Begin VB.Frame fraPaymentDate 
               Enabled         =   0   'False
               Height          =   2295
               Index           =   1
               Left            =   -71700
               TabIndex        =   218
               Top             =   420
               Width           =   3015
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   4
                  Left            =   720
                  TabIndex        =   111
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   4
                  Left            =   1440
                  TabIndex        =   112
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   5
                  Left            =   720
                  TabIndex        =   113
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   5
                  Left            =   1440
                  TabIndex        =   114
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   6
                  Left            =   720
                  TabIndex        =   115
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   6
                  Left            =   1440
                  TabIndex        =   116
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   7
                  Left            =   720
                  TabIndex        =   117
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   118
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "7th"
                  Height          =   195
                  Index           =   41
                  Left            =   390
                  TabIndex        =   222
                  Top             =   1380
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "5th"
                  Height          =   195
                  Index           =   39
                  Left            =   390
                  TabIndex        =   221
                  Top             =   420
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "6th"
                  Height          =   195
                  Index           =   40
                  Left            =   390
                  TabIndex        =   220
                  Top             =   900
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "8th"
                  Height          =   195
                  Index           =   42
                  Left            =   390
                  TabIndex        =   219
                  Top             =   1860
                  Width           =   225
               End
            End
            Begin VB.Frame Frame12 
               Caption         =   "Quarterly Payment Dates"
               Height          =   2295
               Left            =   -72180
               TabIndex        =   205
               Top             =   960
               Width           =   3015
               Begin VB.ComboBox cboD1 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   213
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboM1 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   212
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboD2 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   211
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboM2 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   210
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.ComboBox cboD3 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   209
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.ComboBox cboM3 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   208
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.ComboBox cboD4 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   840
                  TabIndex        =   207
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.ComboBox cboM4 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   206
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  Caption         =   "First"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   217
                  Top             =   360
                  Width           =   375
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Second"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   216
                  Top             =   840
                  Width           =   615
               End
               Begin VB.Label Label46 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Third"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   215
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.Label Label47 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Fourth"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   214
                  Top             =   1800
                  Width           =   615
               End
            End
            Begin VB.Frame fraPaymentDate 
               Enabled         =   0   'False
               Height          =   2295
               Index           =   0
               Left            =   -74640
               TabIndex        =   200
               Top             =   420
               Width           =   3015
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   110
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   3
                  Left            =   720
                  TabIndex        =   109
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   108
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   2
                  Left            =   720
                  TabIndex        =   107
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   106
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   1
                  Left            =   720
                  TabIndex        =   105
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboMonth 
                  Height          =   315
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   104
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboDay 
                  Height          =   315
                  Index           =   0
                  Left            =   720
                  TabIndex        =   103
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "4th"
                  Height          =   195
                  Index           =   38
                  Left            =   390
                  TabIndex        =   204
                  Top             =   1860
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "3rd"
                  Height          =   195
                  Index           =   37
                  Left            =   390
                  TabIndex        =   203
                  Top             =   1380
                  Width           =   225
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "2nd"
                  Height          =   195
                  Index           =   36
                  Left            =   345
                  TabIndex        =   202
                  Top             =   900
                  Width           =   270
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "1st"
                  Height          =   255
                  Index           =   35
                  Left            =   240
                  TabIndex        =   201
                  Top             =   420
                  Width           =   375
               End
            End
            Begin VB.Frame fraPaymentDate 
               Caption         =   "Quarterly Payment Dates"
               Enabled         =   0   'False
               Height          =   2295
               Index           =   3
               Left            =   -72120
               TabIndex        =   195
               Top             =   600
               Width           =   3015
               Begin VB.ComboBox cboQDay 
                  Height          =   315
                  Index           =   0
                  Left            =   720
                  TabIndex        =   127
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboQMth 
                  Height          =   315
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   128
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboQDay 
                  Height          =   315
                  Index           =   1
                  Left            =   720
                  TabIndex        =   129
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboQMth 
                  Height          =   315
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   130
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.ComboBox cboQDay 
                  Height          =   315
                  Index           =   2
                  Left            =   720
                  TabIndex        =   131
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.ComboBox cboQMth 
                  Height          =   315
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   132
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.ComboBox cboQDay 
                  Height          =   315
                  Index           =   3
                  Left            =   720
                  TabIndex        =   133
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.ComboBox cboQMth 
                  Height          =   315
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   134
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "First"
                  Height          =   195
                  Index           =   47
                  Left            =   330
                  TabIndex        =   199
                  Top             =   360
                  Width           =   285
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Second"
                  Height          =   195
                  Index           =   48
                  Left            =   60
                  TabIndex        =   198
                  Top             =   840
                  Width           =   555
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Third"
                  Height          =   195
                  Index           =   49
                  Left            =   255
                  TabIndex        =   197
                  Top             =   1320
                  Width           =   360
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Fourth"
                  Height          =   195
                  Index           =   50
                  Left            =   165
                  TabIndex        =   196
                  Top             =   1800
                  Width           =   450
               End
            End
            Begin VB.Frame fraPaymentDate 
               Caption         =   "Half Yearly Payment Dates"
               Enabled         =   0   'False
               Height          =   1575
               Index           =   4
               Left            =   -72240
               TabIndex        =   192
               Top             =   960
               Width           =   3135
               Begin VB.ComboBox cboHDay 
                  Height          =   315
                  Index           =   0
                  Left            =   840
                  TabIndex        =   135
                  Top             =   360
                  Width           =   615
               End
               Begin VB.ComboBox cboHMth 
                  Height          =   315
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   136
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboHDay 
                  Height          =   315
                  Index           =   1
                  Left            =   840
                  TabIndex        =   137
                  Top             =   840
                  Width           =   615
               End
               Begin VB.ComboBox cboHMth 
                  Height          =   315
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   138
                  Top             =   840
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Second"
                  Height          =   195
                  Index           =   52
                  Left            =   180
                  TabIndex        =   194
                  Top             =   840
                  Width           =   555
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "First"
                  Height          =   195
                  Index           =   51
                  Left            =   450
                  TabIndex        =   193
                  Top             =   360
                  Width           =   285
               End
            End
            Begin VB.Frame fraPaymentDate 
               Caption         =   "Yearly Payment Date"
               Enabled         =   0   'False
               Height          =   975
               Index           =   5
               Left            =   2760
               TabIndex        =   190
               Top             =   1080
               Width           =   3135
               Begin VB.ComboBox cboYMth 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   140
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.ComboBox cboYDay 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   139
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Once:"
                  Height          =   195
                  Index           =   67
                  Left            =   210
                  TabIndex        =   191
                  Top             =   360
                  Width           =   435
               End
            End
         End
         Begin MSForms.CheckBox chkLettingFee 
            Height          =   330
            Left            =   720
            TabIndex        =   83
            Top             =   360
            Width           =   1995
            VariousPropertyBits=   746588185
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "3519;582"
            Value           =   "0"
            Caption         =   "Letting Fee Payable"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkMngFee 
            Height          =   330
            Left            =   -74280
            TabIndex        =   89
            Top             =   360
            Width           =   2715
            VariousPropertyBits=   746588185
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "4789;582"
            Value           =   "0"
            Caption         =   "Management Fee Payable"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkRentPayble 
            Height          =   255
            Left            =   -74280
            TabIndex        =   95
            Top             =   360
            Width           =   1455
            VariousPropertyBits=   746588185
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2566;450"
            Value           =   "0"
            Caption         =   "Rent Payable"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "Issue Fee/Charges:                         (days)"
            Height          =   195
            Index           =   0
            Left            =   -74520
            TabIndex        =   239
            Top             =   480
            Width           =   2940
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            Caption         =   "Issue Payable:                       (days)"
            Height          =   195
            Index           =   13
            Left            =   -67800
            TabIndex        =   238
            Top             =   480
            Width           =   2490
         End
      End
      Begin MSComctlLib.TreeView tvwLandLord 
         Height          =   4815
         Left            =   -74920
         TabIndex        =   301
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8493
         _Version        =   393217
         Indentation     =   441
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   302
         Top             =   1920
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5636
         _Version        =   393216
         BackColorFixed  =   12632256
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.MonthView dtDate 
         Height          =   2370
         Left            =   -70200
         TabIndex        =   303
         Top             =   1080
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16768960
         Appearance      =   1
         StartOfWeek     =   57016322
         CurrentDate     =   38637
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   -70560
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew.frx":0E92
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew.frx":176C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew.frx":2046
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClientNew.frx":2920
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Left            =   -71400
         TabIndex        =   306
         Top             =   405
         Width           =   630
      End
      Begin MSForms.Label lblImageName 
         Height          =   195
         Left            =   -66360
         TabIndex        =   305
         Top             =   480
         Width           =   3120
         Caption         =   "Image Name:"
         Size            =   "5503;344"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdImgLeftMove 
         Height          =   315
         Left            =   -65040
         TabIndex        =   304
         Top             =   3960
         Width           =   555
         Size            =   "979;556"
         Picture         =   "frmClientNew.frx":2C3A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Image imgPremises 
         BorderStyle     =   1  'Fixed Single
         Height          =   3090
         Left            =   -66360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3090
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   380
      Left            =   10800
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveClient 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   380
      Left            =   4392
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteClient 
      Caption         =   "&Delete"
      Height          =   380
      Left            =   8664
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditClient 
      Caption         =   "&Edit"
      Height          =   380
      Left            =   2256
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelChange 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6528
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNewClient 
      Caption         =   "&New"
      Height          =   380
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   11865
      TabIndex        =   66
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton cmdResidency 
         Caption         =   "V"
         Enabled         =   0   'False
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
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdClient 
         Caption         =   "V"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox txtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   2355
      End
      Begin VB.TextBox txtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2620
      End
      Begin VB.TextBox txtYearEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2000
      End
      Begin VB.TextBox txtVATReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox txtResidency 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   120
         Width           =   2000
      End
      Begin VB.TextBox txtAcBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9285
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   2000
      End
      Begin VB.ListBox lstResidency 
         Height          =   450
         ItemData        =   "frmClientNew.frx":308C
         Left            =   5640
         List            =   "frmClientNew.frx":3096
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   2000
      End
      Begin MSForms.ComboBox cboLandLordSageSuppAC 
         Height          =   285
         Left            =   1605
         TabIndex        =   4
         Top             =   1200
         Width           =   2610
         VariousPropertyBits=   746604575
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4604;503"
         TextColumn      =   1
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1762;4233"
      End
      Begin MSForms.ComboBox cboLandLordSageCustAC 
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   840
         Width           =   2610
         VariousPropertyBits=   746604575
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4604;503"
         TextColumn      =   1
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763;4233"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year End:"
         Height          =   195
         Index           =   7
         Left            =   7800
         TabIndex        =   153
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX/VAT Number:"
         Height          =   195
         Index           =   6
         Left            =   7800
         TabIndex        =   152
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Balance:"
         Height          =   195
         Index           =   5
         Left            =   7800
         TabIndex        =   151
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Residency:"
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   150
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Supplier A/C:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   149
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sage Customer A/C:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   148
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   147
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client/Landlord ID:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   146
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Caption         =   "Main"
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
   Begin VB.PictureBox picClientList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1320
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   154
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2175
         Left            =   45
         TabIndex        =   155
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   4403
      ScaleHeight     =   390
      ScaleWidth      =   3255
      TabIndex        =   157
      Top             =   3787
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         TabIndex        =   158
         Top             =   90
         Width           =   3075
      End
   End
   Begin MSForms.Label Label3 
      Height          =   100
      Left            =   120
      TabIndex        =   159
      Top             =   2400
      Width           =   11895
      BackColor       =   8388608
      Size            =   "20981;176"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmClientNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDefaultAccount As Boolean
Private szPropertyID As String
Private iRecharge As Integer
Private bGlobalData As Boolean
Private bNewEdit As Boolean
Private IMAGE_FILE_NAME_ As String
Private szaPremisisIDType() As String

Private Sub cboBank_ID_Click()
   txtBANK_NAME.text = cboBank_ID.Column(1)
   txtBANK_ADDRESS1.text = cboBank_ID.Column(3)
   txtBANK_ADDRESS2.text = cboBank_ID.Column(5)
   txtBANK_ADDRESS3.text = cboBank_ID.Column(6)
   txtBANK_POST_CODE.text = cboBank_ID.Column(4)
   txtBANK_SC.text = cboBank_ID.Column(2)
End Sub

Private Sub cboProperty_Click()
   ClearAllRecharge

   Dim sSQLQuery_ As String, sFilter As String
   Dim szaPropertyID() As String

   MousePointer = vbHourglass

   szaPropertyID = Split(cboProperty.text, " / ")
   szPropertyID = szaPropertyID(0)

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   sSQLQuery_ = "SELECT * " & _
                "FROM tlbAggreement " & _
                "WHERE CLIENT_ID = '" & txtClientID.text & "' And " & _
                     "PROPERTY_ID = '" & szPropertyID & "';"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If adoMain.Recordset.EOF Then
      MsgBox "There is no agreement record setup for this property. Please enter agreement details.", vbCritical + vbOKOnly, "No Agreement"
      LoockingAgrText False
      cboProperty.Locked = True
      cmdAgmntSave.Enabled = True
      optLetting_Fees(0).SetFocus
      MousePointer = vbDefault
      Exit Sub
   Else
      With adoMain
         txtClientID.text = .Recordset.Fields("CLIENT_ID").Value
         szPropertyID = .Recordset.Fields("PROPERTY_ID").Value
         txtSTART_DATE.text = .Recordset.Fields("START_DATE").Value
         txtEND_DATE.text = .Recordset.Fields("END_DATE").Value
         txtREVIEW_DATE.text = .Recordset.Fields("REVIEW_DATE").Value
         txtNOTICE_DAYS.text = .Recordset.Fields("NOTICE_DAYS").Value
         txtRRPA.text = .Recordset.Fields("RRPA").Value
         txtOWNERSHIP_PERCENT.text = .Recordset.Fields("OWNERSHIP_PERCENT").Value
         optLetting_Fees(CInt(.Recordset.Fields("LETTING_FEES").Value)).Value = True
         txtLETTING_FEES_VALUE(CInt(.Recordset.Fields("LETTING_FEES").Value)).text = .Recordset.Fields("LETTING_FEES_VALUE").Value
         optManagement_Fees(CInt(.Recordset.Fields("MGT_FEES").Value)).Value = True
         txtMGT_FEES_VALUE(CInt(.Recordset.Fields("MGT_FEES").Value)).text = .Recordset.Fields("MGT_FEES_VALUE").Value
         optRecharges(CInt(.Recordset.Fields("RECHARGES").Value)).Value = True

         .Recordset.Close
      End With

      cmdAgmntEdit.Enabled = True
      cmdAgmntDelete.Enabled = True
   End If

   MousePointer = vbDefault
End Sub

Private Sub cboProperty_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = CboShowDown(cboProperty.hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub chkLettingFee_LostFocus()
   If chkLettingFee.Value = True And cboLettingChrgType.ListCount = 0 Then
      MsgBox "There is no 'Fee Charge Type' defined. Please input charge type in the Global Form.", vbCritical + vbOKOnly, "No Types"
      chkLettingFee.Value = False
   End If
End Sub

Private Sub chkMngFee_LostFocus()
   If chkMngFee.Value = True And cboMngChrgType.ListCount = 0 Then
      MsgBox "There is no 'Fee Charge Type' defined. Please input charge type in the Global Form.", vbCritical + vbOKOnly, "No Types"
      chkMngFee.Value = False
   End If
End Sub

Private Sub chkRentPayble_LostFocus()
   If chkRentPayble.Value = True And cboRentChrgType.ListCount = 0 Then
      MsgBox "There is no 'Payable Type' defined. Please input payable type in the Global Form.", vbCritical + vbOKOnly, "No Types"
      chkRentPayble.Value = False
   End If
End Sub

Private Sub cmdCancelBank_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   
   CommandButtonEnabled True
   LockingAcText True
   NewBankText True, True
   flxOtherBankDetails_RowColChange
   cmdNewBank.Visible = False
End Sub

Private Sub cmdClientDetailsCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText True
   CommandButtonEnable True
End Sub

Private Sub CommandButtonEnable(bEnable As Boolean)
   cmdClientDetailsEdit.Enabled = bEnable
   cmdClientDetailsSave.Enabled = Not bEnable
   cmdClientDetailsCancel.Enabled = Not bEnable
End Sub

Private Sub cmdClientDetailsEdit_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to edit?", vbQuestion + vbYesNo, "Edit Details") = vbNo Then Exit Sub
   LockingAllText False
   CommandButtonEnable False
End Sub

Private Sub cmdClientDetailsSave_Click()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM Client " & _
           "WHERE ClientID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   
   With rstClient
      .Edit
      !ClientAddressLine1 = txtClientAddressLine1.text
      !ClientAddressLine2 = txtClientAddressLine2.text
      !ClientAddressLine3 = txtClientAddressLine3.text
      !ClientPostCode = txtClientPostCode.text
      !ClientOfficeEmail = txtClientOfficeEmail.text
      !ClientPersonalEmail = txtClientPersonalEmail.text
      !ClientHomeTel = txtClientHomeTel.text
      !ClientMobile = txtClientMobile.text
      !ClientOfficeAddressLine1 = txtClientOfficeAddressLine1.text
      !ClientOfficeAddressLine2 = txtClientOfficeAddressLine2.text
      !ClientOfficeAddressLine3 = txtClientOfficeAddressLine3.text
      !ClientOfficePostCode = txtClientOfficePostCode.text
      !ClientOfficeTel = txtClientOfficeTel.text

      .Update
      .Close
   End With
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   
   MsgBox "Data has been updated successfully", vbInformation + vbOKOnly, "Data Update"
   CommandButtonEnable True
End Sub

Private Sub cmdClinetAddAtch_Click()
   If MsgBox("Do you want to add new file?", vbQuestion + vbYesNo, "Attachment") = vbNo Then Exit Sub
   AddNewAttachment cmbFiles, "Client", txtClientID.text
   MsgBox "File has been saved successfull, Thanks"
End Sub

Private Sub cmdDeleteBank_Click()
   If MsgBox("Do you want to delete current account details?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
   
   flxOtherBankDetails.RemoveItem (flxOtherBankDetails.Row)

   flxOtherBankDetails_RowColChange
   NewBankText True, False
   cmdNewBank.Caption = "New"
   LockingAcText True
   MsgBox "Record has been deleted successfully.", vbInformation + vbOKOnly, "Delete"
End Sub

Private Sub cmdDeleteFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   If MsgBox("Are you sure to delete " & cmbFiles.text & "?", vbQuestion + vbYesNo, "Delete File") = vbNo Then Exit Sub
   DeleteAttachment cmbFiles, cmbFiles.Column(2), txtClientID.text, "Client"
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdEditBank_Click()
   MousePointer = vbHourglass

   cmdNewBank.Caption = "Edit"

   cmdNewBank.Visible = True
   LockingAcText False

   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
   MousePointer = vbDefault
End Sub

Private Sub cmdAddNewBank_Click()
   If MsgBox("Is it default account?", vbQuestion + vbYesNo, "Deafult Account") = vbYes Then
      bDefaultAccount = True
   Else
      bDefaultAccount = False
   End If

   MousePointer = vbHourglass

   PopulateBank
   cmdNewBank.Caption = "New"
   cmdNewBank.Visible = True
   cboBank_ID.SetFocus

   LockingAcText False
   NewBankText True, True
   cboBank_ID.Locked = False
   
   CommandButtonEnabled False
   flxOtherBankDetails.Row = flxOtherBankDetails.Rows - 1
   MousePointer = vbDefault
End Sub

Private Sub CommandButtonEnabled(bEnable As Boolean)
   cmdAddNewBank.Enabled = bEnable
   cmdEditBank.Enabled = bEnable
   cmdDeleteBank.Enabled = bEnable
   cmdSaveBank.Enabled = Not bEnable
   cmdCancelBank.Enabled = Not bEnable
   flxOtherBankDetails.Enabled = bEnable
End Sub

Public Function PopulateBank()
   Dim sSQLQuery_ As String

   adoBank.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT BANK_ID, BANK_NAME, SORT_CODE, " & _
                     "BANK_ADDRESS1, BANK_POST_CODE, " & _
                     "BANK_ADDRESS2, BANK_ADDRESS3 " & _
                "FROM tlbBank"

   adoBank.RecordSource = sSQLQuery_
   adoBank.CommandType = adCmdText
   adoBank.Refresh

   Dim TotalRow, TotalCol As Integer

   TotalRow = adoBank.Recordset.RecordCount
   TotalCol = adoBank.Recordset.Fields.Count

   Dim data() As String

   ReDim data(TotalCol, TotalRow) As String

   Dim i, j As Integer

   For i = 0 To adoBank.Recordset.RecordCount - 1
       For j = 0 To adoBank.Recordset.Fields.Count - 1
           data(j, i) = adoBank.Recordset.Fields(j)
       Next j
       adoBank.Recordset.MoveNext
   Next i

   cboBank_ID.Column() = data()
End Function

Private Sub cmdAddNewClient_Click()
   If MsgBox("Do you want to add new client?", vbYesNo + vbQuestion, "Add New Client") = vbNo Then Exit Sub
   bNewEdit = True
   
   MousePointer = vbHourglass

   SageCustomerAccCombo
   SageSupplierAccCombo

   UnlockMainClientText True
   MainCommandButtonEnable True

   txtClientID.SetFocus
   
   MousePointer = vbDefault
End Sub

Private Sub SageCustomerAccCombo()
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
    
         Dim TotalRow, TotalCol As Long
         Dim data() As String
         Dim i As Integer
             
         TotalRow = oSalesRecord.Count
         TotalCol = 2
         cboLandLordSageCustAC.Clear
         
         ReDim data(TotalCol, TotalRow) As String
         
         oSalesRecord.MoveFirst
         For i = 0 To TotalRow - 1
            'cboTest.AddItem adoClient.Recordset.Fields(1)
            data(0, i) = CStr(oSalesRecord.Fields.Item("ACCOUNT_REF").Value)
            data(1, i) = CStr(oSalesRecord.Fields.Item("NAME").Value)
            oSalesRecord.MoveNext
         Next i
         '
         cboLandLordSageCustAC.Column() = data()
         cboLandLordSageCustAC.ColumnCount = TotalCol
         cboLandLordSageCustAC.BoundColumn = 1
'         cboLandLordSageCustAC.TextColumn = 2

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

Private Sub UnlockMainClientText(bUnlock As Boolean)
   txtClientID.Locked = Not bUnlock
   txtClientName.Locked = Not bUnlock
   cboLandLordSageCustAC.Locked = Not bUnlock
   cboLandLordSageSuppAC.Locked = Not bUnlock
   txtAcBalance.Locked = Not bUnlock
   txtVATReg.Locked = Not bUnlock
   txtYearEndDate.Locked = Not bUnlock
End Sub

Private Sub cmdAgmntEdit_Click()
   If MsgBox("Do you want to edit the agreement?", vbQuestion + vbYesNo, "Edit Agreement") = vbNo Then Exit Sub

   LoockingAgrText False
   cboProperty.Locked = True
End Sub

Private Sub cmdAgmntSave_Click()
   If MsgBox("Are you sure to save?", vbQuestion + vbYesNo, "Data Saving") = vbNo Then Exit Sub

   MousePointer = vbHourglass

   Dim conAgr As New RDO.rdoConnection
   Dim rstAgr As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conAgr.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conAgr.CursorDriver = rdUseIfNeeded
   conAgr.EstablishConnection rdDriverNoPrompt

   szSQL = "DELETE * " & _
           "FROM tlbAggreement " & _
           "WHERE CLIENT_ID = '" & txtClientID.text & "' AND " & _
               "PROPERTY_ID = '" & szPropertyID & "';"
'Debug.Print szSQL
   Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)
   rstAgr.Close
   
   szSQL = "SELECT * " & _
           "FROM tlbAggreement"
   Set rstAgr = conAgr.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   
   With rstAgr
      .AddNew
      rstAgr!CLIENT_ID = txtClientID.text
      rstAgr!PROPERTY_ID = szPropertyID
      rstAgr!AGG_DATE = Format(Now, "DD MMMM YYYY")
      rstAgr!START_DATE = Format(txtSTART_DATE.text, "DD MMMM YYYY")
      rstAgr!END_DATE = Format(txtEND_DATE.text, "DD MMMM YYYY")
      rstAgr!REVIEW_DATE = Format(txtREVIEW_DATE.text, "DD MMMM YYYY")
      rstAgr!NOTICE_DAYS = CInt(txtNOTICE_DAYS.text)
      rstAgr!RRPA = CCur(txtRRPA.text)
      rstAgr!OWNERSHIP_PERCENT = CSng(txtOWNERSHIP_PERCENT.text)
      rstAgr!LETTING_FEES = CStr(IIf(optLetting_Fees(0), 0, 1))
      rstAgr!LETTING_FEES_VALUE = CSng(txtLETTING_FEES_VALUE(CInt(rstAgr!LETTING_FEES)).text)
      rstAgr!MGT_FEES = CStr(IIf(optManagement_Fees(0), 0, 1))
      rstAgr!MGT_FEES_VALUE = CSng(txtMGT_FEES_VALUE(CInt(rstAgr!MGT_FEES)).text)
      rstAgr!RECHARGES = CStr(iRecharge)
      
      .Update
      .Close
   End With
   Set rstAgr = Nothing
   
   conAgr.Close
   Set conAgr = Nothing
   MousePointer = vbDefault
   
   cboProperty.Locked = False
   MsgBox "Agreement has been updated successfully.", vbInformation + vbOKOnly, "Agreement"
   LoockingAgrText True
   Exit Sub
   
ErrorHandler:

   rstAgr.Close
   Set rstAgr = Nothing
   conAgr.Close
   Set conAgr = Nothing
   
   MsgBox ERR.Number & ERR.description & " ", vbCritical + vbOK, "PCM Error: 125"
End Sub

Private Sub cmdClient_Click()
   Call PrepareList

   picClientList.Top = picMain.Top + txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = picMain.Left + txtClientID.Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxClientList
   LoadAllClientFlxGrd
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDeleteClient_Click()
'===========================================================================================
'This button should be visible, because user should not get facility to delete any record.
'we should give user a facility to see or remove the recode from the current list.
'===========================================================================================
   If txtClientID.text = "" Then
      MsgBox "Please select a client to delete.", vbInformation, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Are you sure to delete current client?", vbYesNo + vbInformation, "Confimation") = vbNo Then Exit Sub

   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT * " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If Not rstClient.EOF Then
      MsgBox "This client could not be deleted. This client has property in the database." & _
            (Chr$(13) + Chr$(10)) & "To delete this client delete all properties of this client.", vbCritical + vbOKOnly, "Delete Not Possible"
      rstClient.Close
      Set rstClient = Nothing
      conClient.Close
      Set conClient = Nothing
      Exit Sub
   End If

   rstClient.Close

   MousePointer = vbHourglass

   szSQL = "DELETE * FROM CLIENT WHERE CLIENTID = '" & txtClientID.text & "';"
   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   rstClient.Close
   Set rstClient = Nothing
   conClient.Close
   Set conClient = Nothing

   MsgBox "Client has been deleted successfully.", vbOKOnly + vbInformation, "Delete Confirmation"

   MousePointer = vbDefault
End Sub

Private Sub cmdEditClient_Click()
   If txtClientID.text = "" Then
      MsgBox "Please select a client to edit.", vbCritical + vbOKOnly, "No selection"
      txtClientID.SetFocus
      Exit Sub
   End If

   If MsgBox("Do you want to make change of the current client?", vbYesNo + vbQuestion, "Edit Client") = vbNo Then Exit Sub
   bNewEdit = False

   MousePointer = vbHourglass

   MainCommandButtonEnable True
   
   Dim szTemp As String
   
   If cboLandLordSageSuppAC.ListCount = 0 Then
      szTemp = cboLandLordSageSuppAC.text
      SageSupplierAccCombo
      cboLandLordSageSuppAC.text = szTemp
   End If
   If cboLandLordSageCustAC.ListCount = 0 Then
      szTemp = cboLandLordSageCustAC.text
      SageCustomerAccCombo
      cboLandLordSageCustAC.text = szTemp
   End If
      
   LockingAllText False
   UnlockMainClientText True
   
   MousePointer = vbDefault
End Sub

Private Sub MainCommandButtonEnable(bEnabled As Boolean)
   cmdAddNewClient.Enabled = Not bEnabled
   cmdEditClient.Enabled = Not bEnabled
   cmdSaveClient.Enabled = bEnabled
   cmdDeleteClient.Enabled = Not bEnabled
   cmdCancelChange.Enabled = bEnabled
   
   cmdClient.Enabled = Not bEnabled
   cmdResidency.Enabled = bEnabled
End Sub

Private Sub cmdGridUnitLookup_Click()
   picClientList.Visible = False
End Sub

Private Sub cmdGSCancel_Click()
   If MsgBox("Do you want to cancel changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub

   Dim i As Integer
   
   On Error Resume Next
   For i = 0 To 67
      Label1(i).ForeColor = vbBlack
   Next i
   
   EnableGlobalControl False
End Sub

Private Sub cmdGSEdit_Click()
   MousePointer = vbHourglass

   EnableGlobalControl True

   MousePointer = vbDefault
End Sub

Private Sub EnableGlobalControl(bEnable As Boolean)
   Dim i As Integer
   
   chkLettingFee.Enabled = bEnable
   chkMngFee.Enabled = bEnable
   chkRentPayble.Enabled = bEnable
   For i = 0 To 5
      If i < 3 Then fraFee(i).Enabled = bEnable
      If i < 6 Then fraPaymentDate(i).Enabled = bEnable
   Next i
   
   cmdGSEdit.Enabled = Not bEnable
   cmdGSSave.Enabled = bEnable
   cmdGSCancel.Enabled = bEnable
End Sub

Private Function ControlFilled() As Boolean
   ControlFilled = True
   
   If txtFeeIsuDays.text = "" Then
      Label83(0).ForeColor = vbRed
      ControlFilled = False
   End If
   If txtPayIsuDays.text = "" Then
      Label83(13).ForeColor = vbRed
      ControlFilled = False
   End If
   If chkLettingFee.Value Then
      If cboLettingAM.text = "" Then
         Label1(23).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboLettingFreq.text = "" Then
         Label1(25).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtLettingNtDueDt.text = "" Then
         Label1(24).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtLettingStDt.text = "" Then
         Label1(26).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboLettingChrgType.text = "" Then
         Label1(64).ForeColor = vbRed
         ControlFilled = False
      End If
   End If

   If chkMngFee.Value Then
      If cboMngAM.text = "" Then
         Label1(27).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboMngFreq.text = "" Then
         Label1(29).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtMngNtDueDt.text = "" Then
         Label1(28).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtMngStDt.text = "" Then
         Label1(30).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboMngChrgType.text = "" Then
         Label1(65).ForeColor = vbRed
         ControlFilled = False
      End If
   End If

   If chkRentPayble.Value Then
      If cboRentAM.text = "" Then
         Label1(31).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboRentFreq.text = "" Then
         Label1(32).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtRentNtDueDt.text = "" Then
         Label1(33).ForeColor = vbRed
         ControlFilled = False
      End If
      If txtRentStDt.text = "" Then
         Label1(34).ForeColor = vbRed
         ControlFilled = False
      End If
      If cboRentChrgType.text = "" Then
         Label1(66).ForeColor = vbRed
         ControlFilled = False
      End If
   End If

   Dim i As Integer
   
   For i = 0 To 11         'MONTHLY
      If cboDay(i).text = "" Or cboMonth(i).text = "" Then
         Label1(35 + i).ForeColor = vbRed
         ControlFilled = False
      End If
      If i < 4 Then        'QUARTERLY
         If cboQDay(i).text = "" Or cboQMth(i).text = "" Then
            Label1(47 + i).ForeColor = vbRed
            ControlFilled = False
         End If
      End If
      If i < 2 Then        'HALF YEARLY
         If cboHDay(i).text = "" Or cboHMth(i).text = "" Then
            Label1(51 + i).ForeColor = vbRed
            ControlFilled = False
         End If
      End If
   Next i
   If cboYDay.text = "" Or cboYMth.text = "" Then        'YEARLY
      Label1(67).ForeColor = vbRed
      ControlFilled = False
   End If
End Function

Private Sub cmdGSSave_Click()
   Dim sSQLQuery_ As String, i As Integer

   If Not ControlFilled Then
      MsgBox "Please type/select red marked fields.", vbCritical + vbOKOnly, "Blank Fields"
      Exit Sub
   End If

   MousePointer = vbHourglass

'Saving client's global data
   If cmdGSEdit.Caption = "Edit" Then
      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"
   Else
      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData;"
   End If

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If cmdGSEdit.Caption = "Add New" Then
      adoMain.Recordset.AddNew
   End If

   adoMain.Recordset.Fields("ClientID").Value = txtClientID.text
   If chkLettingFee.Value Then
      adoMain.Recordset.Fields("LettingFee").Value = True
      adoMain.Recordset.Fields("LettingAM").Value = cboLettingAM.text
      adoMain.Recordset.Fields("LettingFreq").Value = cboLettingFreq.text
      adoMain.Recordset.Fields("LettingNtDueDt").Value = CDate(txtLettingNtDueDt.text)
      adoMain.Recordset.Fields("LettingStDt").Value = CDate(txtLettingStDt.text)
      adoMain.Recordset.Fields("LettingChrgType").Value = CInt(cboLettingChrgType.text)
   End If
   If chkMngFee.Value Then
      adoMain.Recordset.Fields("MngFee").Value = True
      adoMain.Recordset.Fields("MngAM").Value = cboMngAM.text
      adoMain.Recordset.Fields("MngFreq").Value = cboMngFreq.text
      adoMain.Recordset.Fields("MngNtDueDt").Value = CDate(txtMngNtDueDt.text)
      adoMain.Recordset.Fields("MngStDt").Value = CDate(txtMngStDt.text)
      adoMain.Recordset.Fields("MngChrgType").Value = CInt(cboMngChrgType.text)
   End If
   If chkRentPayble.Value Then
      adoMain.Recordset.Fields("RentPayble").Value = True
      adoMain.Recordset.Fields("RentAM").Value = cboRentAM.text
      adoMain.Recordset.Fields("RentFreq").Value = cboRentFreq.text
      adoMain.Recordset.Fields("RentNtDueDt").Value = CDate(txtRentNtDueDt.text)
      adoMain.Recordset.Fields("RentStDt").Value = CDate(txtRentStDt.text)
      adoMain.Recordset.Fields("RentChrgType").Value = CInt(cboRentChrgType.text)
   End If
   
   For i = 0 To 11
      adoMain.Recordset.Fields("MonthlyDueDate" & (i + 1) & "").Value = _
                        cboDay(i).text & " " & cboMonth(i).text
      If i < 4 Then _
         adoMain.Recordset.Fields("QuarterlyDueDate" & (i + 1) & "").Value = _
                        cboQDay(i).text & " " & cboQMth(i).text
      If i < 2 Then _
         adoMain.Recordset.Fields("HalfYearlyDueDate" & (i + 1) & "").Value = _
                        cboHDay(i).text & " " & cboHMth(i).text
   Next i
   adoMain.Recordset.Fields("YearlyDueDate").Value = _
                     cboYDay.text & " " & cboYMth.text

   adoMain.Recordset.Fields("FeeIsuDays").Value = CInt(txtFeeIsuDays.text)
   adoMain.Recordset.Fields("PayIsuDays").Value = CInt(txtPayIsuDays.text)
   
   adoMain.Recordset.Update
   adoMain.Recordset.Close

   MousePointer = vbDefault

   If cmdGSEdit.Caption = "Edit" Then
      MsgBox "Data has been updated successfully", vbInformation + vbOKOnly, "Update Data"
   Else
      MsgBox "Data has been saved successfully", vbInformation + vbOKOnly, "Save Data"
   End If

   EnableGlobalControl False
End Sub

Private Sub cmdImgDelete_Click()
   If imgPremises.Picture = 0 Then Exit Sub
   If MsgBox("Are you sure to delete the image?", vbQuestion + vbYesNo, "Delete Image") = vbNo Then Exit Sub
   DeleteImage imgPremises, IMAGE_FILE_NAME_, szaPremisisIDType(0), szaPremisisIDType(1)
   MsgBox "File has been deleted succussfully", vbInformation + vbOKOnly, "Delete File"
End Sub

Private Sub cmdImgLeftMove_Click()
   IMAGE_FILE_NAME_ = MoveNextImage(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), IMAGE_FILE_NAME_, lblImageName)
End Sub

Private Sub cmdNewBank_Click()
   If cmdNewBank.Caption = "New" Then
      NewBankText False, True
      cboBank_ID.Locked = False
      cboBank_ID.Clear
      cboBank_ID.SetFocus
   Else
      NewBankText False, False
      txtBANK_NAME.SetFocus
   End If

   cmdNewBank.Enabled = False
End Sub

Private Sub NewBankText(bLock As Boolean, bNew As Boolean)
'   cboBank_ID.Locked = bLock
   txtBANK_NAME.Locked = bLock
   txtBANK_ADDRESS1.Locked = bLock
   txtBANK_ADDRESS2.Locked = bLock
   txtBANK_ADDRESS3.Locked = bLock
   txtBANK_POST_CODE.Locked = bLock

   If Not bNew Then Exit Sub
   cboBank_ID.text = ""
   txtBANK_NAME.text = ""
   txtBANK_ADDRESS1.text = ""
   txtBANK_ADDRESS2.text = ""
   txtBANK_ADDRESS3.text = ""
   txtBANK_POST_CODE.text = ""
End Sub

Private Sub cmdOpenFile_Click()
   If cmbFiles.text = "" Then Exit Sub
   MousePointer = vbHourglass
   
   If OpenFile(cmbFiles.Column(2), App.Path & "\" & cmbFiles.Column(1)) < 32 Then _
      MsgBox "File has been moved from original location.", vbExclamation
      
   MousePointer = vbDefault
End Sub

Private Sub cmdResidency_Click()
   lstResidency.Top = txtResidency.Top
   lstResidency.Left = txtResidency.Left
   lstResidency.Visible = True
   lstResidency.ZOrder 0
   lstResidency.SetFocus
End Sub

Private Sub cmdSaveBank_Click()
   If cmdNewBank.Caption = "New" Then
      flxOtherBankDetails.AddItem ""
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 1) = cboBank_ID.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 2) = txtBANK_NAME.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 3) = txtBANK_POST_CODE.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 4) = txtBank_AC_Name.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 5) = txtBANK_AC_NUM.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 6) = txtBANK_SC.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 7) = IIf(bDefaultAccount, "YES", "NO")
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 8) = txtBANK_ADDRESS1.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 9) = txtBANK_ADDRESS2.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 10) = txtBANK_ADDRESS3.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 11) = cboPaymentMethod.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Rows - 1, 12) = txtBacsRef.text
   Else
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 1) = cboBank_ID.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 2) = txtBANK_NAME.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 3) = txtBANK_POST_CODE.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 4) = txtBank_AC_Name.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5) = txtBANK_AC_NUM.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6) = txtBANK_SC.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = IIf(bDefaultAccount, "YES", "NO")
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 8) = txtBANK_ADDRESS1.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 9) = txtBANK_ADDRESS2.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 10) = txtBANK_ADDRESS3.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 11) = cboPaymentMethod.text
      flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 12) = txtBacsRef.text
   End If

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String, szWhere As String

   On Error GoTo ErrorHandler

   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt
'
   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "New" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank;"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
      rstBank.AddNew
      rstBank!BANK_ID = cboBank_ID.text
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update
'
      NewBankText True, False
      rstBank.Close
      cmdNewBank.Visible = False
   End If

   If Not cmdNewBank.Enabled And cmdNewBank.Caption = "Edit" Then
      'Set the RDO Connections to the dataset
      szSQL = "SELECT * " & _
              "FROM tlbBank " & _
              "WHERE BANK_ID = '" & cboBank_ID.text & "';"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
'
      rstBank.Edit
      rstBank!BANK_NAME = txtBANK_NAME.text
      rstBank!BANK_ADDRESS1 = txtBANK_ADDRESS1.text
      rstBank!BANK_ADDRESS2 = txtBANK_ADDRESS2.text
      rstBank!BANK_ADDRESS3 = txtBANK_ADDRESS3.text
      rstBank!BANK_POST_CODE = txtBANK_POST_CODE.text
      rstBank.Update
'
      rstBank.Close

      NewBankText True, False
      cmdNewBank.Visible = False
   End If

   If bDefaultAccount And cmdNewBank.Caption = "New" Then
      szSQL = "SELECT * " & _
              "FROM CLIENT " & _
              "WHERE CLIENTID = '" & txtClientID.text & "'"
      Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
      With rstBank
         .Edit
         !BANK_ID = cboBank_ID.text
         .Update
         .Close
      End With
   End If
   
   If cmdNewBank.Caption = "Edit" Then
      szWhere = " Where BANK_AC_NUM = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5) & "' And " & _
                     "BANK_SC = '" & flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6) & "';"
   Else
      szWhere = ""
   End If

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks" & szWhere
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)
   With rstBank
      If cmdNewBank.Caption = "New" Then
         .AddNew
      Else
         .Edit
      End If
      !CLIENT_ID = txtClientID.text
      !BANK_ID = cboBank_ID.text
      !Bank_AC_Name = txtBank_AC_Name.text
      !BANK_AC_NUM = txtBANK_AC_NUM.text
      !BANK_SC = txtBANK_SC.text
      !DEFAULT_AC = bDefaultAccount
      !PaymentMethod = cboPaymentMethod.text
      !BacsRef = txtBacsRef.text
      .Update
   End With

   If cmdNewBank.Caption = "New" Then
      MsgBox "Data has been saved successfully.", vbInformation + vbOKOnly, "Add New"
   Else
      MsgBox "Data has been updated successfully.", vbInformation + vbOKOnly, "Edit"
   End If
   CommandButtonEnabled True

NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub cmdSaveClient_Click()
   If MsgBox("Are you sure to save/update?", vbQuestion + vbYesNo, "Saving Data") = vbNo Then Exit Sub
   If txtClientID.text = "" Then
      MsgBox "Please type client id.", vbCritical + vbOKOnly, "Client"
      txtClientID.SetFocus
      Exit Sub
   End If
   If txtClientName.text = "" Then
      MsgBox "Please type client's name.", vbCritical + vbOKOnly, "Client"
      txtClientName.SetFocus
      Exit Sub
   End If
   If cboLandLordSageCustAC.text = "" Then
      MsgBox "Please select client's Sage Customer Account.", vbCritical + vbOKOnly, "Client"
      cboLandLordSageCustAC.SetFocus
      Exit Sub
   End If
   If cboLandLordSageSuppAC.text = "" Then
      MsgBox "Please select client's Sage Supplier Account.", vbCritical + vbOKOnly, "Client"
      cboLandLordSageSuppAC.SetFocus
      Exit Sub
   End If
   If txtResidency.text = "" Then
      MsgBox "Please select client's residency.", vbCritical + vbOKOnly, "Client"
      txtResidency.SetFocus
      Exit Sub
   End If
   If txtVATReg.text = "" Then
      MsgBox "Please type client's VAT Registration number.", vbCritical + vbOKOnly, "Client"
      txtVATReg.SetFocus
      Exit Sub
   End If
   If txtYearEndDate.text = "" Then
      MsgBox "Please type year end date.", vbCritical + vbOKOnly, "Client"
      txtYearEndDate.SetFocus
      Exit Sub
   End If
   
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   szSQL = "SELECT ClientID, ClientName, LandLordSageCustAC, " & _
                  "LandLordSageSuppAC, Residency, AcBalance, VATReg, " & _
                  "YearEndDate " & _
           "FROM Client " & _
           "WHERE ClientID = '" & txtClientID.text & "';"

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD="
   If PostToDBUsingADODB(Me, picMain, adoConn, szSQL, bNewEdit) Then
      MsgBox "Data has been saved succfully.", vbOKOnly, "Data Save"
   Else
      MsgBox "Data has not been saved.", vbOKOnly, "Data Save"
   End If
   UnlockMainClientText False
   MainCommandButtonEnable False
End Sub

Private Sub cmdUnitMemoCancel_Click()
   If MsgBox("Do you want to cancel the changes?", vbQuestion + vbYesNo, "Cancel") = vbNo Then Exit Sub
   MemoButtonEnable False
End Sub

Private Sub cmdUnitMemoEdit_Click()
   MemoButtonEnable True
End Sub

Private Sub cmdUnitMemoSave_Click()
   If SaveMemo("Client", "ClientMemo", txtClientID.text, "ClientID", txtNote) Then
      MsgBox "Memo has been saved successfully.", vbInformation + vbOKOnly, "Memo"
   End If
   MemoButtonEnable False
End Sub

Private Sub MemoButtonEnable(bEnable As Boolean)
   txtNote.Locked = Not bEnable
   cmdUnitMemoEdit.Enabled = Not bEnable
   cmdUnitMemoSave.Enabled = bEnable
   cmdUnitMemoCancel.Enabled = bEnable
End Sub

Private Sub cmdUploadImageAdd_Click()
   If MsgBox("Do you want to add new image?", vbQuestion + vbYesNo, "Image Attachment") = vbNo Then Exit Sub
   IMAGE_FILE_NAME_ = AddNewImage(imgPremises, szaPremisisIDType(1), szaPremisisIDType(0), lblImageName)
   MsgBox "Image has been uploaded successfull."
End Sub

Private Sub flxClientList_Click()
   Dim sSQLQuery_ As String, sFilter As String

   txtClientID.text = flxClientList.TextMatrix(flxClientList.Row, 1)

   MousePointer = vbHourglass
   fmeLoading.ZOrder 0
   fmeLoading.Visible = True
   fmeLoading.Refresh

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   sSQLQuery_ = "SELECT * " & _
                "FROM CLIENT " & _
                "WHERE CLIENT.ClientID = '" & flxClientList.TextMatrix(flxClientList.Row, 1) & "';"
'Debug.Print sSQLQuery_
   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      MsgBox "Error in Database.", vbExclamation
   Else
      LoadClientProperty

      lblLoading.Caption = "Please wait, tree is building..."
      fmeLoading.Refresh
      DrawLandLordTree tvwLandLord, imgList, txtClientID.text, True

      lblLoading.Caption = "Please wait, global data is loading..."
      LoadTypes
      fmeLoading.Refresh

      sSQLQuery_ = "SELECT * " & _
                   "FROM ClientGlobalData " & _
                   "WHERE ClientID = '" & txtClientID.text & "';"

      adoMain.RecordSource = sSQLQuery_
      adoMain.CommandType = adCmdText
      adoMain.Refresh
      If adoMain.Recordset.RecordCount > 0 Then
         Call LoadGlobalData
         bGlobalData = True
      Else
         bGlobalData = False
      End If

      RetrieveMemo "Client", "ClientMemo", txtClientID.text, "ClientID", txtNote
   End If

   fmeLoading.Visible = False
   MousePointer = vbDefault

   picClientList.Visible = False
End Sub

Private Sub LoadGlobalData()
   Dim szTemp() As String, i As Integer

   If adoMain.Recordset.Fields("LettingFee").Value Then
      chkLettingFee.Value = True
      cboLettingAM.text = adoMain.Recordset.Fields("LettingAM").Value
      cboLettingFreq.text = adoMain.Recordset.Fields("LettingFreq").Value
      txtLettingNtDueDt.text = adoMain.Recordset.Fields("LettingNtDueDt").Value
      txtLettingStDt.text = adoMain.Recordset.Fields("LettingStDt").Value
      cboLettingChrgType.text = adoMain.Recordset.Fields("LettingChrgType").Value
   End If
   If adoMain.Recordset.Fields("MngFee").Value Then
      chkMngFee.Value = True
      cboMngAM.text = adoMain.Recordset.Fields("MngAM").Value
      cboMngFreq.text = adoMain.Recordset.Fields("MngFreq").Value
      txtMngNtDueDt.text = adoMain.Recordset.Fields("MngNtDueDt").Value
      txtMngStDt.text = adoMain.Recordset.Fields("MngStDt").Value
      cboMngChrgType.text = adoMain.Recordset.Fields("MngChrgType").Value
   End If
   If adoMain.Recordset.Fields("RentPayble").Value Then
      chkRentPayble.Value = True
      cboRentAM.text = adoMain.Recordset.Fields("RentAM").Value
      cboRentFreq.text = adoMain.Recordset.Fields("RentFreq").Value
      txtRentNtDueDt.text = adoMain.Recordset.Fields("RentNtDueDt").Value
      txtRentStDt.text = adoMain.Recordset.Fields("RentStDt").Value
      cboRentChrgType.text = adoMain.Recordset.Fields("RentChrgType").Value
   End If

   For i = 0 To 11
      szTemp = Split(adoMain.Recordset.Fields("MonthlyDueDate" & (i + 1) & "").Value)
      cboDay(i).text = szTemp(0)
      cboMonth(i).text = szTemp(1)
      If i < 4 Then
         szTemp = Split(adoMain.Recordset.Fields("QuarterlyDueDate" & (i + 1) & "").Value)
         cboQDay(i).text = szTemp(0)
         cboQMth(i).text = szTemp(1)
      End If
      If i < 2 Then
         szTemp = Split(adoMain.Recordset.Fields("HalfYearlyDueDate" & (i + 1) & "").Value)
         cboHDay(i).text = szTemp(0)
         cboHMth(i).text = szTemp(1)
      End If
   Next i
   szTemp = Split(adoMain.Recordset.Fields("YearlyDueDate").Value)
   cboYDay.text = szTemp(0)
   cboYMth.text = szTemp(1)

   txtFeeIsuDays.text = adoMain.Recordset.Fields("FeeIsuDays").Value
   txtPayIsuDays.text = adoMain.Recordset.Fields("PayIsuDays").Value
End Sub

Private Sub LoadClientProperty()
   Dim conClient As New RDO.rdoConnection
   Dim rstProperty As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   cboProperty.Clear

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT PropertyID, PropertyName  " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyName;"

   Set rstProperty = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstProperty.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstProperty.EOF
      cboProperty.AddItem rstProperty!PROPERTYID & " / " & rstProperty!PropertyName
      rstProperty.MoveNext
   Wend
   
NoRes:
   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstProperty.Close
   conClient.Close
   Set rstProperty = Nothing
   Set conClient = Nothing
End Sub

Private Sub flxOtherBankDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxOtherBankDetails.ToolTipText = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.MouseRow, flxOtherBankDetails.MouseCol)
End Sub

Private Sub flxOtherBankDetails_RowColChange()
   Dim iCol As Integer
   
   MousePointer = vbHourglass

   cboBank_ID.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 1)
   txtBANK_NAME.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 2)
   txtBANK_POST_CODE.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 3)
   txtBank_AC_Name.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 4)
   txtBANK_AC_NUM.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 5)
   txtBANK_SC.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 6)
   bDefaultAccount = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", True, False)
   txtBANK_ADDRESS1.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 8)
   txtBANK_ADDRESS2.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 9)
   txtBANK_ADDRESS3.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 10)
   cboPaymentMethod.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 11)
   txtBacsRef.text = flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 12)
   fraBank(0).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   fraBank(1).Caption = IIf(flxOtherBankDetails.TextMatrix(flxOtherBankDetails.Row, 7) = "YES", "Default Account Details:", "Other Account Details:")
   MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50

   MousePointer = vbHourglass
   tabMain.Tab = 0
   cboPaymentMethod.AddItem "CHEQUE"
   cboPaymentMethod.AddItem "BACS"
   cboPaymentMethod.AddItem "DIRECT DEBIT"
   cboPaymentMethod.AddItem "Bank TRANSFER"
   cboPaymentMethod.AddItem "TT"
   cboPaymentMethod.AddItem "CHAPS"

   FillDaysMonths

   MousePointer = vbDefault
End Sub

Public Sub FillDaysMonths()

    Dim i As Integer, j As Integer
    
    For i = 0 To 11
      For j = 1 To 31
         cboDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboMonth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
   
   For i = 0 To 3
      For j = 1 To 31
         cboQDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboQMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
    
   For i = 0 To 1
      For j = 1 To 31
         cboHDay(i).AddItem Format(j, "00")
      Next j
      For j = 1 To 12
         cboHMth(i).AddItem Format("1/" & j & "/2000", "MMMM")
      Next j
   Next i
    
   For j = 1 To 31
      cboYDay.AddItem Format(j, "00")
   Next j
   For j = 1 To 12
      cboYMth.AddItem Format("1/" & j & "/2000", "MMMM")
   Next j
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub FlxDemandsConfigure(conFlxGrid As Control)
   Dim szHeader As String

   conFlxGrid.Cols = 4
   conFlxGrid.Clear
   szHeader$ = "|<ClientID|<ClientName|<ClientPostCode"
   conFlxGrid.FormatString = szHeader$
   conFlxGrid.ColWidth(0) = 300        'Solid column
   conFlxGrid.ColWidth(1) = 900        'Client ID
   conFlxGrid.ColWidth(2) = 3000       'Client Name
   conFlxGrid.ColWidth(3) = 800        'Post Code
   conFlxGrid.Rows = 2
'
   conFlxGrid.RowHeightMin = 300
End Sub

Private Sub imgClose_Click()
   picClientList.Visible = False
End Sub

Private Sub PropertyDetails(szID As String)
   Dim Conn As New RDO.rdoConnection
   Dim Rst As rdoResultset
   Dim szStr As String, szaTemp() As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   szStr = "SELECT * " & _
         "FROM PROPERTY " & _
         "WHERE PROPERTY.PROPERTYID='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)
   
   txtTVInfoName.text = Rst!PropertyName
   txtTVInfoAdd(0).text = Rst!ProAddressLine1
   txtTVInfoAdd(1).text = Rst!ProAddressLine2
   txtTVInfoAdd(2).text = Rst!ProAddressLine3
   txtTVInfoPC.text = Rst!PROPOSTCODE
   
   Rst.Close
   Set Rst = Nothing
   Conn.Close
   Set Conn = Nothing
End Sub

Private Function UnitDetails(szID As String) As Boolean
   Dim Conn As New RDO.rdoConnection
   Dim Rst As rdoResultset
   Dim szStr As String, szaTemp() As String

   Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn.CursorDriver = rdUseIfNeeded
   Conn.EstablishConnection rdDriverNoPrompt

   szStr = "SELECT * " & _
         "FROM UNITS " & _
         "WHERE UNITS.UnitNumber='" & szID & "';"
   Set Rst = Conn.OpenResultset(szStr, rdOpenStatic, rdConcurReadOnly)

   If Rst.EOF Then
      MsgBox "Error in Database, Please contact with vendor", vbCritical, "Serious Error"
   Else
      If Rst!UnitName <> "" Then txtTVInfoName.text = Rst!UnitName
      txtTVInfoAdd(0).text = IIf(Rst!UnitAddressLine1 <> "", Rst!UnitAddressLine1, "")
      txtTVInfoAdd(1).text = IIf(Rst!UnitAddressLine2 <> "", Rst!UnitAddressLine2, "")
      txtTVInfoAdd(2).text = IIf(Rst!UnitAddressLine3 <> "", Rst!UnitAddressLine3, "")
      txtTVInfoPC.text = IIf(Rst!UnitPostCode <> "", Rst!UnitPostCode, "")
      If Rst!OCCUPIED = "N" Then
         UnitDetails = False
         Exit Function
      Else
         lblTenantIDLink.Caption = Rst!SageAccountNumber
         lblTenantNameLink.Caption = IIf(IsNull(Rst!TenantCompanyName), "", Rst!TenantCompanyName)
         Rst.Close
         Conn.Close
         Set Rst = Nothing
         Set Conn = Nothing
'
         szStr = LeaseDetails(szID)
         If szStr = "NULL" Then
            MsgBox "Please update lease information of this unit.", vbInformation + vbOKOnly, "Error"
            UnitDetails = False
         Else
            UnitDetails = True
            szaTemp = Split(szStr, " # ")
'
            txtPreOccupiedFr.text = szaTemp(0)
            txtPreOccupiedTo.text = szaTemp(1)
            txtPreTenancyType.text = szaTemp(2)
            txtPreRentRvw.text = szaTemp(3)
         End If
      End If
   End If
End Function

Private Sub LoadAllClientFlxGrd()
   Dim conClient As New RDO.rdoConnection
   Dim rstClient As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   Set rstClient = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   If rstClient.EOF Then GoTo NoRes
   
   Dim iRow As Integer
   iRow = 1
   
   While Not rstClient.EOF
      flxClientList.TextMatrix(iRow, 1) = rstClient!ClientID
      flxClientList.TextMatrix(iRow, 2) = rstClient!ClientName
      flxClientList.TextMatrix(iRow, 3) = IIf(IsNull(rstClient!ClientPostCode), "", rstClient!ClientPostCode)
      rstClient.MoveNext
      If Not rstClient.EOF Then flxClientList.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
   Exit Sub
   
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number
   
   rstClient.Close
   conClient.Close
   Set rstClient = Nothing
   Set conClient = Nothing
End Sub

Private Sub lstResidency_DblClick()
   txtResidency.text = lstResidency.text
   lstResidency.Visible = False
End Sub

Private Sub lstResidency_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then lstResidency_DblClick
End Sub

Private Sub optLetting_Fees_Click(Index As Integer)
   txtLETTING_FEES_VALUE(Index).SetFocus
End Sub

Private Sub optManagement_Fees_Click(Index As Integer)
   txtMGT_FEES_VALUE(Index).SetFocus
End Sub

Private Sub optRecharges_Click(Index As Integer)
   iRecharge = Index
End Sub

Private Sub tabDates_Click(PreviousTab As Integer)
   Select Case tabFee.Tab
      Case 0:
         cboDay(0).SetFocus
   End Select
End Sub

Private Sub tabFee_Click(PreviousTab As Integer)
   MousePointer = vbHourglass
   
   Select Case tabFee.Tab
   Case 3:
      tabDates.Tab = 0
   End Select
   
   MousePointer = vbDefault
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
   MousePointer = vbHourglass

   Select Case tabMain.Tab
   Case 1:                    'Property
      tvwLandLord.SetFocus
   Case 2:                    'Agreement
      If cboProperty.ListCount < 1 Then
         MsgBox "No property has been entered for this client.", vbCritical + vbOKOnly, "No Property"
      Else
         cboProperty.SetFocus
      End If
   Case 3:                    'Bank Payment details
      If cboBank_ID.text = "" Or flxOtherBankDetails.TextMatrix(1, 1) = "" Then
         LoadAllBankAC
         flxOtherBankDetails.Row = 0
         flxOtherBankDetails.Col = 0
      End If
   Case 5:                    'Global setting
      tabFee.Tab = 0
      If bGlobalData Then
         cmdGSEdit.Caption = "Edit"
      Else
         cmdGSEdit.Caption = "Add New"
      End If
   Case 6:                    'Memo and File attachment
      If txtClientID.text <> "" Then _
            Call LoadAttachmentFiles(cmbFiles, txtClientID.text, "Client")
   End Select
   MousePointer = vbDefault
End Sub

Private Sub LoadTypes()
   cboLettingChrgType.Clear
   cboMngChrgType.Clear
   cboRentChrgType.Clear

   Dim sSQLQuery_ As String

   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="

   sSQLQuery_ = "SELECT * FROM ChargeTypes ORDER BY ID"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If adoMain.Recordset.RecordCount > 0 Then
 
      Dim TotalRow, TotalCol As Integer
   
      TotalRow = adoMain.Recordset.RecordCount
      TotalCol = adoMain.Recordset.Fields.Count
   
      Dim data() As String
   
      ReDim data(TotalCol, TotalRow) As String
   
      Dim i, j As Integer
      
      For i = 0 To adoMain.Recordset.RecordCount - 1
          For j = 0 To adoMain.Recordset.Fields.Count - 1
              data(j, i) = adoMain.Recordset.Fields(j)
          Next j
          adoMain.Recordset.MoveNext
      Next i
   
      cboLettingChrgType.Column() = data()
      cboMngChrgType.Column() = data()
   End If
'********************************************************************
   sSQLQuery_ = "SELECT * FROM PayableTypes ORDER BY ID"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If adoMain.Recordset.RecordCount > 0 Then
      TotalRow = adoMain.Recordset.RecordCount
      TotalCol = adoMain.Recordset.Fields.Count
   
      ReDim data(TotalCol, TotalRow) As String
   
      For i = 0 To adoMain.Recordset.RecordCount - 1
          For j = 0 To adoMain.Recordset.Fields.Count - 1
              data(j, i) = adoMain.Recordset.Fields(j)
          Next j
          adoMain.Recordset.MoveNext
      Next i
   
      cboRentChrgType.Column() = data()
   End If
End Sub

Private Sub LoadAllBankAC()
   ConfigureFlxOtherBank

   Dim conBank As New RDO.rdoConnection
   Dim rstBank As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conBank.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conBank.CursorDriver = rdUseIfNeeded
   conBank.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT tlbClientBanks.*, tlbBank.* " & _
           "FROM tlbClientBanks, tlbBank, Client " & _
           "WHERE Client.ClientID = '" & txtClientID.text & "' And " & _
               "Client.BANK_ID = tlbBank.BANK_ID And " & _
               "tlbBank.BANK_ID = tlbClientBanks.BANK_ID"
   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   If Not rstBank.EOF Then
      cboBank_ID.text = rstBank!BANK_ID
      txtBANK_NAME.text = rstBank!BANK_NAME
      txtBANK_ADDRESS1.text = rstBank!BANK_ADDRESS1
      txtBANK_ADDRESS2.text = rstBank!BANK_ADDRESS2
      txtBANK_ADDRESS3.text = rstBank!BANK_ADDRESS3
      txtBANK_POST_CODE.text = rstBank!BANK_POST_CODE
      cboPaymentMethod.text = rstBank!PaymentMethod
      txtBank_AC_Name.text = rstBank!Bank_AC_Name
      txtBANK_SC.text = rstBank!BANK_SC
      txtBANK_AC_NUM.text = rstBank!BANK_AC_NUM
      txtBacsRef.text = rstBank!BacsRef
   End If
   rstBank.Close

   szSQL = "SELECT * " & _
           "FROM tlbClientBanks, tlbBank " & _
           "WHERE CLIENT_ID = '" & txtClientID.text & "' And " & _
               "tlbBank.BANK_ID = tlbClientBanks.BANK_ID " & _
           "ORDER BY Bank_AC_Name;"

   Set rstBank = conBank.OpenResultset(szSQL, rdOpenDynamic, rdConcurRowVer)

   If rstBank.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstBank.EOF
      flxOtherBankDetails.TextMatrix(iRow, 1) = rstBank!BANK_ID
      flxOtherBankDetails.TextMatrix(iRow, 2) = rstBank!BANK_NAME
      flxOtherBankDetails.TextMatrix(iRow, 3) = rstBank!BANK_POST_CODE
      flxOtherBankDetails.TextMatrix(iRow, 4) = rstBank!Bank_AC_Name
      flxOtherBankDetails.TextMatrix(iRow, 5) = rstBank!BANK_AC_NUM
      flxOtherBankDetails.TextMatrix(iRow, 6) = rstBank!BANK_SC
      flxOtherBankDetails.TextMatrix(iRow, 7) = IIf(rstBank!DEFAULT_AC, "YES", "NO")
      flxOtherBankDetails.TextMatrix(iRow, 8) = rstBank!BANK_ADDRESS1
      flxOtherBankDetails.TextMatrix(iRow, 9) = rstBank!BANK_ADDRESS2
      flxOtherBankDetails.TextMatrix(iRow, 10) = rstBank!BANK_ADDRESS3
      flxOtherBankDetails.TextMatrix(iRow, 11) = rstBank!PaymentMethod
      flxOtherBankDetails.TextMatrix(iRow, 12) = rstBank!BacsRef

      rstBank.MoveNext
      If Not rstBank.EOF Then flxOtherBankDetails.AddItem ""
      iRow = iRow + 1
   Wend
NoRes:
   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   rstBank.Close
   conBank.Close
   Set rstBank = Nothing
   Set conBank = Nothing
End Sub

Private Sub ConfigureFlxOtherBank()
   Dim szHeader As String, i As Integer

   flxOtherBankDetails.Clear
   flxOtherBankDetails.Cols = 13

   szHeader = "|<BANK_ID|<BANK_NAME|<BANK_POST_CODE|<BANK_AC_NAME|<BANK_AC_NUM|<BANK_SC|<DEFAULT_AC"
   flxOtherBankDetails.FormatString = szHeader

   flxOtherBankDetails.ColWidth(0) = 400

   For i = 1 To flxOtherBankDetails.Cols - 1
      flxOtherBankDetails.ColWidth(i) = (flxOtherBankDetails.Width - 600) / 7
   Next i
   flxOtherBankDetails.ColWidth(8) = 0
   flxOtherBankDetails.ColWidth(9) = 0
   flxOtherBankDetails.ColWidth(10) = 0
   flxOtherBankDetails.ColWidth(11) = 0      'PaymentMethod
   flxOtherBankDetails.ColWidth(12) = 0      'BacsRef

   flxOtherBankDetails.RowHeightMin = 315
End Sub
Private Sub tvwLandLord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      szaPremisisIDType = Split(tvwLandLord.SelectedItem.key, "@")
      fraType.Caption = szaPremisisIDType(1)

      IMAGE_FILE_NAME_ = ImageLoader(imgPremises, szaPremisisIDType(0), szaPremisisIDType(1), lblImageName)

      txtTVInfoName.text = tvwLandLord.SelectedItem.text
      txtTVInfoAdd(0).text = ""
      txtTVInfoAdd(1).text = ""
      txtTVInfoAdd(2).text = ""
      txtTVInfoPC.text = ""
      
      If szaPremisisIDType(1) = "CLIENT" Then
         fraOccupied.Visible = False
      End If
      If szaPremisisIDType(1) = "PROPERTY" Then
         fraOccupied.Visible = False
         PropertyDetails szaPremisisIDType(0)
      End If
      If szaPremisisIDType(1) = "UNITS" Then
         If UnitDetails(szaPremisisIDType(0)) Then
            fraOccupied.Visible = True
         Else
            fraOccupied.Visible = False
         End If
      End If
   End If
End Sub

Private Sub LockingAcText(bLock As Boolean)
'   cmdNewBank.Enabled = Not bLock
   
   txtBank_AC_Name.Locked = bLock
   txtBANK_SC.Locked = bLock
   txtBANK_AC_NUM.Locked = bLock
   txtBacsRef.Locked = bLock
   
   If cmdNewBank.Caption = "Edit" Then Exit Sub
   
   txtBank_AC_Name.text = ""
   txtBANK_SC.text = ""
   txtBANK_AC_NUM.text = ""
   txtBacsRef.text = ""
End Sub

Private Sub LockingAllText(bLock As Boolean)
   txtClientID.Locked = bLock
   txtClientName.Locked = bLock
'   txtLandLordSageCustAC.Locked = bLock
'   txtLandLordSageSuppAC.Locked = bLock
   cmdResidency.Enabled = Not bLock
   txtAcBalance.Locked = bLock
   txtVATReg.Locked = bLock
   txtYearEndDate.Locked = bLock
   txtClientAddressLine1.Locked = bLock
   txtClientAddressLine2.Locked = bLock
   txtClientAddressLine3.Locked = bLock
   txtClientPostCode.Locked = bLock
   txtClientHomeTel.Locked = bLock
   txtClientOfficeTel.Locked = bLock
   txtClientMobile.Locked = bLock
   txtClientPersonalEmail.Locked = bLock
   txtClientOfficeEmail.Locked = bLock
   txtClientOfficeAddressLine1.Locked = bLock
   txtClientOfficeAddressLine2.Locked = bLock
   txtClientOfficeAddressLine3.Locked = bLock
   txtClientOfficePostCode.Locked = bLock
End Sub

Private Sub SageSupplierAccCombo()
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

         Dim TotalRow, TotalCol As Long
         Dim data() As String
         Dim i As Integer
             
         TotalRow = oPurchaseRecord.Count
         TotalCol = 2
         cboLandLordSageSuppAC.Clear
         
         ReDim data(TotalCol, TotalRow) As String
         
         oPurchaseRecord.MoveFirst
         For i = 0 To TotalRow - 1
            'cboTest.AddItem adoClient.Recordset.Fields(1)
            data(0, i) = CStr(oPurchaseRecord.Fields.Item("ACCOUNT_REF").Value)
            data(1, i) = CStr(oPurchaseRecord.Fields.Item("NAME").Value)
            oPurchaseRecord.MoveNext
         Next i
         '
         cboLandLordSageSuppAC.Column() = data()
         cboLandLordSageSuppAC.ColumnCount = TotalCol
         cboLandLordSageSuppAC.BoundColumn = 1
'         cboLandLordSageSuppAC.TextColumn = 2
         
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

Private Sub LoockingAgrText(bLock As Boolean)
   cmdAgmntEdit.Enabled = bLock
   cmdAgmntSave.Enabled = Not bLock
   
   fraLettingFees.Enabled = Not bLock
   fraManagementFees.Enabled = Not bLock
   fraRecharges.Enabled = Not bLock

   txtLETTING_FEES_VALUE(0).Locked = bLock
   txtLETTING_FEES_VALUE(1).Locked = bLock
   txtMGT_FEES_VALUE(0).Locked = bLock
   txtMGT_FEES_VALUE(1).Locked = bLock
   txtSTART_DATE.Locked = bLock
   txtEND_DATE.Locked = bLock
   txtREVIEW_DATE.Locked = bLock
   txtNOTICE_DAYS.Locked = bLock
   txtRRPA.Locked = bLock
   txtOWNERSHIP_PERCENT.Locked = bLock
End Sub

Private Sub txtBANK_AC_NUM_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

'Private Sub txtBANK_SC_Change()
'MsgBox "ASDFASEFD"
'End Sub
'
Private Sub txtBANK_SC_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtEND_DATE_Change()
   TextBoxChangeDate txtEND_DATE
End Sub

Private Sub txtEND_DATE_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtEND_DATE, KeyAscii
End Sub

Private Sub txtEND_DATE_LostFocus()
   TextBoxFormatDate txtEND_DATE
End Sub

Private Sub txtLettingNtDueDt_Change()
   TextBoxChangeDate txtLettingNtDueDt
End Sub

Private Sub txtLettingNtDueDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtLettingNtDueDt, KeyAscii
End Sub

Private Sub txtLettingNtDueDt_LostFocus()
   TextBoxFormatDate txtLettingNtDueDt
End Sub

Private Sub txtLettingStDt_Change()
   TextBoxChangeDate txtLettingStDt
End Sub

Private Sub txtLettingStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtLettingStDt, KeyAscii
End Sub

Private Sub txtLettingStDt_LostFocus()
   TextBoxFormatDate txtLettingStDt
End Sub

Private Sub txtMngNtDueDt_Change()
   TextBoxChangeDate txtMngNtDueDt
End Sub

Private Sub txtMngNtDueDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtMngNtDueDt, KeyAscii
End Sub

Private Sub txtMngNtDueDt_LostFocus()
   TextBoxFormatDate txtMngNtDueDt
End Sub

Private Sub txtMngStDt_Change()
   TextBoxChangeDate txtMngStDt
End Sub

Private Sub txtMngStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtMngStDt, KeyAscii
End Sub

Private Sub txtMngStDt_LostFocus()
   TextBoxFormatDate txtMngStDt
End Sub

Private Sub txtNOTICE_DAYS_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtRentNtDueDt_Change()
   TextBoxChangeDate txtRentNtDueDt
End Sub

Private Sub txtRentNtDueDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtRentNtDueDt, KeyAscii
End Sub

Private Sub txtRentNtDueDt_LostFocus()
   TextBoxFormatDate txtRentNtDueDt
End Sub

Private Sub txtRentStDt_Change()
   TextBoxChangeDate txtRentStDt
End Sub

Private Sub txtRentStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtRentStDt, KeyAscii
End Sub

Private Sub txtRentStDt_LostFocus()
   TextBoxFormatDate txtRentStDt
End Sub

Private Sub txtResidency_GotFocus()
   If cmdResidency.Enabled Then cmdResidency.SetFocus
End Sub

Private Sub txtREVIEW_DATE_Change()
   TextBoxChangeDate txtREVIEW_DATE
End Sub

Private Sub txtREVIEW_DATE_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtREVIEW_DATE, KeyAscii
End Sub

Private Sub txtREVIEW_DATE_LostFocus()
   TextBoxFormatDate txtREVIEW_DATE
End Sub

Private Sub txtRRPA_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtSTART_DATE_Change()
   TextBoxChangeDate txtSTART_DATE
End Sub

Private Sub txtSTART_DATE_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSTART_DATE, KeyAscii
End Sub

Private Sub txtSTART_DATE_LostFocus()
   TextBoxFormatDate txtSTART_DATE
End Sub

Private Sub ClearAllRecharge()
   optLetting_Fees(0).Value = False
   optLetting_Fees(1).Value = False
   txtLETTING_FEES_VALUE(0).text = ""
   txtLETTING_FEES_VALUE(1).text = ""
   optManagement_Fees(0).Value = False
   optManagement_Fees(1).Value = False
   txtMGT_FEES_VALUE(0).text = ""
   txtMGT_FEES_VALUE(1).text = ""
   optRecharges(0).Value = False
   optRecharges(1).Value = False
   optRecharges(2).Value = False
   txtSTART_DATE.text = ""
   txtEND_DATE.text = ""
   txtREVIEW_DATE.text = ""
   txtNOTICE_DAYS.text = ""
   txtRRPA.text = ""
   txtOWNERSHIP_PERCENT.text = ""
End Sub

Private Sub txtYearEndDate_Change()
   TextBoxChangeDate txtYearEndDate
End Sub

Private Sub txtYearEndDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtYearEndDate, KeyAscii
End Sub

Private Sub txtYearEndDate_LostFocus()
   TextBoxFormatDate txtYearEndDate
End Sub
