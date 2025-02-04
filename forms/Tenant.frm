VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmTenant 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kingsgate - Tenant Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tenant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleLeft       =   250
   ScaleMode       =   0  'User
   ScaleTop        =   250
   ScaleWidth      =   9180
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   240
      TabIndex        =   25
      Top             =   1440
      Width           =   7335
      Begin VB.Frame Frame1 
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
         Height          =   3495
         Index           =   3
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox txt22 
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
            Left            =   1290
            MaxLength       =   20
            TabIndex        =   63
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txt21 
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
            Left            =   1290
            MaxLength       =   20
            TabIndex        =   62
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txt20 
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
            Left            =   1290
            MaxLength       =   12
            TabIndex        =   61
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txt19 
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
            Left            =   1290
            MaxLength       =   40
            TabIndex        =   60
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txt18 
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
            Left            =   1290
            MaxLength       =   40
            TabIndex        =   59
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txt17 
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
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   58
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txt16 
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
            Left            =   1290
            MaxLength       =   40
            TabIndex        =   57
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Left            =   120
            TabIndex        =   70
            Top             =   3120
            Width           =   300
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
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
            Left            =   120
            TabIndex        =   69
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
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
            Left            =   120
            TabIndex        =   68
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 1:"
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
            Left            =   90
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 2:"
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
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 3:"
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
            Left            =   120
            TabIndex        =   65
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 4:"
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
            Left            =   120
            TabIndex        =   64
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Office Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txt11 
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
            Left            =   1305
            MaxLength       =   40
            TabIndex        =   48
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txt15 
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
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   47
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txt13 
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
            Left            =   1305
            MaxLength       =   12
            TabIndex        =   46
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txt12 
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
            Left            =   1305
            MaxLength       =   40
            TabIndex        =   45
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txt10 
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
            Left            =   1305
            MaxLength       =   40
            TabIndex        =   44
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txt9 
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
            Left            =   1305
            MaxLength       =   40
            TabIndex        =   43
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txt14 
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
            Left            =   1305
            MaxLength       =   20
            TabIndex        =   42
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Post Code:"
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
            Left            =   120
            TabIndex        =   55
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 1:"
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
            Left            =   105
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Left            =   120
            TabIndex        =   53
            Top             =   3240
            Width           =   300
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
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
            Left            =   120
            TabIndex        =   52
            Top             =   2760
            Width           =   810
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 2:"
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
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 3:"
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
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0EFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address Line 4:"
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
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2535
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox txt8 
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
            Left            =   1380
            MaxLength       =   20
            TabIndex        =   37
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txt7 
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
            Left            =   1380
            MaxLength       =   40
            TabIndex        =   36
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txt6 
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
            Left            =   1380
            MaxLength       =   40
            TabIndex        =   35
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Line:"
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
            Left            =   225
            TabIndex        =   40
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
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
            Left            =   240
            TabIndex        =   39
            Top             =   960
            Width           =   465
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
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
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Contact 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   5295
         Begin VB.TextBox txt5 
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
            Left            =   1260
            MaxLength       =   20
            TabIndex        =   30
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txt4 
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
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   29
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txt3 
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
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   28
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Line:"
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
            Left            =   225
            TabIndex        =   33
            Top             =   1440
            Width           =   810
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
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
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0EFF0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
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
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   600
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4095
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7223
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Contact 1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Contact 2"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Office Address"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Billing Address"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAddNewFile 
      Caption         =   "Attach File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   8400
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   8895
      Begin VB.TextBox txt23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1920
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "Tenant.frx":0442
         Top             =   120
         Width           =   4815
      End
      Begin VB.OptionButton optHO 
         Caption         =   "Invoice Head Office"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   2055
      End
      Begin VB.OptionButton optBil 
         Caption         =   "Invoice Billing Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6840
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label21 
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
         Left            =   6600
         TabIndex        =   22
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   8895
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Tenant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Tenant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add NewTenant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "&Save New Tenant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelNew 
         Caption         =   "&Cancel New Tenant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveEdit 
         Caption         =   "&Save Changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelEdit 
         Caption         =   "&Cancel Changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOpenFile 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox cmbFiles 
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
      ItemData        =   "Tenant.frx":044A
      Left            =   3840
      List            =   "Tenant.frx":044C
      TabIndex        =   8
      Text            =   "Files"
      Top             =   8400
      Width           =   3375
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   7680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboUnit 
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
      Left            =   5760
      TabIndex        =   6
      Text            =   "cboUnit"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txt1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.ComboBox cboSage 
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
      Left            =   1680
      TabIndex        =   5
      Text            =   "cboSage"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cboTenants 
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
      Left            =   1680
      TabIndex        =   4
      Text            =   "cboTenants"
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Select Tenant"
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
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Current Rental:"
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
      Left            =   4590
      TabIndex        =   3
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Sage A/C Number:"
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
      Left            =   225
      TabIndex        =   2
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Company Name:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1170
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMain 
         Caption         =   "Main"
      End
      Begin VB.Menu mnuShopCentre 
         Caption         =   "Shopping Centre"
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuLease 
         Caption         =   "Lease"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global Data"
      End
      Begin VB.Menu mnuDemands 
         Caption         =   "Demands"
      End
   End
   Begin VB.Menu mnuTenants 
      Caption         =   "Tenants"
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuAddNew 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmTenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TenantCode As String
Dim TenantName As String
Public OldUnit As String
Public NewUnit As String
Public OldSageAct As String
Dim Conn1 As New RDO.rdoConnection
Dim Env1 As rdoEnvironment
Dim Envs1 As rdoEnvironments
Dim Rst1 As rdoResultset
Dim Conn2 As New RDO.rdoConnection
Dim Env2 As rdoEnvironment
Dim Envs2 As rdoEnvironments
Dim Rst2 As rdoResultset
Dim SQLStr1 As String
Dim SQLStr2 As String
Private mintCurFrame As Integer ' Current Frame visible

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub cboTenants_Click()
    Dim i, j, match As Integer
    match = 0
    
    If cboTenants.Text = "" Then
        MsgBox "You must select a Tenant to view!", vbOKOnly + vbCritical, "No tenant selected"
        Exit Sub
    End If
    j = cboTenants.ListCount - 1
    For i = 0 To j
        If cboTenants.List(i) = cboTenants.Text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Tenant selected is invalid.", vbOKOnly + vbCritical, "Invalid Tenant"
        cboTenants.Text = ""
        Exit Sub
    End If
    
    ' Set the RDO Conn1ection to the dataset
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseIfNeeded
    Conn1.EstablishConnection rdDriverNoPrompt
    
    For i = 2 To 10
        If Mid(cboTenants.Text, i, 3) = " / " Then
            TenantCode = Left(cboTenants.Text, i - 1)
        End If
    Next i
        
    'Get record for selected Tenant.
    SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    'Fill text boxes with tenant details.
    cboSage.Text = TenantCode
    If IsNull(Rst1!CompanyName) Then txt1.Text = "" Else txt1.Text = Rst1!CompanyName
    If IsNull(Rst1!CurrentRental) Then cboUnit.Text = "" Else cboUnit.Text = Rst1!CurrentRental
    If IsNull(Rst1!Contact1) Then txt3.Text = "" Else txt3.Text = Rst1!Contact1
    If IsNull(Rst1!Email1) Then txt4.Text = "" Else txt4.Text = Rst1!Email1
    If IsNull(Rst1!DirectLine1) Then txt5.Text = "" Else txt5.Text = Rst1!DirectLine1
    If IsNull(Rst1!Contact2) Then txt6.Text = "" Else txt6.Text = Rst1!Contact2
    If IsNull(Rst1!Email2) Then txt7.Text = "" Else txt7.Text = Rst1!Email2
    If IsNull(Rst1!DirectLine2) Then txt8.Text = "" Else txt8.Text = Rst1!DirectLine2
    If IsNull(Rst1!HOAddressLine1) Then txt9.Text = "" Else txt9.Text = Rst1!HOAddressLine1
    If IsNull(Rst1!HOAddressLine2) Then txt10.Text = "" Else txt10.Text = Rst1!HOAddressLine2
    If IsNull(Rst1!HOAddressLine3) Then txt11.Text = "" Else txt11.Text = Rst1!HOAddressLine3
    If IsNull(Rst1!HOAddressLine4) Then txt12.Text = "" Else txt12.Text = Rst1!HOAddressLine4
    If IsNull(Rst1!HOPostCode) Then txt13.Text = "" Else txt13.Text = Rst1!HOPostCode
    If IsNull(Rst1!HOTelephone) Then txt14.Text = "" Else txt14.Text = Rst1!HOTelephone
    If IsNull(Rst1!HOFax) Then txt15.Text = "" Else txt15.Text = Rst1!HOFax
    If IsNull(Rst1!BillAddressLine1) Then txt16.Text = "" Else txt16.Text = Rst1!BillAddressLine1
    If IsNull(Rst1!BillAddressLine2) Then txt17.Text = "" Else txt17.Text = Rst1!BillAddressLine2
    If IsNull(Rst1!BillAddressLine3) Then txt18.Text = "" Else txt18.Text = Rst1!BillAddressLine3
    If IsNull(Rst1!BillAddressLine4) Then txt19.Text = "" Else txt19.Text = Rst1!BillAddressLine4
    If IsNull(Rst1!BillPostCode) Then txt20.Text = "" Else txt20.Text = Rst1!BillPostCode
    If IsNull(Rst1!BillTelephone) Then txt21.Text = "" Else txt21.Text = Rst1!BillTelephone
    If IsNull(Rst1!BillFax) Then txt22.Text = "" Else txt22.Text = Rst1!BillFax
    If IsNull(Rst1!Comments) Then txt23.Text = "" Else txt23.Text = Rst1!Comments
    If Rst1!InvoiceTo = "B" Then optBil.Value = True Else optHO.Value = True
    
    Rst1.Close
    Conn1.Close
    
    Call DisableTextBoxes

    LodeAttachedFiles
End Sub

Private Sub LodeAttachedFiles()
    cmbFiles.Clear
    
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseOdbc
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT FileName FROM AttachedFile where SageAccountNumber = '" & cboSage.Text & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

    While Not Rst1.EOF
        cmbFiles.AddItem Rst1!FileName
        Rst1.MoveNext
    Wend
    Rst1.Close
    Conn1.Close
End Sub

Private Sub cmdAdd_Click()

    Call AddNew

End Sub

Private Sub cmdAddNewFile_Click()
    Dim ofn As OPENFILENAME, retVal As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = CurDir
    ofn.lpstrTitle = "Our File Open Title"
    ofn.flags = 0

    retVal = GetOpenFileName(ofn)

    If (retVal) Then
        Dim TxtFile As String, ff As Integer
        TxtFile = "C:\Filelist.txt"
        ff = FreeFile
'        Open TxtFile For Append As #ff
'        Print #ff, Trim$(ofn.lpstrFile)
        Close #ff
    End If

    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseOdbc
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT * FROM AttachedFile"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

    Rst1.AddNew
    Rst1!filePath = Trim$(ofn.lpstrFile)
    Rst1!FileName = Trim$(ofn.lpstrFileTitle)
    Rst1!SageAccountNumber = cboSage.Text
    Rst1!Tenant = cboTenants.Text

    Rst1.Update
    Rst1.Close
    Conn1.Close
    
    cmbFiles.AddItem Trim$(ofn.lpstrFileTitle)

    MsgBox "File has been attached successfull, Thanks"
End Sub

Private Sub cmdAttachmentDoc_Click()
'    MsgBox Me.Height
'    If Me.Height = 8040 Then
'        Me.Height = 8520
'    Else
'        Me.Height = 8040
'    End If
    
End Sub

Private Sub cmdCancelEdit_Click()

Dim i As Integer

' Set the RDO Connection to the dataset
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseIfNeeded
Conn1.EstablishConnection rdDriverNoPrompt

For i = 2 To 9
    If Mid(cboTenants.Text, i, 3) = " / " Then
        TenantCode = Left(cboTenants.Text, i - 1)
    End If
Next i

'Get record for selected Tenant.
SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

'Fill text boxes with tenant details.
cboSage.Text = TenantCode
If IsNull(Rst1!CompanyName) Then txt1.Text = "" Else txt1.Text = Rst1!CompanyName
If IsNull(Rst1!CurrentRental) Then cboUnit.Text = "" Else cboUnit.Text = Rst1!CurrentRental
If IsNull(Rst1!Contact1) Then txt3.Text = "" Else txt3.Text = Rst1!Contact1
If IsNull(Rst1!Email1) Then txt4.Text = "" Else txt4.Text = Rst1!Email1
If IsNull(Rst1!DirectLine1) Then txt5.Text = "" Else txt5.Text = Rst1!DirectLine1
If IsNull(Rst1!Contact2) Then txt6.Text = "" Else txt6.Text = Rst1!Contact2
If IsNull(Rst1!Email2) Then txt7.Text = "" Else txt7.Text = Rst1!Email2
If IsNull(Rst1!DirectLine2) Then txt8.Text = "" Else txt8.Text = Rst1!DirectLine2
If IsNull(Rst1!HOAddressLine1) Then txt9.Text = "" Else txt9.Text = Rst1!HOAddressLine1
If IsNull(Rst1!HOAddressLine2) Then txt10.Text = "" Else txt10.Text = Rst1!HOAddressLine2
If IsNull(Rst1!HOAddressLine3) Then txt11.Text = "" Else txt11.Text = Rst1!HOAddressLine3
If IsNull(Rst1!HOAddressLine4) Then txt12.Text = "" Else txt12.Text = Rst1!HOAddressLine4
If IsNull(Rst1!HOPostCode) Then txt13.Text = "" Else txt13.Text = Rst1!HOPostCode
If IsNull(Rst1!HOTelephone) Then txt14.Text = "" Else txt14.Text = Rst1!HOTelephone
If IsNull(Rst1!HOFax) Then txt15.Text = "" Else txt15.Text = Rst1!HOFax
If IsNull(Rst1!BillAddressLine1) Then txt16.Text = "" Else txt16.Text = Rst1!BillAddressLine1
If IsNull(Rst1!BillAddressLine2) Then txt17.Text = "" Else txt17.Text = Rst1!BillAddressLine2
If IsNull(Rst1!BillAddressLine3) Then txt18.Text = "" Else txt18.Text = Rst1!BillAddressLine3
If IsNull(Rst1!BillAddressLine4) Then txt19.Text = "" Else txt19.Text = Rst1!BillAddressLine4
If IsNull(Rst1!BillPostCode) Then txt20.Text = "" Else txt20.Text = Rst1!BillPostCode
If IsNull(Rst1!BillTelephone) Then txt21.Text = "" Else txt21.Text = Rst1!BillTelephone
If IsNull(Rst1!BillFax) Then txt22.Text = "" Else txt22.Text = Rst1!BillFax
If IsNull(Rst1!Comments) Then txt23.Text = "" Else txt23.Text = Rst1!Comments
If Rst1!InvoiceTo = "B" Then optBil.Value = True Else optHO.Value = True

Rst1.Close
Conn1.Close

Call CancelAddEdit
Call DisableTextBoxes

End Sub

Private Sub cmdCancelNew_Click()

Call CancelAddEdit
Call ResetScreen
Call DisableTextBoxes
Call EmptyBoxes

End Sub

Private Sub cmdDelete_Click()

Call Delete

End Sub

Private Sub cmdEdit_Click()

Call Edit

End Sub

Private Sub cmdOpenFile_Click()
    If cmbFiles.Text <> "" Then
        Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
        Conn1.CursorDriver = rdUseOdbc
        Conn1.EstablishConnection rdDriverNoPrompt
        
        SQLStr1 = "SELECT FilePath FROM AttachedFile where SageAccountNumber = '" & cboSage.Text & "' And FileName = '" & cmbFiles.Text & "'"
        Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        If Rst1.RowCount = 1 Then
            Shell Rst1!filePath
        End If
        Rst1.Close
        Conn1.Close
    End If


'While Not
'    cmbFiles.AddItem Rst1!FileName
'    Rst1.MoveNext
'Wend

End Sub

Private Sub cmdPrint_Click()

MousePointer = vbHourglass

If FileExists(App.Path & "\Tenant" & SCID & ".rpt") = False Then
    MsgBox "Unable to Print Details", vbOKOnly + vbInformation, "Unable To Print"
    Exit Sub
End If

CR1.ReportFileName = App.Path & "\Tenant" & SCID & ".rpt"
CR1.SelectionFormula = "{Tenants.SageAccountNumber} = '" & cboSage.Text & "'"
CR1.PrintReport

MousePointer = vbDefault

End Sub

Private Sub cmdSaveEdit_Click()

Dim i As Integer
Dim j As Integer
Dim match As Integer
match = 0

If cboSage.Text = "" Then
    MsgBox "You must select a Sage Account Number!", vbOKOnly + vbCritical, "Sage Account Number required."
    Exit Sub
Else
    j = cboSage.ListCount
    For i = 0 To j - 1
        If cboSage.List(i) = cboSage.Text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Sage Account Number selected is invalid.", vbOKOnly + vbCritical, "Invalid Sage Account Number"
        cboSage.Text = OldSageAct
        Exit Sub
    End If
End If

match = 0

If cboUnit.Text <> "" Then
    j = cboUnit.ListCount - 1
    For i = 0 To j
        If cboUnit.List(i) = cboUnit.Text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Current Rental selected is not a valid unit number.", vbOKOnly + vbCritical, "Invalid Current Rental"
        cboUnit.Text = OldUnit
        Exit Sub
    End If
End If

For i = 10 To 2
    If Mid(cboTenants.Text, i, 3) = " / " Then
        TenantCode = Left(cboTenants.Text, i - 1)
        TenantName = Right(cboTenants.Text, Len(cboTenants.Text) - i + 1)
    End If
Next i

Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseIfNeeded
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

Rst1.Edit

Rst1!CompanyName = txt1.Text
Rst1!SageAccountNumber = cboSage.Text
Rst1!Contact1 = txt3.Text
Rst1!Email1 = txt4.Text
Rst1!DirectLine1 = txt5.Text
Rst1!Contact2 = txt6.Text
Rst1!Email2 = txt7.Text
Rst1!DirectLine2 = txt8.Text
Rst1!HOAddressLine1 = txt9.Text
Rst1!HOAddressLine2 = txt10.Text
Rst1!HOAddressLine3 = txt11.Text
Rst1!HOAddressLine4 = txt12.Text
Rst1!HOPostCode = txt13.Text
Rst1!HOTelephone = txt14.Text
Rst1!HOFax = txt15.Text
Rst1!BillAddressLine1 = txt16.Text
Rst1!BillAddressLine2 = txt17.Text
Rst1!BillAddressLine3 = txt18.Text
Rst1!BillAddressLine4 = txt19.Text
Rst1!BillPostCode = txt20.Text
Rst1!BillTelephone = txt21.Text
Rst1!BillFax = txt22.Text
Rst1!CurrentRental = cboUnit.Text
Rst1!Comments = txt23.Text
If optBil.Value = True Then Rst1!InvoiceTo = "B" Else Rst1!InvoiceTo = "H"

Rst1.Update
Rst1.Close
Conn1.Close

NewUnit = cboUnit.Text

Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn2.CursorDriver = rdUseIfNeeded
Conn2.EstablishConnection rdDriverNoPrompt

If NewUnit <> OldUnit Then
    If NewUnit = "" Then
        SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & OldUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)

        Rst2.Edit
        Rst2!Occupied = "N"
        Rst2!TenantCompanyName = ""
        Rst2!SageAccountNumber = ""
        Rst2.Update
   End If

   If OldUnit = "" Then
        SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & NewUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
        
        Rst2.Edit
        Rst2!Occupied = "Y"
        Rst2!TenantCompanyName = txt1.Text
        Rst2!SageAccountNumber = cboSage.Text
        Rst2.Update
   End If

   If OldUnit <> "" And NewUnit <> "" Then
        SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & OldUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
        
        Rst2.Edit
        Rst2!Occupied = "N"
        Rst2!TenantCompanyName = ""
        Rst2!SageAccountNumber = ""
        Rst2.Update
        Rst2.Close

        SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & NewUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
    
        Rst2.Edit
        Rst2!Occupied = "Y"
        Rst2!TenantCompanyName = txt1.Text
        Rst2!SageAccountNumber = cboSage.Text
        Rst2.Update
        Rst2.Close

        SQLStr2 = "SELECT * FROM LeaseDetails WHERE UnitNumber = '" & OldUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)

        Rst2.Edit
        Rst2!UnitNumber = NewUnit
        Rst2.Update
        
   End If

   Rst2.Close
Else
    If txt1.Text <> TenantName Then
        SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & NewUnit & "'"
        Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)
    
        Rst2.Edit
        Rst2!TenantCompanyName = txt1.Text
        Rst2.Update
        Rst2.Close
    End If
End If

If txt1.Text <> TenantName Or cboSage.Text <> TenantCode Then 'amend on all demands that have not been exported to Sage.
    SQLStr1 = "SELECT * FROM LeaseDetails WHERE SageAccountNumber = '" & TenantCode & "'"
    Set Rst2 = Conn2.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
    If Rst2.EOF = False Then
        Rst2.Edit
        Rst2!CompanyName = txt1.Text
        Rst2!SageAccountNumber = cboSage.Text
        Rst2.Update
    End If
    Rst2.Close
    SQLStr1 = "SELECT * FROM DemandRecords WHERE ExportedToSage <> 'Y' AND SageAccountNumber = '" & TenantCode & "'"
    Set Rst2 = Conn2.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
    If Rst2.EOF = False Then
        While Rst2.EOF = False
            Rst2.Edit
            Rst2!TenantCompanyName = txt1.Text
            Rst2!SageAccountNumber = cboSage.Text
            Rst2.Update
            Rst2.MoveNext
        Wend
    End If
    Rst2.Close
End If
Conn2.Close

cboTenants.Text = cboSage.Text & " / " & txt1.Text
    
MsgBox "Your changes have been saved.", vbOKOnly + vbInformation, "Saved"

cmdAdd.Visible = True
cmdDelete.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdEdit.Visible = True
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False
mnuAddNew.Enabled = True
mnuDel.Enabled = True
mnuEdit.Enabled = True
cmdPrint.Visible = True

NewUnit = ""
OldUnit = ""

Call ResetScreen
Call DisableTextBoxes

End Sub

Private Sub cmdSaveNew_Click()
    
Dim i, j, match As Integer
match = 0
    
If cboSage.Text = "" Then
    MsgBox "You must select a Sage Account Number!", vbOKOnly + vbCritical, "Sage Account Number required."
    Exit Sub
Else
    j = cboSage.ListCount - 1
    For i = 0 To j
        If cboSage.List(i) = cboSage.Text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Sage Account Number selected is invalid.", vbOKOnly + vbCritical, "Sage Account Number is invalid"
        cboSage.Text = ""
        Exit Sub
    End If
End If

match = 0

If cboUnit.Text <> "" Then
    j = cboUnit.ListCount - 1
    For i = 0 To j
        If cboUnit.List(i) = cboUnit.Text Then
            match = 1
            Exit For
        End If
    Next i
    If match = 0 Then
        MsgBox "Current Rental selected is not a valid unit number.", vbOKOnly + vbCritical, "Invalid Current Rental"
        cboUnit.Text = ""
        Exit Sub
    End If
End If

If txt1.Text = "" Then
    MsgBox "You must enter a Tenant Company Name!", vbCritical + vbOKOnly, "Missing Company Name"
    Exit Sub
End If

Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseOdbc
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM Tenants"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

Rst1.AddNew

Rst1!SageAccountNumber = cboSage.Text
Rst1!CompanyName = txt1.Text
Rst1!Contact1 = txt3.Text
Rst1!Email1 = txt4.Text
Rst1!DirectLine1 = txt5.Text
Rst1!Contact2 = txt6.Text
Rst1!Email2 = txt7.Text
Rst1!DirectLine2 = txt8.Text
Rst1!HOAddressLine1 = txt9.Text
Rst1!HOAddressLine2 = txt10.Text
Rst1!HOAddressLine3 = txt11.Text
Rst1!HOAddressLine4 = txt12.Text
Rst1!HOPostCode = txt13.Text
Rst1!HOTelephone = txt14.Text
Rst1!HOFax = txt15.Text
Rst1!BillAddressLine1 = txt16.Text
Rst1!BillAddressLine2 = txt17.Text
Rst1!BillAddressLine3 = txt18.Text
Rst1!BillAddressLine4 = txt19.Text
Rst1!BillPostCode = txt20.Text
Rst1!BillTelephone = txt21.Text
Rst1!BillFax = txt22.Text
Rst1!CurrentRental = cboUnit.Text
Rst1!Comments = txt23.Text
If optBil.Value = True Then Rst1!InvoiceTo = "B" Else Rst1!InvoiceTo = "H"

Rst1.Update
Rst1.Close
Conn1.Close

If cboUnit.Text <> "" Then
    
    Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn2.CursorDriver = rdUseOdbc
    Conn2.EstablishConnection rdDriverNoPrompt
    
    SQLStr2 = "SELECT * FROM Units WHERE UnitNumber = '" & cboUnit.Text & "'"
    Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenDynamic, rdConcurRowVer)

    Rst2.Edit

    Rst2!Occupied = "Y"
    Rst2!TenantCompanyName = txt1.Text
    Rst2!SageAccountNumber = cboSage.Text
    Rst2.Update
           
    Rst2.Close
    Conn2.Close
End If

cboTenants.Text = cboSage.Text & " / " & txt1.Text
    
MsgBox "The new tenant details have been saved.", vbOKOnly + vbInformation, "New Tenant"
    
Call ResetScreen
Call DisableTextBoxes

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorTrap
    
    Me.Caption = gCurrentShopCentreName & " - Tenants"
    'Me.Move (Screen.Width - Width) / 2, 0
    
    cboSage.Clear
    cboUnit.Text = ""
    cboTenants.Text = ""
    
    Call ResetScreen
    Call DisableTextBoxes
    
ErrorTrap:
        If Err.Number <> 0 Then
            If Err.Number = 40002 Then
                If MsgBox("DSN - " & Adsn & " does not exist.", vbRetryCancel, "DSN Error") = vbCancel Then
                    Resume Next
                Else
                    Resume
                End If
            Else
                MsgBox Err.Number & " - " & Err.Description
            End If
        End If
        Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMMain.fraCmdButton.Enabled = True
    Unload Me
End Sub

Private Sub mnuAddNew_Click()
    Call AddNew
End Sub

Private Sub mnuDel_Click()
    Call Delete
End Sub

Private Sub mnuDemands_Click()
    Unload Me
    Load frmDemands
    frmDemands.Show
End Sub

Private Sub mnuEdit_Click()

    Call Edit

End Sub

Private Sub mnuExit_Click()

    Call ExitProgram

End Sub

Public Sub EnableTextBoxes()

cboTenants.Enabled = False
cboSage.Enabled = True
txt1.Enabled = True
cboUnit.Enabled = True
txt3.Enabled = True
txt4.Enabled = True
txt5.Enabled = True
txt6.Enabled = True
txt7.Enabled = True
txt8.Enabled = True
txt9.Enabled = True
txt10.Enabled = True
txt11.Enabled = True
txt12.Enabled = True
txt13.Enabled = True
txt14.Enabled = True
txt15.Enabled = True
txt16.Enabled = True
txt17.Enabled = True
txt18.Enabled = True
txt19.Enabled = True
txt20.Enabled = True
txt21.Enabled = True
txt22.Enabled = True
txt23.Enabled = True
optHO.Enabled = True
optBil.Enabled = True

End Sub

Public Sub DisableTextBoxes()

cboTenants.Enabled = True
cboSage.Enabled = False
txt1.Enabled = False
cboUnit.Enabled = False
txt3.Enabled = False
txt4.Enabled = False
txt5.Enabled = False
txt6.Enabled = False
txt7.Enabled = False
txt8.Enabled = False
txt9.Enabled = False
txt10.Enabled = False
txt11.Enabled = False
txt12.Enabled = False
txt13.Enabled = False
txt14.Enabled = False
txt15.Enabled = False
txt16.Enabled = False
txt17.Enabled = False
txt18.Enabled = False
txt19.Enabled = False
txt20.Enabled = False
txt21.Enabled = False
txt22.Enabled = False
txt23.Enabled = False
optBil.Enabled = False
optHO.Enabled = False

End Sub


Public Sub GetSageActUnits()

Dim a As Integer
Dim j As Integer
Dim b As Integer
Dim k As Integer
Dim match As Integer
Dim temp1 As String
Dim temp2 As String
Dim chk1 As Boolean
chk1 = False

On Error GoTo ErrTrap1

If cboSage.Text = "" Then
    cboSage.Clear
    temp1 = ""
Else
    temp1 = cboSage.Text
    cboSage.Clear
End If

'Get all the Sage Account Numbers from Sage and put in an array.
' Set the RDO Env1ironment
Set Envs1 = rdoEngine.rdoEnvironments
Set Env1 = Envs1(0)
  
' Use Line100's scrolling cursor
Env1.CursorDriver = rdUseServer

' Set the RDO Conn1ection to the dataset
Set Conn1 = Env1.OpenConnection("", rdDriverNoPrompt, False, "DSN=" & Sdsn & ";UID=PCM;PWD=PCM")

'SQLStr1 = "SELECT ACCOUNT_NUMBER FROM SALES_LEDGER ORDER BY ACCOUNT_NUMBER"
SQLStr1 = "SELECT ACCOUNT_REF FROM SALES_LEDGER ORDER BY ACCOUNT_REF"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

'Get all SageAccountNumbers from Tenants table and put in array.
Conn2.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn2.CursorDriver = rdUseIfNeeded
Conn2.EstablishConnection rdDriverNoPrompt

SQLStr2 = "SELECT SageAccountNumber FROM Tenants ORDER BY SageAccountNumber"
Set Rst2 = Conn2.OpenResultset(SQLStr2, rdOpenStatic, rdConcurReadOnly)

If Rst1.EOF = False Then
    If Rst2.EOF = True Then 'add all sage accounts
        While Rst1.EOF = False
            cboSage.AddItem Rst1!ACCOUNT_REF
            Rst1.MoveNext
        Wend
    Else
        Rst1.MoveFirst
        While Rst1.EOF = False
            chk1 = False
            Rst2.MoveFirst
            While Rst2.EOF = False
                If Rst2!SageAccountNumber = Rst1!ACCOUNT_REF Then chk1 = True
                Rst2.MoveNext
            Wend
            If chk1 = False Then cboSage.AddItem Rst1!ACCOUNT_REF
            Rst1.MoveNext
        Wend
    End If
End If

Rst2.Close
Conn2.Close
    
Rst1.Close
Conn1.Close
Env1.Close

If temp1 <> "" Then
    cboSage.Text = temp1
    cboSage.AddItem temp1, 0
End If
If cboUnit.Text = "" Then
    cboUnit.Clear
Else
    temp2 = cboUnit.Text
    cboUnit.Clear
End If

' Get all the unoccupied unit numbers and put in cboUnit.
' Set the RDO Conn1ection to the dataset
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseIfNeeded
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT UnitNumber FROM Units WHERE Occupied = 'N' ORDER BY UnitNumber"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            cboUnit.AddItem Rst1!UnitNumber
            Rst1.MoveNext
        Wend
    End If

Rst1.Close
Conn1.Close

If temp2 <> "" Then
    cboUnit.Text = temp2
    cboUnit.AddItem temp2, 0
End If

Exit Sub

ErrTrap1:
    If Err.Number = 40002 Then
        If MsgBox("DSN - " & Sdsn & " not found. Please check with your system adminstrator.", vbRetryCancel + vbCritical, "DSN Set Up Error") = vbRetry Then
            Resume
        Else
            Resume Next
        End If
    Else
        If Err.Number = 40009 Then Resume Next
        If Err.Number <> 0 Then
            MsgBox Err.Number & " - " & Err.Description
            Resume Next
        End If
    End If
    
End Sub

Public Sub ResetScreen()

Dim temp

cboTenants.Enabled = True

If cboTenants.Text = "" Then
    cboTenants.Clear
Else
    temp = cboTenants.Text
    cboTenants.Clear
    cboTenants.Text = temp
End If

'Reset screen to show all Company Names and Sage Account numer in cboTenants
' Set the RDO Conn1ection to the dataset
Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn1.CursorDriver = rdUseIfNeeded
Conn1.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT SageAccountNumber, CompanyName FROM Tenants ORDER BY SageAccountNumber"
Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            cboTenants.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
            Rst1.MoveNext
        Wend
    End If

Rst1.Close
Conn1.Close

cmdAdd.Visible = True
cmdDelete.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdEdit.Visible = True
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False
mnuEdit.Enabled = True
mnuAddNew.Enabled = True
mnuDel.Enabled = True
cmdPrint.Visible = True

End Sub

Public Sub EmptyBoxes()

txt1.Text = ""
cboUnit.Text = ""
cboSage.Text = ""
cboTenants.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
txt6.Text = ""
txt7.Text = ""
txt8.Text = ""
txt9.Text = ""
txt10.Text = ""
txt11.Text = ""
txt12.Text = ""
txt13.Text = ""
txt14.Text = ""
txt15.Text = ""
txt16.Text = ""
txt17.Text = ""
txt18.Text = ""
txt19.Text = ""
txt20.Text = ""
txt21.Text = ""
txt22.Text = ""
txt23.Text = ""
optBil.Value = True

End Sub

Public Sub CancelAddEdit()

cmdAdd.Visible = True
cmdDelete.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdEdit.Visible = True
cmdSaveEdit.Visible = False
cmdCancelEdit.Visible = False
mnuEdit.Enabled = True
mnuAddNew.Enabled = True
mnuDel.Enabled = True
cmdPrint.Visible = True

End Sub

Public Sub AddNew()

    cboTenants.Enabled = False
    
    TenantCode = ""
    
    cmdAdd.Visible = False
    cmdDelete.Visible = False
    cmdSaveNew.Visible = True
    cmdCancelNew.Visible = True
    cmdEdit.Visible = False
    cmdSaveEdit.Visible = False
    cmdCancelEdit.Visible = False
    mnuAddNew.Enabled = False
    mnuDel.Enabled = False
    mnuEdit.Enabled = False
    cmdPrint.Visible = False
    
    Call EnableTextBoxes
    Call EmptyBoxes
    Call GetSageActUnits
    
    cboUnit.Text = ""
    OldSageAct = ""
    OldUnit = ""

End Sub

Public Sub Edit()
    If cboTenants.Text = "" Then
        MsgBox "You must select a tenant to Edit!", vbOKOnly + vbCritical, "No tenant selected"
    Else
        OldUnit = cboUnit.Text
        OldSageAct = cboSage.Text
        Call EnableTextBoxes
        
        cmdAdd.Visible = False
        cmdDelete.Visible = False
        cmdSaveNew.Visible = False
        cmdCancelNew.Visible = False
        cmdEdit.Visible = False
        cmdSaveEdit.Visible = True
        cmdCancelEdit.Visible = True
        mnuEdit.Enabled = False
        mnuAddNew.Enabled = False
        mnuDel.Enabled = False
        cmdPrint.Visible = False
        
        Call GetSageActUnits
    End If
End Sub

Public Sub Delete()

Dim Response

If cboTenants.Text = "" Then
    MsgBox "You must select a tenant to delete!", vbOKOnly + vbCritical, "No tenant selected"
Else
   
    Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn1.CursorDriver = rdUseOdbc
    Conn1.EstablishConnection rdDriverNoPrompt

    SQLStr1 = "SELECT SageAccountNumber FROM LeaseDetails WHERE SageAccountNumber = '" & cboSage.Text & "'"
    Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
        
    If Rst1.EOF = False Then
        MsgBox "Unable to delete tenant: " & txt1.Text & " - a lease record exists for tenant.", vbOKOnly + vbCritical, "Delete Tenant"
        Rst1.Close
        Conn1.Close
        Exit Sub
    Else
        Rst1.Close
        If cboUnit.Text <> "" Then
            MsgBox "Can not delete tenant - tenant is occupying a unit!", vbOKOnly + vbCritical, "Delete Tenant"
            Conn1.Close
            Exit Sub
        End If
        
        Response = MsgBox("Are you sure you want to delete tenant: " & txt1.Text & "?", vbYesNo + vbQuestion, "Delete tenant")
    
        If Response = vbYes Then
            
            SQLStr1 = "SELECT * FROM Tenants WHERE SageAccountNumber = '" & cboSage.Text & "'"
            Set Rst1 = Conn1.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)
                    
            Rst1.Delete
            Rst1.Close
            Conn1.Close
            
            Call EmptyBoxes
            Call ResetScreen
        End If
    End If
    
End If

End Sub

Private Sub mnuGlobal_Click()

Load frmGlobal
Unload Me
frmGlobal.Show

End Sub

Private Sub mnuLease_Click()

Load frmLease
Unload Me
frmLease.Show

End Sub

Private Sub mnuMain_Click()

'Load frmMain
Unload Me
'frmMain.Show
frmMMain.fraCmdButton.Enabled = True
End Sub

Private Sub mnuShopCentre_Click()

Load frmShoppingCentre
Unload Me
frmShoppingCentre.Show

End Sub

Private Sub mnuUnits_Click()

Load frmUnit
Unload Me
frmUnit.Show

End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index - 1 = mintCurFrame Then Exit Sub     ' No need to change frame.
   ' Otherwise, hide old frame, show new.
   Frame1(TabStrip1.SelectedItem.Index - 1).Visible = True
   Frame1(mintCurFrame).Visible = False
   ' Set mintCurFrame to new value.
   mintCurFrame = TabStrip1.SelectedItem.Index - 1
End Sub
