VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fund"
   ClientHeight    =   11595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16815
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFund.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   16815
   Begin VB.Frame fraWizardMaster 
      BackColor       =   &H8000000B&
      Height          =   10995
      Left            =   4680
      TabIndex        =   26
      Top             =   315
      Visible         =   0   'False
      Width           =   16215
      Begin VB.CommandButton cmdCloseFrame 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   14535
         TabIndex        =   56
         Top             =   10485
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnassign 
         Caption         =   "UnAssign Fund"
         Height          =   375
         Left            =   7065
         TabIndex        =   53
         Top             =   10485
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3390
         Index           =   3
         Left            =   11160
         TabIndex        =   30
         Top             =   1125
         Width           =   4935
         Begin VB.CheckBox chkProperties 
            Caption         =   "Select Property:"
            Height          =   195
            Left            =   135
            TabIndex        =   37
            Top             =   90
            Width           =   3840
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperty 
            Height          =   2715
            Left            =   45
            TabIndex        =   10
            Top             =   630
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   4789
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Client ID"
            Height          =   195
            Left            =   3825
            TabIndex        =   55
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Property Name"
            Height          =   195
            Left            =   1710
            TabIndex        =   46
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Property ID"
            Height          =   195
            Left            =   405
            TabIndex        =   45
            Top             =   360
            Width           =   810
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0FFFF&
            FillStyle       =   0  'Solid
            Height          =   285
            Index           =   3
            Left            =   45
            Top             =   315
            Width           =   4770
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5775
         Index           =   4
         Left            =   135
         TabIndex        =   34
         Top             =   4590
         Width           =   15735
         Begin VB.CheckBox chkSel 
            Caption         =   "Check1"
            Height          =   195
            Left            =   135
            TabIndex        =   51
            Top             =   540
            Width           =   240
         End
         Begin VB.CommandButton cmdPreviewAssignment 
            Caption         =   "&Preview Assignment"
            Height          =   375
            Left            =   135
            TabIndex        =   40
            Top             =   90
            Width           =   3150
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPreviewAssignment 
            Height          =   4830
            Left            =   135
            TabIndex        =   35
            Top             =   855
            Width           =   15540
            _ExtentX        =   27411
            _ExtentY        =   8520
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Fund Category"
            Height          =   195
            Index           =   15
            Left            =   9315
            TabIndex        =   50
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Fund Name"
            Height          =   195
            Index           =   14
            Left            =   6705
            TabIndex        =   49
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Fund Code"
            Height          =   195
            Index           =   13
            Left            =   4815
            TabIndex        =   48
            Top             =   540
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Property ID"
            Height          =   195
            Index           =   12
            Left            =   2610
            TabIndex        =   47
            Top             =   540
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Client ID"
            Height          =   195
            Index           =   11
            Left            =   855
            TabIndex        =   36
            Top             =   540
            Width           =   630
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0FFFF&
            FillStyle       =   0  'Solid
            Height          =   285
            Index           =   0
            Left            =   135
            Top             =   495
            Width           =   15525
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3390
         Index           =   1
         Left            =   135
         TabIndex        =   32
         Top             =   1125
         Width           =   6015
         Begin VB.CheckBox chkFund 
            Caption         =   "Please Select a Fund:"
            Height          =   195
            Left            =   135
            TabIndex        =   39
            Top             =   90
            Width           =   2490
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFundWizard 
            Height          =   2715
            Left            =   135
            TabIndex        =   33
            Top             =   630
            Width           =   5820
            _ExtentX        =   10266
            _ExtentY        =   4789
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Fund  Category"
            Height          =   195
            Left            =   4005
            TabIndex        =   54
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "FundName"
            Height          =   195
            Left            =   2295
            TabIndex        =   42
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "FundCode"
            Height          =   195
            Left            =   630
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0FFFF&
            FillStyle       =   0  'Solid
            Height          =   285
            Index           =   1
            Left            =   135
            Top             =   315
            Width           =   5805
         End
      End
      Begin VB.CommandButton cmdGridClose 
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
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   225
         Width           =   255
      End
      Begin VB.CommandButton cmdCancelWizard 
         Caption         =   "&Back"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   10485
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveAssignment 
         Caption         =   "&Save Assignment"
         Height          =   375
         Left            =   8505
         TabIndex        =   9
         Top             =   10485
         Width           =   1980
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3390
         Index           =   2
         Left            =   6210
         TabIndex        =   29
         Top             =   1125
         Width           =   4980
         Begin VB.CheckBox chkClient 
            Caption         =   "Please Select Client:"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   90
            Width           =   2130
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
            Height          =   2715
            Left            =   45
            TabIndex        =   7
            Top             =   630
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   4789
            _Version        =   393216
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Client Name"
            Height          =   195
            Left            =   1620
            TabIndex        =   44
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFFF&
            Caption         =   "Client ID"
            Height          =   195
            Left            =   450
            TabIndex        =   43
            Top             =   360
            Width           =   630
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00E0FFFF&
            FillStyle       =   0  'Solid
            Height          =   285
            Index           =   2
            Left            =   45
            Top             =   315
            Width           =   4905
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   90
         ScaleHeight     =   825
         ScaleWidth      =   15660
         TabIndex        =   27
         Top             =   225
         Width           =   15690
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to the Fund Assignment Wizard"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   270
            Left            =   5175
            TabIndex        =   28
            Top             =   270
            Width           =   4485
         End
      End
      Begin VB.Shape Shape1 
         Height          =   10770
         Left            =   90
         Top             =   180
         Width           =   16035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   9090
      Width           =   16665
      Begin VB.CommandButton cmdViewAssignment 
         Caption         =   "View Fund assignment"
         Height          =   375
         Left            =   7380
         TabIndex        =   52
         Top             =   1530
         Width           =   1980
      End
      Begin VB.CommandButton cmdAssignFund 
         Caption         =   "Assign Fund"
         Height          =   375
         Left            =   6210
         TabIndex        =   5
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add &New"
         Height          =   375
         Left            =   495
         TabIndex        =   0
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1665
         TabIndex        =   1
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3945
         TabIndex        =   3
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cl&ose"
         Height          =   375
         Left            =   14760
         TabIndex        =   6
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2790
         TabIndex        =   2
         Top             =   1530
         Width           =   1120
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5025
         TabIndex        =   4
         Top             =   1530
         Width           =   1120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   765
         Width           =   435
      End
      Begin MSForms.TextBox txtName 
         Height          =   315
         Left            =   1845
         TabIndex        =   12
         Top             =   765
         Width           =   10545
         VariousPropertyBits=   746604571
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "18600;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDemandTypeCategory 
         Height          =   315
         Left            =   1845
         TabIndex        =   13
         Top             =   1125
         Width           =   10545
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "18600;556"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;5000"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFDFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   720
         TabIndex        =   18
         Top             =   1080
         Width           =   750
      End
      Begin MSForms.TextBox txtCode 
         Height          =   315
         Left            =   1845
         TabIndex        =   11
         Top             =   405
         Width           =   10545
         VariousPropertyBits=   746604571
         MaxLength       =   100
         BorderStyle     =   1
         Size            =   "18600;556"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   17
         Top             =   405
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8925
      Left            =   45
      TabIndex        =   16
      Top             =   90
      Width           =   16710
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxFund 
         Height          =   8370
         Left            =   90
         TabIndex        =   20
         Top             =   450
         Width           =   16125
         _ExtentX        =   28443
         _ExtentY        =   14764
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
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
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "No."
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Code"
         Height          =   195
         Index           =   2
         Left            =   585
         TabIndex        =   23
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Category"
         Height          =   195
         Index           =   4
         Left            =   5490
         TabIndex        =   22
         Top             =   225
         Width           =   10815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Name"
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   21
         Top             =   225
         Width           =   3240
      End
   End
   Begin VB.Label lblFrameIndex 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "1"
      Height          =   195
      Left            =   14490
      TabIndex        =   25
      Top             =   135
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   1095
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "frmFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iNewEdit As Byte
Dim iAssignMode  As Integer
Dim FullMatrix() As String
Dim szPropertyList As String
Dim szClientList As String
Dim szFunds As String
Private Sub cboDemandTypeCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub chkClient_Click()
   Dim iRow As Integer, i As Integer

   If chkClient.Value Then
      For i = 1 To flxClients.Rows - 1
         flxClients.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxClients.Rows - 2
         flxClients.TextMatrix(i, 0) = "X"
      Next i

      flxClients.row = flxClients.Rows - 1

      flxClients_Click
      chkProperties.Value = 1
   Else
      For i = 1 To flxClients.Rows - 1
         flxClients.TextMatrix(i, 0) = ""
      Next i
      chkProperties.Value = 0
      ConfigFlxProperties
   End If
End Sub

Private Sub chkFund_Click()
   Dim iRow As Integer, i As Integer

   If chkFund.Value Then
      For i = 1 To flxFundWizard.Rows - 1
         flxFundWizard.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxFundWizard.Rows - 1
         flxFundWizard.TextMatrix(i, 0) = "X"
      Next i

      flxFundWizard.row = flxFundWizard.Rows - 1

      'flxFundWizard_Click
   Else
      For i = 1 To flxFundWizard.Rows - 1
         flxFundWizard.TextMatrix(i, 0) = ""
      Next i

      'ConfigflxFundWizardSplit
   End If
End Sub

Private Sub chkProperties_Click()
     Dim iRow As Integer, i As Integer

   If chkProperties.Value Then
      For i = 1 To flxProperty.Rows - 1
         flxProperty.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxProperty.Rows - 1
         flxProperty.TextMatrix(i, 0) = "X"
      Next i

      flxProperty.row = flxProperty.Rows - 1

      'flxProperty_Click
   Else
      For i = 1 To flxProperty.Rows - 1
         If flxProperty.TextMatrix(i, 1) <> "No Property" Then
            flxProperty.TextMatrix(i, 0) = ""
         End If
      Next i

      'ConfigflxPropertySplit
   End If
End Sub

Private Sub chkSel_Click()
    Dim iRow As Integer, i As Integer

   If chkSel.Value Then
      For i = 1 To flxPreviewAssignment.Rows - 1
         flxPreviewAssignment.TextMatrix(i, 0) = ""
      Next i
      For i = 1 To flxPreviewAssignment.Rows - 1
         flxPreviewAssignment.TextMatrix(i, 0) = "X"
      Next i

      flxPreviewAssignment.row = flxPreviewAssignment.Rows - 1

      'flxPreviewAssignment_Click
   Else
      For i = 1 To flxPreviewAssignment.Rows - 1
         flxPreviewAssignment.TextMatrix(i, 0) = ""
      Next i

      'ConfigflxPreviewAssignmentSplit
   End If
End Sub

Private Sub cmdAddNew_Click()
   cmdAddNew.Enabled = False
   cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   flxFund.Enabled = False

   iNewEdit = 1
   Label1(3).Caption = ""
   txtName.text = ""
   txtCode.text = ""
   cboDemandTypeCategory.text = ""
   txtCode.SetFocus
   txtCode.Locked = False
   txtName.Locked = False
   cboDemandTypeCategory.Locked = False
End Sub

Private Sub cmdAssignFund_Click()
    iAssignMode = 1 'false means Assign
    cmdUnassign.Visible = False
    chkFund.Value = 0
    chkClient.Value = 0
    chkProperties.Value = 0
    cmdPreviewAssignment.Caption = "&Preview Assignment"
    cmdSaveAssignment.Caption = "&Save Assignment"
    cmdSaveAssignment.Visible = True
    fraWizardMaster.Left = Frame2.Left + 30
    fraWizardMaster.Top = Frame2.Top + 30
'    fraWizardMaster.Width = 16260
'    fraWizardMaster.Height = 9995
    
    flxPreviewAssignment.Clear
    flxPreviewAssignment.Cols = 5
    flxPreviewAssignment.Rows = 1
    flxPreviewAssignment.RowHeight(0) = 0
    flxPreviewAssignment.ColWidth(0) = 1800
    flxPreviewAssignment.ColWidth(1) = 1600
    flxPreviewAssignment.ColWidth(2) = 2600
    flxPreviewAssignment.ColWidth(3) = 1800
    flxPreviewAssignment.ColWidth(4) = 4000
    
    
    flxProperty.Clear
    Call LoadFundWizard
    fraWizardMaster.Visible = True
    Frame1(0).Enabled = False
    Frame2.Enabled = False
    
    Dim szSQL As String, r As Integer
    Dim adoConn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    'dim https://we.tl/t-nSgKEMpVeq as stirng
'    Dim szHeader As stirng
    '   connect to database
    adoConn.Open getConnectionString
    
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT;"
    adoRST.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly
'    szHeader$ = "|< CLIENT ID|< CLIENT NAME"
'    flxClients.FormatString = szHeader$
    flxClients.ColWidth(0) = 200
    flxClients.ColWidth(1) = 1500
    flxClients.ColWidth(2) = 2800
    flxClients.Rows = 1
    flxClients.Cols = 3
    flxClients.RowHeight(0) = 0
'    flxClients.TextMatrix(0, 1) = "Client ID"
'    flxClients.TextMatrix(0, 2) = "Client Name"
    r = 1
    'flxClients.AddItem ""
    While Not adoRST.EOF
        flxClients.AddItem ""
        flxClients.TextMatrix(r, 1) = adoRST.Fields.Item("CLIENTID").Value
        flxClients.TextMatrix(r, 2) = adoRST.Fields.Item("CLIENTNAME").Value
        
        r = r + 1
        adoRST.MoveNext
    Wend
    adoRST.Close
    adoConn.Close
    Set adoRST = Nothing
    Set adoConn = Nothing
End Sub

Private Sub cmdCancel_Click()
   If MsgBox("Do you like to discard the changes?", vbQuestion + vbYesNo, "Fund") = vbNo Then Exit Sub

   cmdAddNew.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   flxFund.Enabled = True
   txtName.text = ""
   txtCode.text = ""
   cboDemandTypeCategory.text = ""
   Label1(3).Caption = ""

   iNewEdit = 0
End Sub

Private Sub cmdCancelWizard_Click()
    Frame1(0).Enabled = True
    Frame2.Enabled = True
    fraWizardMaster.Visible = False
End Sub

Private Sub cmdClose_Click()
   If iNewEdit <> 0 Then
      If MsgBox("Do you want to close without saving the changes?", vbQuestion + vbYesNo, "Fund") = vbNo Then Exit Sub
   End If

   Unload Me
End Sub

Private Sub cmdCloseFrame_Click()
    Frame1(0).Enabled = True
    Frame2.Enabled = True
    fraWizardMaster.Visible = False
End Sub

Private Sub cmdDelete_Click()
   If flxFund.TextMatrix(flxFund.row, 0) = "" Then Exit Sub

   Dim conConn As New ADODB.Connection

   conConn.Open getConnectionString

   If HasFundUsed(flxFund.TextMatrix(flxFund.row, 0), conConn) Then
      MsgBox "This fund cannot be deleted as it is being used.", vbInformation + vbOKOnly, "Deleting Fund"
   Else
        If MsgBox("Are you sure you wish to delete this fund? ", vbYesNo, "Deleting Fund") = vbYes Then
              conConn.Execute "DELETE * FROM Fund WHERE FundID = " & flxFund.TextMatrix(flxFund.row, 0) & ";"
              flxFund.RemoveItem flxFund.row
              ShowMsgInTaskBar "The Fund has been deleted", "Y", "P"
        End If
   End If

   conConn.Close
   Set conConn = Nothing
End Sub

Private Sub cmdEdit_Click()
   If flxFund.row = 0 Then
      MsgBox "Please select a Fund from the grid.", vbCritical + vbOKOnly, "Fund"
      Exit Sub
   End If

   cmdAddNew.Enabled = False
   cmdEdit.Enabled = False
   cmdSave.Enabled = True
   cmdCancel.Enabled = True
   flxFund.Enabled = False

   iNewEdit = 2
   txtName.Locked = False
   cboDemandTypeCategory.Locked = False
   Label1(3).Caption = flxFund.TextMatrix(flxFund.row, 0)
   txtCode.text = flxFund.TextMatrix(flxFund.row, 1)
   txtName.text = flxFund.TextMatrix(flxFund.row, 2)
   cboDemandTypeCategory.Value = flxFund.TextMatrix(flxFund.row, 4)
   SelTxtInCtrl txtName
   txtName.SetFocus
End Sub

Private Sub cmdGridClose_Click()
    Frame1(0).Enabled = True
    Frame2.Enabled = True
    fraWizardMaster.Visible = False
End Sub

Private Function CreateListOfProp() As Integer
   Dim i As Integer

   szPropertyList = ""
   
   For i = 0 To flxProperty.Rows - 1
      If flxProperty.TextMatrix(i, 0) = "X" Then
         szPropertyList = "'" & flxProperty.TextMatrix(i, 1) & "'" & ", " & szPropertyList
      End If
   Next i
   If Len(szPropertyList) > 2 Then
      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
      CreateListOfProp = Len(szPropertyList)
      Exit Function
   End If
   CreateListOfProp = 0
End Function
Private Function CreateListOfClient() As Integer
   Dim i As Integer

   szClientList = ""
   
   For i = 0 To flxClients.Rows - 1
      If flxClients.TextMatrix(i, 0) = "X" Then
         szClientList = "'" & flxClients.TextMatrix(i, 1) & "'" & ", " & szClientList
      End If
   Next i
   If Len(szClientList) > 2 Then
      szClientList = Left(szClientList, Len(szClientList) - 2)
      CreateListOfClient = Len(szClientList)
      Exit Function
   End If
   CreateListOfClient = 0
End Function
Private Function SelFunds() As Integer
   szFunds = ""
    Dim i As Integer
   For i = 1 To flxFundWizard.Rows - 1
      If flxFundWizard.TextMatrix(i, 0) = "X" Then
         szFunds = flxFundWizard.TextMatrix(i, 4) & ", " & szFunds
      End If
   Next i

   If szFunds = "" Then
        SelFunds = 0
   Else
        szFunds = Left(szFunds, Len(szFunds) - 2)
   End If
End Function
Private Sub cmdPreviewAssignment_Click()
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim iRow As Integer
    Dim a As Integer
    Dim b As Integer
    iRow = 1
    flxPreviewAssignment.Clear
    flxPreviewAssignment.Cols = 7
    flxPreviewAssignment.Rows = 1
    flxPreviewAssignment.RowHeight(0) = 0
    flxPreviewAssignment.ColWidth(0) = 230
    flxPreviewAssignment.ColWidth(1) = 1600
    flxPreviewAssignment.ColWidth(2) = 2800
    flxPreviewAssignment.ColWidth(3) = 2000
    flxPreviewAssignment.ColWidth(4) = 2600
    flxPreviewAssignment.ColWidth(5) = 6000
    flxPreviewAssignment.ColWidth(6) = 0
  
    Dim rsFundMatrix As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    If iAssignMode = 2 Then '2 means view assignment
            Call CreateListOfProp
            Call SelFunds
            adoConn.Open getConnectionString
            If szFunds = "" Then
                'don't show anything
            ElseIf CreateListOfClient = 0 Then
                MsgBox "Please select a client", vbInformation, "Warning"
            ElseIf szFunds <> "" And szPropertyList <> "" And szClientList <> "" Then
                    szPropertyList = Replace(szPropertyList, "No Property", "")
                rsFundMatrix.Open "Select *  from  fundmatrix where fundID in (" & _
                                    szFunds & ") and PropertyID in (" & szPropertyList & ") and ClientID in (" & szClientList & ") " & _
                                    "AND isDeleted=false order by ClientID,PropertyID,FundID", adoConn, adOpenStatic, adLockReadOnly
                      While Not rsFundMatrix.EOF
                            flxPreviewAssignment.AddItem ""
                            flxPreviewAssignment.TextMatrix(iRow, 1) = rsFundMatrix("ClientID").Value  'clientID
                            flxPreviewAssignment.TextMatrix(iRow, 2) = IIf(rsFundMatrix("PropertyID").Value <> "", rsFundMatrix("PropertyID").Value, "No Property") 'Property ID
                            flxPreviewAssignment.TextMatrix(iRow, 3) = rsFundMatrix("FundCode").Value  'FundCode
                            flxPreviewAssignment.TextMatrix(iRow, 4) = rsFundMatrix("FundName").Value  'FundName
                            flxPreviewAssignment.TextMatrix(iRow, 5) = rsFundMatrix("FundCategory").Value 'CategoryCode
                            flxPreviewAssignment.TextMatrix(iRow, 6) = rsFundMatrix("FundID").Value  'FundID
                            iRow = iRow + 1
                            rsFundMatrix.MoveNext
                     Wend
                     rsFundMatrix.Close
                     Set rsFundMatrix = Nothing
            Else
                 'don't show anything
                'rsFundMatrix.Open "Select *  from from fundmatrix where fundID ='" & strSelectedFunds & "' and PropertyID='" & strSelectedProperties & "'", adoconn, adOpenStatic, adLockReadOnly
            End If
            adoConn.Close
            'Exit Sub
    Else
         For K = 1 To flxProperty.Rows - 1
            If flxProperty.TextMatrix(K, 0) = "X" Then
                For i = 1 To flxFundWizard.Rows - 1
                      If flxFundWizard.TextMatrix(i, 0) = "X" Then
                            flxPreviewAssignment.AddItem ""
                            flxPreviewAssignment.TextMatrix(iRow, 1) = flxProperty.TextMatrix(K, 3)
                            flxPreviewAssignment.TextMatrix(iRow, 2) = flxProperty.TextMatrix(K, 1)
                            flxPreviewAssignment.TextMatrix(iRow, 3) = flxFundWizard.TextMatrix(i, 1) 'FundCode
                            flxPreviewAssignment.TextMatrix(iRow, 4) = flxFundWizard.TextMatrix(i, 2) 'FundName
                            flxPreviewAssignment.TextMatrix(iRow, 5) = flxFundWizard.TextMatrix(i, 3) 'CategoryCode
                            flxPreviewAssignment.TextMatrix(iRow, 6) = flxFundWizard.TextMatrix(i, 4) 'FundID
                            iRow = iRow + 1
                      End If
                Next i
             End If
         Next K
    End If
    
    
    'adoconon.Close
'Leave the cross matrix and  load the flxPreviewAssignment values from the database and then filter them as per selected condition
'           Exit Sub
       If iAssignMode = 1 Or iAssignMode = 2 Then '1 means assign mode on  2 means you are viewing assignment from database

                'this will remove the duplicate values in client and properties
                Dim PropertyArray() As String
                Dim ClientArray() As String
                ReDim PropertyArray(flxPreviewAssignment.Rows - 1, 0)
                ReDim ClientArray(flxPreviewAssignment.Rows - 1, 0)
                'saving all property ID,client ID in an array
                For a = 1 To flxPreviewAssignment.Rows - 2
                       PropertyArray(a, 0) = flxPreviewAssignment.TextMatrix(a, 2)
                       ClientArray(a, 0) = flxPreviewAssignment.TextMatrix(a, 1)
                Next a
            'tree build ignoring no property case
                For a = 1 To flxPreviewAssignment.Rows - 2
                    For b = a + 1 To flxPreviewAssignment.Rows - 1
                        If flxPreviewAssignment.TextMatrix(a, 2) = flxPreviewAssignment.TextMatrix(b, 2) And flxPreviewAssignment.TextMatrix(a, 2) <> "No Property" Then
                             ' duplicate value is found in properties
                               flxPreviewAssignment.TextMatrix(b, 2) = ""
                        End If
                    Next b
                Next a
        
            'tree building only for no property
                For a = 1 To flxPreviewAssignment.Rows - 1
                     If flxPreviewAssignment.TextMatrix(a, 2) = "No Property" Then
                            If flxPreviewAssignment.TextMatrix(a, 2) = PropertyArray(a - 1, 0) Then
                                If flxPreviewAssignment.TextMatrix(a, 2) = PropertyArray(a - 1, 0) And flxPreviewAssignment.TextMatrix(a, 1) <> ClientArray(a - 1, 0) Then
                                Else
                                        flxPreviewAssignment.TextMatrix(a, 2) = ""
                                End If
                                        
                            End If
                     End If
                Next a
            'tree building only for Client ID
                For a = 1 To flxPreviewAssignment.Rows - 2
                    For b = a + 1 To flxPreviewAssignment.Rows - 1
                        If flxPreviewAssignment.TextMatrix(a, 1) = flxPreviewAssignment.TextMatrix(b, 1) Then
                             ' duplicate value is found  in client
                             flxPreviewAssignment.TextMatrix(b, 1) = ""
                        End If
                    Next b
                Next a

       End If
End Sub

Private Sub cmdSave_Click()
   If iNewEdit = 0 Then Exit Sub
   If Trim(txtCode.text) = "" Then
      MsgBox "Please type the Code of the Fund.", vbCritical + vbOKOnly, "Fund"
      txtCode.text = ""
      txtCode.SetFocus
      Exit Sub
   End If
   
   Dim iRow As Integer
   
   For iRow = 1 To flxFund.Rows - 1
      If iNewEdit = 1 Then             'New mode
         If flxFund.TextMatrix(iRow, 1) = txtCode.text Then
            MsgBox "The Code cannot be duplicated. Please enter another code.", vbCritical + vbOKOnly, "Fund"
            txtCode.text = ""
            txtCode.SetFocus
            Exit Sub
         End If
      End If
      If iNewEdit = 2 Then       'Edit mode
         If flxFund.TextMatrix(iRow, 1) = txtCode.text And _
               Label1(3).Caption <> flxFund.TextMatrix(iRow, 0) Then
            MsgBox "The Code cannot be duplicated. Please enter another code.", vbCritical + vbOKOnly, "Fund"
            txtCode.text = ""
            txtCode.SetFocus
            Exit Sub
         End If
      End If
   Next iRow

   If Trim(txtCode.text) = "" Then
      MsgBox "Please type the Code of the Fund.", vbCritical + vbOKOnly, "Fund"
      txtCode.text = ""
      txtCode.SetFocus
      Exit Sub
   End If
   If Trim(txtName.text) = "" Then
      MsgBox "Please type the Name of the Fund.", vbCritical + vbOKOnly, "Fund"
      txtName.text = ""
      cboDemandTypeCategory.text = ""
      Label1(3).Caption = ""
      txtName.SetFocus
      Exit Sub
   End If
   If cboDemandTypeCategory.text = "" Then
      MsgBox "Please select the category.", vbCritical + vbOKOnly, "Fund"
      cboDemandTypeCategory.SetFocus
      Exit Sub
   End If

   On Error GoTo ErrorHandler

   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   If iNewEdit = 1 Then
      szSQL = "INSERT INTO FUND (FundCode, FundName, CategoryCode) " & _
              "VALUES ('" & Trim(txtCode.text) & "', '" & Trim(txtName.text) & "', " & cboDemandTypeCategory.Value & ");"
      adoConn.Execute szSQL
      adoConn.Execute "UPDATE Fund SET szFundID = CSTR(FundID);"
      MsgBox "Fund has been added successfully.", vbInformation + vbOKOnly, "Fund"
   End If
   If iNewEdit = 2 Then
      szSQL = "UPDATE FUND " & _
              "SET FundName = '" & Trim(txtName.text) & "', " & _
                  "FundCode = '" & Trim(txtCode.text) & "', " & _
                  "CategoryCode = " & cboDemandTypeCategory.Value & " " & _
              "WHERE FundID = " & Val(flxFund.TextMatrix(flxFund.row, 0)) & ";"
      adoConn.Execute szSQL
      MsgBox "Fund has been edited successfully.", vbInformation + vbOKOnly, "Fund"
   End If

   cmdAddNew.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
   cmdCancel.Enabled = False
   flxFund.Enabled = True
   Label1(3).Caption = ""
   txtName.text = ""
   txtCode.text = ""
   cboDemandTypeCategory.text = ""

   szSQL = "SELECT F.FundID, F.FundCode, F.FundName, SC.Value AS CategoryCode, SC.Code AS CatID " & _
           "FROM FUND AS F, SecondaryCode AS SC " & _
           "WHERE F.CategoryCode = CINT(SC.Code) AND " & _
               "SC.PrimaryCode = 'DCTG' " & _
           "ORDER BY F.FundID;"
   populateGridDefinedHeader adoConn, szSQL, flxFund

   adoConn.Close
   Set adoConn = Nothing

   iNewEdit = 0
   FocusControl cmdAddNew
   Exit Sub
ErrorHandler:
   If iNewEdit = 1 Then MsgBox "System could not add new Fund." & Chr(13) & Err.Number & " " & Err.description, vbCritical + vbOKOnly, "Error"
   If iNewEdit = 2 Then MsgBox "System could not edit Fund." & Chr(13) & Err.Number & " " & Err.description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub LoadDemandCategory(adoConn As ADODB.Connection)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer
   
   cboDemandTypeCategory.Clear

   szSQL = "SELECT Code, Value " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'DCTG';"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.RecordCount < 1 Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If
 
   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To adoRST.RecordCount - 1
       For j = 0 To adoRST.Fields.Count - 1
           Data(j, i) = adoRST.Fields(j)
       Next j
       adoRST.MoveNext
   Next i

   cboDemandTypeCategory.Column() = Data()
   
   adoRST.Close
   Set adoRST = Nothing
End Sub



Private Sub addPropertiesTowizard(strClientID As String, adoConn As ADODB.Connection, ByRef iRow As Integer)
        Dim szSQL As String
        Dim adoRST As New ADODB.Recordset
        'Dim iRow As Integer
        szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
               "FROM  PROPERTY where clientID = '" & strClientID & "'" & _
               "ORDER BY ClientID,PROPERTYID;"
        adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
        'iRow = 1

        While Not adoRST.EOF 'While Not adoRst.EOF
            flxProperty.AddItem ""
            flxProperty.TextMatrix(iRow, 0) = "X"
            flxProperty.TextMatrix(iRow, 1) = adoRST("PROPERTYID").Value
            flxProperty.TextMatrix(iRow, 2) = adoRST("PROPERTYNAME").Value
            flxProperty.TextMatrix(iRow, 3) = adoRST("ClientID").Value
            flxProperty.RowHeight(iRow) = 280
            iRow = iRow + 1
    
            adoRST.MoveNext
        'End If
        Wend
        'for No Properties
            flxProperty.AddItem ""
            flxProperty.TextMatrix(iRow, 0) = "X"
            flxProperty.TextMatrix(iRow, 1) = "No Property"
            flxProperty.TextMatrix(iRow, 2) = "No Property"
            flxProperty.TextMatrix(iRow, 3) = strClientID
            flxProperty.RowHeight(iRow) = 0
            iRow = iRow + 1
End Sub

Private Sub cmdSaveAssignment_Click()
    Dim adoConn As New ADODB.Connection
    Dim rsFundMatrix As New ADODB.Recordset
    Dim K As Integer
    Dim i As Integer
    Dim iRow As Integer
    iRow = 1
    ''Exit Sub added by anol 2021-05-20
    cmdPreviewAssignment_Click
    'Now Need to add validation if there is single client selected
    adoConn.Open getConnectionString
    'Load the full matrix in an array
        ReDim FullMatrix(flxPreviewAssignment.Rows, 5)
        For K = 1 To flxProperty.Rows - 1
            If flxProperty.TextMatrix(K, 0) = "X" Then
                For i = 1 To flxFundWizard.Rows - 1
                      If flxFundWizard.TextMatrix(i, 0) = "X" Then
                           FullMatrix(iRow, 0) = flxProperty.TextMatrix(K, 3) 'client ID
                           FullMatrix(iRow, 1) = flxProperty.TextMatrix(K, 1) ' Property ID
                           If FullMatrix(iRow, 1) = "No Property" Then
                                FullMatrix(iRow, 1) = ""
                           End If
                           FullMatrix(iRow, 2) = flxFundWizard.TextMatrix(i, 4) ' FundID
                           FullMatrix(iRow, 3) = flxFundWizard.TextMatrix(i, 2) ' Fund name
                           FullMatrix(iRow, 4) = flxFundWizard.TextMatrix(i, 3) ' Fund Category
                           FullMatrix(iRow, 5) = flxFundWizard.TextMatrix(i, 1) ' Fund Code
                           iRow = iRow + 1
                      End If
                Next i
            End If
        Next K
      If iAssignMode = 1 Then '1 means assign mode on
            For i = 1 To UBound(FullMatrix())
                 If FullMatrix(i, 2) <> "" Then
                     rsFundMatrix.Open "Select * from FundMatrix where fundID=" & _
                         FullMatrix(i, 2) & " And PropertyID='" & FullMatrix(i, 1) & "'  And clientID='" & FullMatrix(i, 0) & "' and isDeleted=false", adoConn, adOpenDynamic, adLockOptimistic
                         If rsFundMatrix.EOF Then
                                 rsFundMatrix.AddNew
                                 rsFundMatrix!ClientID = FullMatrix(i, 0)
                                 rsFundMatrix!propertyID = FullMatrix(i, 1)
                                 rsFundMatrix!fundID = FullMatrix(i, 2)
                                 rsFundMatrix!FundName = FullMatrix(i, 3)
                                 rsFundMatrix!FundCategory = FullMatrix(i, 4)
                                 rsFundMatrix!FundCode = Trim(FullMatrix(i, 5))
                                 rsFundMatrix.Update
                         End If
                         rsFundMatrix.Close
                         Set rsFundMatrix = Nothing
                 End If
            Next i
       Else
              For i = 1 To flxPreviewAssignment.Rows - 1
                        If flxPreviewAssignment.TextMatrix(i, 0) = "X" Then
                              adoConn.Execute "Update FundMatrix set isDeleted=true where fundID=" & flxPreviewAssignment.TextMatrix(i, 6) & " AND PropertyID ='" & flxPreviewAssignment.TextMatrix(i, 2) & "'   AND clientID ='" & flxPreviewAssignment.TextMatrix(i, 1) & "'  "
                        End If
              Next i
       End If
     
       adoConn.Close
       Set adoConn = Nothing
       If iAssignMode = 1 Then 'assign mode on
            MsgBox "Your Fund assignment has been saved.", vbInformation, "OK"
            FocusControl cmdCloseFrame
       Else
            MsgBox "Fund unassignment has been saved.", vbInformation, "OK"
            FocusControl cmdCloseFrame
       End If
End Sub

Private Sub cmdUnassign_Click()
    cmdSaveAssignment.Visible = False
    Dim adoConn As New ADODB.Connection
    Dim rsFundMatrix As New ADODB.Recordset
    Dim K As Integer
    Dim i As Integer
    Dim iRow As Integer
    iRow = 1
    'Now Need to add validation if there is single client selected
    adoConn.Open getConnectionString

    Call CreateListOfProp
    Call SelFunds

    If szFunds = "" Then
        'don't show anything
    ElseIf szFunds <> "" And szPropertyList <> "" Then
            szPropertyList = Replace(szPropertyList, "No Property", "")
        rsFundMatrix.Open "Select *  from  fundmatrix where fundID in (" & _
                            szFunds & ") and PropertyID in (" & szPropertyList & ") AND isDeleted=false order by ClientID,PropertyID,FundID", adoConn, adOpenStatic, adLockReadOnly
        ReDim FullMatrix(rsFundMatrix.RecordCount + 1, 5)
              While Not rsFundMatrix.EOF
                    FullMatrix(iRow, 0) = rsFundMatrix("ClientID").Value 'clientID
                    FullMatrix(iRow, 1) = rsFundMatrix("PropertyID").Value  ' IIf(rsFundMatrix("PropertyID").Value <> "", rsFundMatrix("PropertyID").Value, "No Property")  'Property ID
                    FullMatrix(iRow, 2) = rsFundMatrix("FundCode").Value   'FundCode
                    FullMatrix(iRow, 3) = rsFundMatrix("FundName").Value 'FundName
                    FullMatrix(iRow, 4) = rsFundMatrix("FundCategory").Value  'CategoryCode
                    FullMatrix(iRow, 5) = rsFundMatrix("FundID").Value   'FundID
                    iRow = iRow + 1
                    rsFundMatrix.MoveNext
             Wend
             rsFundMatrix.Close
             Set rsFundMatrix = Nothing
    Else
         'don't show anything
    End If

    For i = 1 To flxPreviewAssignment.Rows - 1
              If flxPreviewAssignment.TextMatrix(i, 0) = "X" Then
                    adoConn.Execute "Update FundMatrix set isDeleted=true where fundID=" & FullMatrix(i, 5) & " AND PropertyID ='" & FullMatrix(i, 1) & "'   AND clientID ='" & FullMatrix(i, 0) & "'  "
              End If
    Next i
     
    adoConn.Close
    Set adoConn = Nothing
    Call cmdPreviewAssignment_Click 'view the final results
    MsgBox "Fund unassignment data has been saved.", vbInformation, "OK"
    
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdViewAssignment_Click()
    iAssignMode = 2 'false means view assignment
    cmdUnassign.Visible = True
    chkFund.Value = 0
    chkClient.Value = 0
    chkProperties.Value = 0
    cmdPreviewAssignment.Caption = "&Preview saved assignment"
    cmdSaveAssignment.Visible = False
    
    fraWizardMaster.Left = Frame2.Left + 30
    fraWizardMaster.Top = Frame2.Top + 30
'    fraWizardMaster.Width = 16260
'    fraWizardMaster.Height = 10995
    
    flxPreviewAssignment.Clear
    flxPreviewAssignment.Cols = 5
    flxPreviewAssignment.Rows = 1
    flxPreviewAssignment.RowHeight(0) = 0
    flxPreviewAssignment.ColWidth(0) = 1800
    flxPreviewAssignment.ColWidth(1) = 1600
    flxPreviewAssignment.ColWidth(2) = 2600
    flxPreviewAssignment.ColWidth(3) = 1800
    flxPreviewAssignment.ColWidth(4) = 2600
    
    
    flxProperty.Clear
    Call LoadFundWizard
    fraWizardMaster.Visible = True
    Frame1(0).Enabled = False
    Frame2.Enabled = False
    
    Dim szSQL As String, r As Integer
    Dim adoConn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    'dim https://we.tl/t-nSgKEMpVeq as stirng
'    Dim szHeader As stirng
    '   connect to database
    adoConn.Open getConnectionString
    
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT;"
    adoRST.Open szSQL, adoConn, adOpenDynamic, adLockReadOnly
'    szHeader$ = "|< CLIENT ID|< CLIENT NAME"
'    flxClients.FormatString = szHeader$
    flxClients.ColWidth(0) = 200
    flxClients.ColWidth(1) = 1500
    flxClients.ColWidth(2) = 4000
    flxClients.Rows = 1
    flxClients.Cols = 3
    flxClients.RowHeight(0) = 0

    r = 1
    'flxClients.AddItem ""
    While Not adoRST.EOF
        flxClients.AddItem ""
        flxClients.TextMatrix(r, 1) = adoRST.Fields.Item("CLIENTID").Value
        flxClients.TextMatrix(r, 2) = adoRST.Fields.Item("CLIENTNAME").Value
        
        r = r + 1
        adoRST.MoveNext
    Wend
    adoRST.Close
    adoConn.Close
    Set adoRST = Nothing
    Set adoConn = Nothing
End Sub

Private Sub flxClients_Click()
    Dim iCount As Integer
    Dim iRow As Integer
    Dim szSQL As String
    Dim adoConn As New ADODB.Connection
    Dim adoRST As New ADODB.Recordset
    adoConn.Open getConnectionString
    SelectFlxGridRow 0, flxClients, flxClients.row
    flxProperty.Clear
    Call ConfigFlxProperties
    Dim i As Integer
    i = 1
    For iRow = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(iRow, 0) = "X" Then
                iCount = iCount + 1
                addPropertiesTowizard flxClients.TextMatrix(iRow, 1), adoConn, i
                'i = i + 1
            End If
    Next

    adoConn.Close
    Set adoConn = Nothing
End Sub
Private Sub ConfigFlxProperties()
    Dim szHeader As String
    szHeader$ = "<|<Property ID|<Property Name|<"
    flxProperty.FormatString = szHeader
    flxProperty.Clear
    flxProperty.Cols = 4
    flxProperty.Rows = 1
'    flxProperty.TextMatrix(0, 1) = "Property ID"
'    flxProperty.TextMatrix(0, 2) = "Property Name"
    flxProperty.RowHeight(0) = 0
    flxProperty.ColWidth(0) = 300                   '"X"
    flxProperty.ColWidth(1) = 900                'Property ID
    flxProperty.ColWidth(2) = 2150                'Property Name
    flxProperty.ColWidth(3) = 1100                'Client ID
End Sub

Private Sub flxFund_Click()
     Label1(3).Caption = flxFund.TextMatrix(flxFund.row, 0)
   txtCode.text = flxFund.TextMatrix(flxFund.row, 1)
   txtName.text = flxFund.TextMatrix(flxFund.row, 2)
   cboDemandTypeCategory.Value = flxFund.TextMatrix(flxFund.row, 4)
End Sub

Private Sub flxFundWizard_Click()
    SelectFlxGridRow 0, flxFundWizard, flxFundWizard.row
End Sub
Private Function SelectFlxGridRow2(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
'      conFlxGrid.row = iSelRow
'      For iRow = 1 To conFlxGrid.Cols - 1
'         conFlxGrid.col = iRow
'         conFlxGrid.CellBackColor = RGB(255, 255, 255)
'      Next iRow
      SelectFlxGridRow2 = -1
   Else
        'Here I have Implemented if no value in the grid row then do not select anol 2020-11-04
      
            conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
            conFlxGrid.row = iSelRow
'            For iRow = 1 To conFlxGrid.Cols - 1
'               conFlxGrid.col = iRow
'               conFlxGrid.CellBackColor = RGB(174, 179, 233)
'            Next iRow
            SelectFlxGridRow2 = 1
   End If
End Function
Private Function SelectFlxGridRow(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
'      conFlxGrid.row = iSelRow
'      For iRow = 1 To conFlxGrid.Cols - 1
'         conFlxGrid.col = iRow
'         conFlxGrid.CellBackColor = RGB(255, 255, 255)
'      Next iRow
      SelectFlxGridRow = -1
   Else
        'Here I have Implemented if no value in the grid row then do not select anol 2020-11-04
      If conFlxGrid.TextMatrix(iSelRow, iColID + 1) <> "" Then
            conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
            conFlxGrid.row = iSelRow
'            For iRow = 1 To conFlxGrid.Cols - 1
'               conFlxGrid.col = iRow
'               conFlxGrid.CellBackColor = RGB(174, 179, 233)
'            Next iRow
            SelectFlxGridRow = 1
      Else
            SelectFlxGridRow = -1
      End If
   End If
End Function

Private Sub flxPreviewAssignment_Click()
    If iAssignMode = 0 Then ' Mode 0 means assign mode is not on
        SelectFlxGridRow 0, flxPreviewAssignment, flxPreviewAssignment.row
    ElseIf iAssignMode = 1 Or iAssignMode = 2 Then
        SelectFlxGridRow2 0, flxPreviewAssignment, flxPreviewAssignment.row
    End If
End Sub

Private Sub flxProperty_Click()
         SelectFlxGridRow 0, flxProperty, flxProperty.row
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
    Me.ZOrder 0
    Me.Height = 12015
    Me.Width = 16905
    Me.BackColor = MODULEBACKCOLOR
    Frame1(0).BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    fraWizardMaster.BackColor = &HE9E8E7   'MODULEBACKCOLOR
    Frame1(1).BackColor = &HE9E8E7 ' MODULEBACKCOLOR
    Frame1(2).BackColor = &HE9E8E7 'MODULEBACKCOLOR
    Frame1(3).BackColor = &HE9E8E7 'MODULEBACKCOLOR
    Frame1(4).BackColor = &HE9E8E7 ' MODULEBACKCOLOR
    chkFund.BackColor = MODULEBACKCOLOR
    chkClient.BackColor = MODULEBACKCOLOR
    chkProperties.BackColor = MODULEBACKCOLOR
    
    txtCode.Locked = True
    txtName.Locked = True
    cboDemandTypeCategory.Locked = True
    Dim adoConn As New ADODB.Connection
    Dim szSQL As String

   adoConn.Open getConnectionString
   szSQL = "SELECT F.FundID, F.FundCode, F.FundName, SC.Value AS CategoryCode, SC.Code AS CatID " & _
           "FROM FUND AS F, SecondaryCode AS SC " & _
           "WHERE F.CategoryCode = CINT(SC.Code) AND " & _
               "SC.PrimaryCode = 'DCTG' " & _
           "ORDER BY F.FundID;"
'Debug.Print szSQL
   ConfigFlxFund
   populateGridDefinedHeader adoConn, szSQL, flxFund

   LoadDemandCategory adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadFundWizard()
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    Dim szSQL As String
    Dim iRow As Integer
    Call ConfigFlxFundWizard
    adoConn.Open getConnectionString
    szSQL = "SELECT F.FundID, F.FundCode, F.FundName, SC.Value AS CategoryCode, SC.Code AS CatID " & _
           "FROM FUND AS F, SecondaryCode AS SC " & _
           "WHERE F.CategoryCode = CINT(SC.Code) AND " & _
           "SC.PrimaryCode = 'DCTG' " & _
           "ORDER BY F.FundCode;"
    rsFund.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
    iRow = 1
    While Not rsFund.EOF
            flxFundWizard.AddItem ""
            flxFundWizard.TextMatrix(iRow, 1) = Space(1) & CStr(rsFund("FundCode").Value) ' I use this space to make the column make left align later we need to remove it
'            flxFundWizard.ColAlignment(1) = vbAlignLeft
            flxFundWizard.TextMatrix(iRow, 2) = rsFund("FundName").Value
            flxFundWizard.TextMatrix(iRow, 3) = rsFund("CategoryCode").Value
            flxFundWizard.TextMatrix(iRow, 4) = rsFund("FundID").Value
            rsFund.MoveNext
            iRow = iRow + 1
            
    Wend
    rsFund.Close
    Set rsFund = Nothing
    adoConn.Close
    Set adoConn = Nothing
    
End Sub
Private Sub ConfigFlxFundWizard()
   Dim szHeader As String, iCol As Integer

   flxFundWizard.Clear
   flxFundWizard.Cols = 5
   flxFundWizard.Rows = 1
   flxFundWizard.RowHeight(0) = 0
   szHeader$ = "<FundID|<FundCode|<FundName|<CategoryCode|CatID"
   flxFundWizard.FormatString = szHeader$
   flxFundWizard.ColWidth(0) = 230
   flxFundWizard.ColWidth(1) = 1400
   flxFundWizard.ColWidth(2) = 2090
   flxFundWizard.ColWidth(3) = 1590
   flxFundWizard.ColWidth(4) = 0
End Sub
Private Sub ConfigFlxFund()
   Dim szHeader As String, iCol As Integer

   flxFund.Clear
   flxFund.Cols = 5
   flxFund.Rows = 2
   flxFund.RowHeight(0) = 0
   szHeader$ = "<FundID|<FundCode|<FundName|<CategoryCode|CatID"
   flxFund.FormatString = szHeader$

   flxFund.ColWidth(0) = Label1(2).Left - Label1(1).Left
   Debug.Print Label1(2).Left - Label1(1).Left
   flxFund.ColWidth(1) = Label1(7).Left - Label1(2).Left
   Debug.Print Label1(7).Left - Label1(2).Left
   flxFund.ColWidth(2) = Label1(4).Left - Label1(7).Left
   Debug.Print Label1(4).Left - Label1(7).Left
   flxFund.ColWidth(3) = flxFund.Width + flxFund.Left - Label1(4).Left - 120
   Debug.Print flxFund.Width + flxFund.Left - Label1(4).Left - 120
   flxFund.ColWidth(4) = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
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

Private Sub txtCode_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   txtCode.text = UCase(txtCode.text)
End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

