VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManagementFeeSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Management Fee - Options"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManagementFeeSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   9570
   Visible         =   0   'False
   Begin VB.CheckBox ChkAllClient 
      Appearance      =   0  'Flat
      Caption         =   "Select All Clients"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   765
      Width           =   1815
   End
   Begin VB.TextBox txtComparenextDueDate1 
      Height          =   330
      Left            =   10350
      TabIndex        =   28
      Top             =   8235
      Width           =   645
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
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8820
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtAssignedProperty 
      Alignment       =   2  'Center
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
      Left            =   6930
      MaxLength       =   10
      TabIndex        =   25
      Top             =   8820
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkAssignProperty 
      Caption         =   "Do not assign a property when charging"
      Height          =   285
      Left            =   2250
      TabIndex        =   24
      Top             =   8820
      Width           =   4830
   End
   Begin VB.TextBox txtLastChargeDate 
      Alignment       =   2  'Center
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
      Left            =   4545
      MaxLength       =   10
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   5085
      ScaleHeight     =   450
      ScaleWidth      =   2655
      TabIndex        =   21
      Top             =   5175
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while loading ...."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   270
         TabIndex        =   22
         Top             =   135
         Width           =   2745
      End
   End
   Begin VB.TextBox txtClientSearch 
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
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   3
      Top             =   405
      Width           =   1575
   End
   Begin VB.TextBox txtPostingDate 
      Alignment       =   2  'Center
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
      Left            =   7230
      MaxLength       =   10
      TabIndex        =   2
      Top             =   80
      Width           =   1575
   End
   Begin VB.PictureBox fmeLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   3855
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label lblLoading 
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while system processing email..."
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   15
         Width           =   3675
      End
   End
   Begin VB.CommandButton cmdGDPOk 
      Caption         =   "&Ok"
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
      Left            =   5895
      TabIndex        =   10
      Top             =   9315
      Width           =   1440
   End
   Begin VB.TextBox txtChargeDate 
      Alignment       =   2  'Center
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
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   0
      Top             =   80
      Width           =   1575
   End
   Begin VB.CheckBox chkDT 
      Appearance      =   0  'Flat
      Caption         =   "Select All Demand Types"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5610
      TabIndex        =   6
      Top             =   4710
      Width           =   2055
   End
   Begin VB.CheckBox chkProp 
      Appearance      =   0  'Flat
      Caption         =   "Select All Properties"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4470
      TabIndex        =   7
      Top             =   765
      Width           =   1815
   End
   Begin VB.CommandButton cmdGDPCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   7695
      TabIndex        =   11
      Top             =   9315
      Width           =   1440
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClients 
      Height          =   3585
      Left            =   60
      TabIndex        =   5
      Top             =   1110
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   3585
      Left            =   4365
      TabIndex        =   9
      Top             =   1110
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDemandTypes 
      Height          =   3645
      Left            =   3735
      TabIndex        =   8
      Top             =   5040
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
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
      _Band(0).Cols   =   4
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCategory 
      Height          =   3645
      Left            =   60
      TabIndex        =   12
      Top             =   5040
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   6429
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483630
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
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned Property:"
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
      Left            =   6030
      TabIndex        =   27
      Top             =   9135
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Charge Date:"
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
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   23
      Top             =   45
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date:"
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
      Height          =   195
      Index           =   0
      Left            =   6225
      TabIndex        =   20
      Top             =   75
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Category:"
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
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   4800
      Width           =   1065
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Date :"
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
      Height          =   195
      Index           =   30
      Left            =   60
      TabIndex        =   16
      Top             =   75
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Names:"
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
      Height          =   195
      Left            =   3780
      TabIndex        =   15
      Top             =   4770
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client (Search)"
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
      Left            =   60
      TabIndex        =   14
      Top             =   405
      Width           =   1035
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   4440
      TabIndex        =   13
      Top             =   450
      Width           =   645
   End
End
Attribute VB_Name = "frmManagementFeeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iSelDemandCategory  As Integer
Public iSelDemandTypes     As Integer
Public iSelProperties      As Integer     'Number of properties selected by the user
Public szCallingFrom       As String

Private Type SendDemandByEmail
   szLesseeID    As String
   szLesseeEmail As String
   szClient      As String
   colAtt        As Collection
   colURN        As Collection
End Type

Private uLessee()   As SendDemandByEmail
Private iLes        As Integer, szEmailDmdIdList   As String
Private szaProp()   As String, iProperty           As Integer
Private szaDT()     As String, iDT                 As Integer
Private iSelClient  As Integer
Public boldatechaged As Boolean
Dim icounter1 As Long
'  frmManagementFeeSelection.szCallingFrom = "ManagementFee Preview" one style

Private Sub ChkAllClient_Click()
    Dim i As Integer

   If ChkAllClient.Value = 1 Then
      For i = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(i, 0) <> "X" Then
            SelectFlxGridRow 0, flxClients, i
            iSelClient = iSelClient + 1
         End If
      Next i
      iSelProperties = flxClients.Rows - 1
   Else
      For i = 1 To flxClients.Rows - 1
        iSelClient = 0
         If flxClients.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxClients, i
            'iSelClient = iSelClient + 1
         End If
      Next i
      iSelProperties = 0
   End If
   FilterProperties
'   If iSelProperties > 0 Then
'        LoadFlxGrids
'    End If
End Sub

Private Sub chkAssignProperty_Click()
'    If chkAssignProperty.Value = 1 Then
'         Dim szSelectedClient As String
'        Dim rCount As Integer
'         For rCount = 1 To flxClients.Rows - 1
'                If flxClients.TextMatrix(rCount, 0) = "X" Then
'                   szSelectedClient = flxClients.TextMatrix(rCount, 1)
'                   Exit For
'                End If
'            Next
'           frmSelectionProperty.szSelectedClientID = szSelectedClient
'           If szSelectedClient = "" Then
'                MsgBox "Please selelct a client"
'                Exit Sub
'           End If
'        frmSelectionProperty.Show
'        frmSelectionProperty.ZOrder 0
'    Else
'        txtAssignedProperty.text = ""
'    End If
End Sub

Private Sub chkDT_Click()
    Dim i As Integer
    iSelProperties = 0
    iSelDemandTypes = 0
    For i = 1 To flxDemandTypes.Rows - 1
        'If flxDemandTypes.TextMatrix(i, 0) = "X" Then
         '   iSelProperties = iSelProperties + 1
            If chkDT.Value = 1 Then
                  flxDemandTypes.TextMatrix(i, 0) = "X"
                  iSelDemandTypes = iSelDemandTypes + 1
            Else
                  flxDemandTypes.TextMatrix(i, 0) = ""
                  iSelDemandTypes = 0
            End If
'            txtLastChargeDate.text = findLastChargeDate(flxDemandTypes.TextMatrix(i, 1))
        'End If
    Next i
    If chkDT.Value = 1 Then
        FocusControl cmdGDPOk
    End If
End Sub

Private Sub chkDT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDemandTypes.SetFocus
    End If
End Sub

Private Sub chkProp_Click()
   If iSelClient = 0 Then
      chkProp.Value = 0
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      Exit Sub
   End If

   Dim i As Integer

   If chkProp.Value = 1 Then
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) <> "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
      iSelProperties = flxProperties.Rows - 1
   Else
      For i = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxProperties, i
         End If
      Next i
      iSelProperties = 0
   End If
   If iSelProperties > 0 Then
        LoadFlxGrids
    End If
   FilterDemandTypes
End Sub

Private Sub cmdACList_Click(Index As Integer)
    Dim szSelectedClient As String
    Dim rCount As Integer
     For rCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(rCount, 0) = "X" Then
               szSelectedClient = flxClients.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
       frmSelectionProperty.szSelectedClientID = szSelectedClient
        If szSelectedClient = "" Then
                MsgBox "Please selelct a client"
                Exit Sub
           End If
    frmSelectionProperty.Show
    frmSelectionProperty.ZOrder 0
End Sub

'Private Sub chkSCS_Click()
'   If chkSCS.Value Then
'      txtSCDateFrom.Locked = False
'      txtSCDateFrom.SetFocus
'      txtSCDateTo.Locked = False
'   Else
'      txtSCDateFrom.Locked = True
'      txtSCDateTo.Locked = True
'      txtSCDateFrom.text = ""
'      txtSCDateTo.text = ""
'   End If
'End Sub

Private Sub cmdGDPCancel_Click()
   frmDemands3.Enabled = True
   Unload Me
End Sub

Private Sub flxCategory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxDemandTypes.MousePointer = vbArrow
End Sub

Private Sub flxCategory_RowColChange()
   SelectFlxGridRow 0, flxCategory, flxCategory.row
   If flxCategory.TextMatrix(flxCategory.row, 0) = "X" Then
      iSelDemandCategory = iSelDemandCategory + 1
   Else
      iSelDemandCategory = iSelDemandCategory - 1
   End If

   FilterDemandTypes
End Sub

Private Function SelectFlxGridRowNocolor(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
   Dim iRow As Integer

   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
        ' conFlxGrid.CellBackColor = RGB(255, 255, 255)
      'Next iRow
      SelectFlxGridRowNocolor = -1
   Else
      conFlxGrid.TextMatrix(iSelRow, iColID) = "X"

      conFlxGrid.row = iSelRow
      'For iRow = 1 To conFlxGrid.Cols - 1
         'conFlxGrid.col = iRow
         'conFlxGrid.CellBackColor = RGB(174, 179, 233)
      'Next iRow
      SelectFlxGridRowNocolor = 1
   End If
End Function
'Private Function SelectFlxGridRowNocolor(iColID As Integer, conFlxGrid As MSHFlexGrid, iSelRow As Integer) As Integer
'   Dim iRow As Integer
'
'   If conFlxGrid.TextMatrix(iSelRow, iColID) = "X" Then
'      conFlxGrid.TextMatrix(iSelRow, iColID) = ""
'      conFlxGrid.row = iSelRow
'      'For iRow = 1 To conFlxGrid.Cols - 1
'         'conFlxGrid.col = iRow
'        ' conFlxGrid.CellBackColor = RGB(255, 255, 255)
'      'Next iRow
'      SelectFlxGridRowNocolor = -1
'   Else
'      conFlxGrid.TextMatrix(iSelRow, iColID) = "X"
'
'      conFlxGrid.row = iSelRow
'      'For iRow = 1 To conFlxGrid.Cols - 1
'         'conFlxGrid.col = iRow
'         'conFlxGrid.CellBackColor = RGB(174, 179, 233)
'      'Next iRow
'      SelectFlxGridRowNocolor = 1
'   End If
'End Function
Private Sub flxClients_Click()
        'SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
        SelectFlxGridRowNocolor 0, flxClients, flxClients.row
        iSelClient = 1
        
        FilterProperties
        flxDemandTypes.Clear
        flxDemandTypes.Rows = 2
        
        Dim rCount As Integer
        Dim adoConn1 As New ADODB.Connection
        Dim szSelectedClient As String
        Dim szSelectedClientName As String
        Dim PurchaseLedgerControl As String
    
        If flxClients.TextMatrix(flxClients.row, 0) = "X" Then

           szSelectedClient = flxClients.TextMatrix(flxClients.row, 1)
          
                If icounter1 > 1 Then
                
                          'txtChargeDate.text = Format(Date, "dd/MM/yyyy")
                          'txtPostingDate.text = txtChargeDate.text
                          icounter1 = icounter1 + 1
                End If
           
           szSelectedClientName = flxClients.TextMatrix(flxClients.row, 2)
        End If
        If szSelectedClient = "" Then Exit Sub
        If adoConn1.State = 0 Then
            adoConn1.Open getConnectionString
        End If
        
'        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Management Fee Payable (P&L)", szSelectedClient)
'        If (PurchaseLedgerControl = "") Then
'            MsgBox "Please set up Management Fees control accounts for '" & szSelectedClientName & "' "
'            Exit Sub
'        End If
'
'
'        PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Management Fees Control Account (B/S)", szSelectedClient)
'        If (PurchaseLedgerControl = "") Then
'            MsgBox "Please set up Management Fees control accounts for '" & szSelectedClientName & "' "
'            Exit Sub
'        End If
'
        If szSelectedClient = "" Then Exit Sub
'        adoConn1.Open getConnectionString
'        If IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoConn1) = 0 Then
'            MsgBox "The posting date cannot fall within a closed financial period for the client :" & szSelectedClient, vbInformation, "Warning"
'            adoConn1.Close
''            FocusControl txtChargeDate
'            Exit Sub
'        ElseIf IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoConn1) = 9 Then
'            MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
'            adoConn1.Close
''            FocusControl txtChargeDate
'            Exit Sub
'        End If
        adoConn1.Close
        
                
                
                
End Sub

Private Sub flxClients_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
'           SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
           iSelClient = 1
           FilterProperties
           flxDemandTypes.Clear
    End If
End Sub

Private Sub flxClients_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClients_Click
        flxProperties.SetFocus
    End If
End Sub

Private Sub flxDemandTypes_Click()
   'from colrow change to click by anol 16 aug 2016
   SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
    'I want to select only 1 row
'   SelectOnly1RowFlxGrid flxDemandTypes, flxDemandTypes.row, 0
   If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
      iSelDemandTypes = iSelDemandTypes + 1
   Else
      iSelDemandTypes = iSelDemandTypes - 1
   End If
End Sub

Private Sub flxDemandTypes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
         SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
        If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
           iSelDemandTypes = iSelDemandTypes + 1
        Else
           iSelDemandTypes = iSelDemandTypes - 1
        End If
    End If
End Sub

Private Sub flxDemandTypes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        flxDemandTypes_Click
'        optAutoGenSig.SetFocus
    End If
End Sub

Private Sub flxProperties_Click()
    'from colrow change to click by anol 16 Aug 2016
   SelectFlxGridRow 0, flxProperties, flxProperties.row
   'I want to select only 1 row
'    SelectOnly1RowFlxGrid flxProperties, flxProperties.row, 0
'   If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'      iSelProperties = iSelProperties + 1
'   Else
'      iSelProperties = iSelProperties - 1
'   End If
            'modified by anol 20160920
    Dim i As Integer
    iSelProperties = 0
    For i = 1 To flxProperties.Rows - 1
        If flxProperties.TextMatrix(i, 0) = "X" Then
            iSelProperties = iSelProperties + 1
'            txtLastChargeDate.text = findLastChargeDate(flxProperties.TextMatrix(i, 1))
'            If txtLastChargeDate.text = "" Then
'                txtLastChargeDate.Locked = False
'            Else
'                txtLastChargeDate.Locked = True
'            End If
        End If
    Next i
    If iSelProperties > 0 Then
        LoadFlxGrids
    End If
    FilterDemandTypes
    
    
End Sub
Private Function findLastChargeDate(strPropertyID As String)
    Dim adoconn As New ADODB.Connection
    Dim rsChargedate As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsChargedate.Open "Select max(R.ChargeDate) as chrgDate from tlbReceiptSplit S,tlbreceipt R,Units U where U.UnitNumber=R.UnitID AND U.PropertyID='" & strPropertyID & "'", adoconn, adOpenStatic, adLockReadOnly
    If Not rsChargedate.EOF Then
        findLastChargeDate = IIf(IsNull(rsChargedate("chrgDate").Value) = True, "", rsChargedate("chrgDate").Value)
    End If
    rsChargedate.Close
    Set rsChargedate = Nothing
    adoconn.Close
    Set adoconn = Nothing

End Function

Private Sub flxProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
           SelectFlxGridRow 0, flxProperties, flxProperties.row
'           If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'              iSelProperties = iSelProperties + 1
'           Else
'              iSelProperties = iSelProperties - 1
'           End If
            'modified by anol 20160920
            Dim i As Integer
            iSelProperties = 0
            For i = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(i, 0) = "X" Then
                    iSelProperties = iSelProperties + 1
                End If
            Next i
           FilterDemandTypes
    End If
End Sub

Private Sub flxProperties_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxProperties_Click
        chkDT.SetFocus
        'Call flxProperties_Click
    End If
End Sub





Private Sub txtClientSearch_Change()
            'Updated by anol 22 Dec 2015
   Dim i As Integer
'   boldatechaged = False
    For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.TextMatrix(i, 0) = ""
   Next i
   
   For i = flxClients.Rows - 1 To 1 Step -1
            flxClients.RowHeight(i) = 240
            If InStr(1, UCase(flxClients.TextMatrix(i, 1)), UCase(txtClientSearch.text), vbTextCompare) = 0 And txtClientSearch.text <> "" Then
                flxClients.RowHeight(i) = 0
            End If

      If flxClients.RowHeight(i) = 240 Then
            flxClients.row = i
      End If
   Next i
End Sub

Private Sub txtClientSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        flxClients.SetFocus
    End If
End Sub



Private Sub txtChargeDate_Change()
   TextBoxChangeDate txtChargeDate
   'boldatechaged = False
End Sub

Private Sub txtChargeDate_GotFocus()
   SelTxtInCtrl txtChargeDate
End Sub

Private Sub txtChargeDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtClientSearch.SetFocus
    End If
   TextBoxKeyPrsDate txtChargeDate, KeyAscii
End Sub

Private Sub txtChargeDate_LostFocus()
   TextBoxFormatDate txtChargeDate
    If IsDate(txtChargeDate.text) = True Then
        txtPostingDate.text = txtChargeDate.text
   End If
   
   Dim rCount As Integer
   Dim szSQL As String
   Dim adoConn1 As New ADODB.Connection
   Dim szSelectedClient As String
   Dim szSelectedClientName As String
   Dim PurchaseLedgerControl As String
   For rCount = 1 To flxClients.Rows - 1
         If flxClients.TextMatrix(rCount, 0) = "X" Then
                szSelectedClient = flxClients.TextMatrix(rCount, 1)
                szSelectedClientName = flxClients.TextMatrix(rCount, 2)
                If szSelectedClient = "" Then Exit Sub
'                adoConn1.Open getConnectionString
'                If IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoConn1) = 0 Then
'                    MsgBox "The posting date cannot fall within a closed financial period for client :" & szSelectedClient, vbInformation, "Warning"
'                    adoConn1.Close
'                    'do not focus it to txt invoice date then you are in a loop
'                    Exit Sub
'                ElseIf IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoConn1) = 9 Then
'                    MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
'                    adoConn1.Close
'                    'FocusControl txtChargeDate
'                     'do not focus it to txt invoice date then you are in a loop
'                    Exit Sub
'                End If
'                adoConn1.Close
         End If
   Next
   
    
   
End Sub

Private Sub ClearLocks(adoconn As ADODB.Connection)
   adoconn.Execute "UPDATE Property SET RunningAutoDemand = '' WHERE RunningAutoDemand = '" & WS_Name & "';"
End Sub

Private Function IsDemandRunning(szProperties As String, adoconn As ADODB.Connection) As Boolean
   Dim iProp   As Integer
   Dim szSQL   As String
   Dim adoRst  As New ADODB.Recordset

   On Error GoTo ErrLocking

   szSQL = ""
   IsDemandRunning = False
   adoconn.BeginTrans

   adoconn.Execute "UPDATE Property SET RunningAutoDemand = '' WHERE RunningAutoDemand = '" & WS_Name & "';"

   szSQL = "SELECT * FROM Property WHERE ClientID = '" & _
                     flxClients.TextMatrix(flxClients.row, 1) & "';"
   adoRst.Open szSQL, adoconn, adOpenForwardOnly, adLockOptimistic

   While Not adoRst.EOF
      For iProp = 1 To flxProperties.Rows - 1
         If flxProperties.TextMatrix(iProp, 0) = "X" Then
            If adoRst.Fields.Item("PropertyID").Value = flxProperties.TextMatrix(iProp, 1) Then
               If adoRst.Fields.Item("RunningAutoDemand").Value <> "" And _
                     adoRst.Fields.Item("RunningAutoDemand").Value <> WS_Name Then
'                  szProperties = szProperties + ", " + adoRst.Fields.Item("PropertyName").Value
                  szProperties = adoRst.Fields.Item("PropertyName").Value + Chr(13) + szProperties
                  IsDemandRunning = True
               Else
                  adoRst.Fields.Item("RunningAutoDemand").Value = WS_Name
                  adoRst.Update
               End If
            End If
         End If
      Next iProp
      adoRst.MoveNext
   Wend

   adoconn.CommitTrans

   adoRst.Close
   Set adoRst = Nothing

   Exit Function

ErrLocking:
   IsDemandRunning = True
   adoconn.RollbackTrans
   Set adoRst = Nothing
End Function
Private Function Check_Fund_LRentCharge(adoConn1 As ADODB.Connection) As Boolean
    Dim rsLRentCharges As New ADODB.Recordset
    Dim rsLease As New ADODB.Recordset
    Dim szSQL As String
    Dim szSAGEID As String
    rsLRentCharges.Open "Select * from LRentCharges where RentChargeDept='' or isnull(RentChargeDept)", adoConn1, adOpenStatic, adLockReadOnly
    If Not rsLRentCharges.EOF Then
        szSQL = "Select * from LeaseDetails L where L.LeaseID='" & rsLRentCharges("LeaseID").Value & "'"
        rsLease.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
        If Not rsLease.EOF Then
            szSAGEID = rsLease("SageAccountNumber").Value
        End If
        rsLease.Close
        MsgBox "There is a 'Fund Name'  missing against Lease ID : '" & szSAGEID & "' . Please correct this before generating your demands", vbCritical, "Fund Name missing  "
        Check_Fund_LRentCharge = True
    End If
    rsLRentCharges.Close

End Function

Private Sub cmdGDPOk_Click()
   Dim i As Integer
   Dim rsDemandType As New ADODB.Recordset
   Dim rCount As Integer
   Dim szSelectedClient As String
   If txtChargeDate.text = "" Then
      ShowMsgInTaskBar "Please type the Issue date.", , "N"
      txtChargeDate.SetFocus
      Exit Sub
   End If

   If IsDate(txtPostingDate.text) = False Then
        ShowMsgInTaskBar "Please enter posting date.", , "N"
        txtPostingDate.SetFocus
        Exit Sub
   End If
   FocusControl cmdGDPCancel
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'##################################### Management Fee  PREVIEW   ####################
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If szCallingFrom = "ManagementFee Preview" Then
        Call GeneratePIPreview
        Picture1.Visible = False
        Picture1.Refresh
        Exit Sub
   End If
    If szCallingFrom = "ManagementFee" Then
        Call GeneratePI
        Picture1.Visible = False
        Picture1.Refresh
        Exit Sub
   End If



End Sub

Private Function checkPurINVWithNLPostingPILCA(adoconn As ADODB.Connection, PILCA As String, TransactionType As Integer) As Boolean
    'Written by anol 2020-05-11 issue 841   remember true =good ,false is BAD
    'This fucntion shall return false if any inconsistent data found after comparing tblPurInv and NLPosting
    'remember that in the following query I cannot use NLPost=true in tblPurInv Because I am Updating this field for all records after committing a transaction
    'So I am not using NLpost=true field so that it includes current/latest Pur inv for comparison
    Dim szSQL As String
    Dim rsInv As New ADODB.Recordset

    szSQL = " Select transaction_ref,*  From ((Select sum( abs(amount))  as Totalamount,transaction_ref,TRANSACTION_TYPE,clientID " & _
            " from NLPosting where TRANSACTION_TYPE=" & TransactionType & "  and Nominal_code='" & PILCA & "'  and Deleteflag=false group by" & _
            "  transaction_ref,TRANSACTION_TYPE,clientID) as X INNER JOIN (Select  sum(TOTAL_AMOUNT)as TAmount,slnumber,transactionType,CL_ID from" & _
            " tblpurinv where transactionType=" & TransactionType & " group by slnumber,transactionType,CL_ID) as" & _
            " Y ON  X.ClientID=Y.CL_ID AND X.transaction_ref=cstr(Y.slnumber) AND  X.TRANSACTION_TYPE=Y.transactionType) where Totalamount<>Tamount"
     rsInv.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
     If rsInv.EOF Then
        checkPurINVWithNLPostingPILCA = True ' that means there is not problematic data
     Else
        MsgBox "There are existing invoice(s) with reference ID: " & SQL2String(rsInv, 0) & " that do not match the Nominal Ledger. Please contact PCM Support.", vbInformation, "PI/PC: " & SQL2String(rsInv, 0)
         checkPurINVWithNLPostingPILCA = True
     End If

End Function

Private Function checksetupDONE(szProperty As String) As Boolean
        Dim adoconn As New ADODB.Connection
        Dim rsCharge As New ADODB.Recordset
        Dim szSQL  As String
        adoconn.Open getConnectionString
        szSQL = "SELECT agr.CHARGE_METHOD,agr.LastChargeDate, agr.TotalAmount,agr.Amount,agr.Fund, agr.NtDueDate,agr.FDD,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as Frequency  " & _
                                      "FROM tlbAgreement agr, ClientProAgr CPA,  ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
                                      "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
                                      "C.ID = agr.CHARGE_TYPE And " & _
                                      "CPA.PropertyID = '" & szProperty & "' "
        rsCharge.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If rsCharge.EOF Then
            checksetupDONE = True
        End If
        rsCharge.Close
        
        
        adoconn.Close
End Function
Private Function checkGlobalDataEntered(szProperty As String) As Boolean
        Dim adoconn As New ADODB.Connection
        Dim rsCharge As New ADODB.Recordset
        Dim szSQL  As String
        adoconn.Open getConnectionString
        szSQL = "SELECT *,YDueDate from globaldata where PropertyID = '" & szProperty & "' "
        rsCharge.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not rsCharge.EOF Then
            If IsNull(rsCharge("YDueDate").Value) = True Then
                checkGlobalDataEntered = True
            End If
        End If
        rsCharge.Close
        
        adoconn.Close
End Function
Private Function hasChargeTypeinlist(szProperty1 As String) As Boolean
    Dim iCount As Integer
    For iCount = 1 To flxDemandTypes.Rows - 1
         If flxDemandTypes.TextMatrix(iCount, 2) <> "" Then
                'If flxDemandTypes.TextMatrix(iCount, 0) = "X" Then
                     If flxDemandTypes.TextMatrix(iCount, 1) = szProperty1 Then
                        hasChargeTypeinlist = True
                     End If
                'End If
         End If
     Next
End Function
Private Sub GeneratePI()
        Dim descriptionAndDate As String
        Dim lSlNumber As Long
        Dim adoconn As New ADODB.Connection
        Dim adoConnTransactions As New ADODB.Connection
        Dim adoPIHeader As New ADODB.Recordset
        Dim adoPISplit As New ADODB.Recordset
        Dim szSQL As String
        Dim szMYID As String
        Dim szFundID As String
        Dim szSelectedPayableTypeID As String
        Dim szSelectedClient As String
        Dim szPropertySelection1 As String
        Dim szSQL1 As String
        Dim rsfixedMethod As New ADODB.Recordset
        Dim rsfixedMethodDetails As New ADODB.Recordset
        Dim j As Integer
        Dim percentageOramount As Double
        Dim dblGrandTotal As Double
        Dim dtNextDue As Date
        Dim dtFDD As Date
        Dim dtNDD As Date
        Dim dblFeqID As Integer
        Dim dblTotalAmount As Double
        Dim lngMgtFeeSL As Long
        Dim iClientCount As Integer
        Dim bControlACForPayable As Boolean
        Dim FinalControlACForPayable As String
        Dim iCountPI As Integer
        Dim strFundName As String
        Dim strStopDate As String
        Dim intNoOfDaysToSendMFB4Due As Integer
        Dim rsManagingAgent As New ADODB.Recordset
        Dim szManagingAgent() As String
        Dim iManagingAgentCount As Integer
        Dim szSQLManagingAgent As String
        Dim szTemp
        Dim iCount As Long
        Dim iCount1 As Long
        Dim iPropertyCount As Long
        Dim rCount As Long
        Dim dblFundId As Integer
        Dim i As Integer
        Dim lT_ID As Long
        Dim strSelectedFundID As String
        Dim szSQL5 As String
        Dim rstSet As New ADODB.Recordset
        Dim dblCapAmount As Double
        Dim dblDemandTypeId As Integer
        Dim rsGlobalData As New ADODB.Recordset
        Dim dblNoOfDaysToSendMFB4Due As Integer
        Dim strFromDate As Date
        Dim strToDate As Date
        
        Dim VAT_ID As String
        Dim VAT_CODE As String
        Dim VAT_RATE As Double
        Dim strLastChargeDate As Date
        Dim SQLforInsert As String
        'Dim rsGlobaldata As New ADODB.Recordset
        Dim dicWarningProp As New Dictionary
        Dim dicWarningAgreement As New Dictionary
        Dim dicWarningFinPeriod As New Dictionary
        Dim szPropertySelectionALL As String
        
        If txtChargeDate.text = "" Then
            MsgBox "Please enter invoice date", vbInformation, "Warning "
            FocusControl txtChargeDate
            Exit Sub
        End If
        strSelectedFundID = ""
        For iCount = 1 To flxDemandTypes.Rows - 1
            If flxDemandTypes.TextMatrix(iCount, 2) <> "" Then
                   iCount1 = iCount1 + 1
                   If flxDemandTypes.TextMatrix(iCount, 0) = "X" Then
                        strSelectedFundID = strSelectedFundID + "" + flxDemandTypes.TextMatrix(iCount, 2) + ","
                   End If
            End If
        Next
        If Len(strSelectedFundID) > 0 Then
                strSelectedFundID = Left(strSelectedFundID, Len(strSelectedFundID) - 1)
        End If

        For rCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(rCount, 0) = "X" Then
               szSelectedClient = flxClients.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
        For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    If szPropertySelectionALL = "" Then
                        szPropertySelectionALL = "'" + flxProperties.TextMatrix(rCount, 1) + "'"
                    Else
                        szPropertySelectionALL = szPropertySelectionALL + ",'" + flxProperties.TextMatrix(rCount, 1) & "'"
                    End If

                End If
        Next
        Debug.Print time & " 1 start"
        
        For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    szPropertySelection1 = flxProperties.TextMatrix(rCount, 1)
                    If checksetupDONE(szPropertySelection1) = True Then
                         dicWarningAgreement.Add rCount, szPropertySelection1
                         'MsgBox "Please enter a valid setup for the property: " & szPropertySelection1, vbInformation, "Client agreement Fees and charges setup"
                    End If
'                    If checkGlobalDataEntered(szPropertySelection1) = True Then
'                         MsgBox "Please enter a valid Client Global Data setup for the property: " & szPropertySelection1, vbInformation, "Client Global Data setup"
'                    End If
'                    If hasChargeTypeinlist(szPropertySelection1) = False Then
'                              MsgBox "Please setup a charge type for the selected property: " & szPropertySelection1, vbInformation, "Warning"
'                    End If
                End If
        Next
        Debug.Print time & " 2 end"
'        adoConn.Open getConnectionString

       
        If szPropertySelection1 = "" Then
            MsgBox "Please select a Property.", vbInformation, "Warning"
            Exit Sub
        End If
'        If iCount1 = 0 Then
'            MsgBox "Please setup a charge type for the selected property.", vbInformation, "Warning"
'            Exit Sub
'        End If
        If strSelectedFundID = "" Then
'            MsgBox "Please select a charge type.", vbInformation, "Warning"
            Exit Sub
        End If
        If chkAssignProperty.Value = 1 Then
            If MsgBox("Are you sure you wish to generate your management fees without assigning a property", vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
        
      'Test before write if data is fine
'      If checkPurINVWithNLPostingPILCA(adoConn, FinalControlACForPayable, "6") = True Then
'        MsgBox "Good data"
'        Exit Sub
'      End If
 '      ************************************Write tblPurInv **************************************
'            adoConn.Close
            Picture1.Visible = True 'loading msg
            Picture1.Refresh
            
            'Loop for Properties loop for each PI creation
            For iClientCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(iClientCount, 0) = "X" Then 'also check financial year is correct period and control account is there and it is set in a variable before looping
                        szSelectedClient = flxClients.TextMatrix(iClientCount, 1)
                        If szSelectedClient = "" Then Exit Sub
                        If adoconn.State = 1 Then
                            adoconn.Close
                        End If
                        adoconn.Open getConnectionString
                        If FinalControlACForPayable = "" Then
                                  FinalControlACForPayable = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", szSelectedClient)
                        End If
                            'if still control acount is empty generate a warning message and exit this sub procedure
                        If FinalControlACForPayable = "" Then
                                MsgBox "Control Account is not set for this Payable type for Client:" & szSelectedClient, vbInformation, "Warning"
                                GoTo EndOfAgreement
                                Exit Sub
                        End If
                        If IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoconn) = 0 Then
                            MsgBox "The posting date cannot fall within a closed financial period for the client :" & szSelectedClient, vbInformation, "Warning"
                            adoconn.Close
                            'FocusControl txtChargeDate
                            'Exit Sub
                            GoTo EndOfAgreement
                        ElseIf IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoconn) = 9 Then
                            'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                             dicWarningFinPeriod.Add iClientCount, szSelectedClient
                             adoconn.Close
                            'FocusControl txtChargeDate
                            'Exit Sub
                            GoTo EndOfAgreement
                        End If
                        adoconn.Close
            
            For iPropertyCount = 1 To flxProperties.Rows - 1
                    If iPropertyCount >= flxProperties.Rows Then Exit For
                    If flxProperties.TextMatrix(iPropertyCount, 0) = "X" And szSelectedClient = flxProperties.TextMatrix(iPropertyCount, 3) Then
                            Debug.Print iPropertyCount
                            szPropertySelection1 = flxProperties.TextMatrix(iPropertyCount, 1)
                            
                            If checkGlobalDataEntered(szPropertySelection1) = True Then
                                 'MsgBox "Please enter a valid Client Global Settings setup for the property: " & szPropertySelection1, vbInformation, "Client Global Data setup"
                                 dicWarningProp.Add iPropertyCount, szPropertySelection1
                                 GoTo EndOfAgreement
                            End If
                            If adoconn.State = 1 Then
                                adoconn.Close
                            End If
                            adoconn.Open getConnectionString
                            
                            
                            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
                            "FROM tlbAgreement agr, ClientProAgr CPA,  ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
                            "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
                            "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                            "CPA.PropertyID = '" & szPropertySelection1 & "' And F.FundID IN(" & strSelectedFundID & ")"
                            rsManagingAgent.Open szSQLManagingAgent, adoconn, adOpenDynamic, adLockOptimistic
                            szTemp = SQL2String(rsManagingAgent, 0)
                            rsManagingAgent.Close
                            If Len(szTemp) > 0 Then
                                  szManagingAgent = Split(szTemp, ",")
                            Else
                                 adoconn.Close
                                GoTo EndOfAgreement
                            End If
                            
                            
                            szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
                            rstSet.Open szSQL5, adoconn, adOpenStatic, adLockReadOnly
                            lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
                            rstSet.Close
                            Set rstSet = Nothing
                            adoconn.Close
                            
                            For iManagingAgentCount = 0 To UBound(szManagingAgent) 'this for shall end after the creation of the PI preview
                            
                            adoconn.Open getConnectionString
                            'adoConn.BeginTrans
                            
                           ' Exit Sub
                            
                            Dim rsCharge As New ADODB.Recordset
                            szSQL = "SELECT AGREEMENT_ID,agr.Capamount, agr.StopDate, CPA.agreementStartDate,CPA.agreementEndDate,agr.CHARGE_METHOD, " & _
                                "agr.LastChargeDate,agr.StopDate, agr.TotalAmount,agr.Amount,agr.Fund,F.FundName, agr.NtDueDate,agr.FDD,agr.Frequency as FrequencyID,agr.EachPeriod, " & _
                                "(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as FrequencyName,ManagingAgentID,NtDueDate  " & _
                               "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
                               "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
                               "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                               "CPA.PropertyID = '" & szPropertySelection1 & "' ANd F.FundID IN(" & strSelectedFundID & ") AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
                              
                              
                            '    szSQL = "SELECT agr.TotalAmount,agr.Fund,agr.Frequency ,agr.NtDueDate,agr.FDD " & _
                            '            "FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
                            '            "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
                            '            "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
                            '            "CPA.PropertyID = '" & szPropertySelection1 & "'"
                            rsCharge.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                            If rsCharge.EOF Then
                                rsCharge.Close
                                Set rsCharge = Nothing
                                adoconn.Close
                                
                                Set adoconn = Nothing
                                MsgBox "Please enter a valid setup for the property: " & szPropertySelection1, vbInformation, "Client agreement Fees and charges setup"
                                GoTo EndOfAgreement
                            End If
                            
                            strStopDate = rsCharge("StopDate").Value
                            'dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value
                            i = 1
                             dblGrandTotal = 0

                                
                         rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                            szPropertySelection1 & "' AND vatOptionEnabled=true", adoconn, adOpenStatic, adLockReadOnly
                         If Not rsGlobalData.EOF Then
                                  VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                  VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                  VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                         Else
                                  VAT_ID = -1
                                  VAT_RATE = 0
                                  VAT_CODE = ""
                         End If
                         rsGlobalData.Close
                         Set rsGlobalData = Nothing
                         
                         
                             Dim rsGlobalData1 As New ADODB.Recordset
                             Dim bolVatOptionEnabled As Boolean
                             Dim bolOptedTotax As String
                             Dim strManagingAgentID As String
                             rsGlobalData1.Open "Select vatOptionEnabled from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                                szPropertySelection1 & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                
                            If Not rsGlobalData1.EOF Then
                                    bolVatOptionEnabled = IIf(IsNull(rsGlobalData1("vatOptionEnabled").Value), False, rsGlobalData1("vatOptionEnabled").Value)
                                    strManagingAgentID = rsCharge("ManagingAgentID").Value
                            End If
                            rsGlobalData1.Close
                            rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoconn, adOpenStatic, adLockReadOnly
                            If Not rsGlobalData1.EOF Then
                                    bolOptedTotax = rsGlobalData1("optedTotax").Value
                            Else
                                    bolOptedTotax = False
                            End If
                            rsGlobalData1.Close
                      
                  ' Debug.Print 0 / 0
                        'days before B4Due
                             'Exit Sub
                    'days before B4Due end
                    Dim icounterTemp As Integer
                    icounterTemp = 1
                            While Not rsCharge.EOF
                                  szMYID = UniqueID()
'                                  adoConnTransactions.Open getConnectionString
'                                  adoConnTransactions.BeginTrans
                                  dblTotalAmount = rsCharge.Fields.Item("TotalAmount").Value
                                  dblCapAmount = rsCharge.Fields.Item("CapAmount").Value
                                  dblFundId = rsCharge.Fields.Item("Fund").Value
                                  strFundName = rsCharge.Fields.Item("FundName").Value
                                  If Not IsNull(rsCharge.Fields.Item("NtDueDate").Value) Then
                                      dtNextDue = rsCharge.Fields.Item("NtDueDate").Value
                                  End If
                                  If Not IsNull(rsCharge.Fields.Item("FDD").Value) Then
                                      dtFDD = rsCharge.Fields.Item("FDD").Value
                                  End If
                                  If Not IsNull(rsCharge.Fields.Item("FrequencyID").Value) Then
                                        If rsCharge.Fields.Item("FrequencyID").Value <> "" Then
                                            dblFeqID = rsCharge.Fields.Item("FrequencyID").Value
                                        End If
                                  End If
                           

                                 strLastChargeDate = IIf(IsNull(rsCharge("LastChargeDate").Value), "", rsCharge("LastChargeDate").Value)
                                 If IsDate(strLastChargeDate) = False Then
                                        rsCharge.Close
'                                        adoConn.RollbackTrans
                                        adoconn.Close
                                        MsgBox "Please enter a last charge date for Property:" & szPropertySelection1, vbInformation, "Warning"
                                        'GoTo EndOfAgreement
                                        GoTo EndOfOneManagingAgentforOneAgreement
                                 End If
                                 If strStopDate = "" Then
                                        'Exit Sub
                                    Else
                                        If IsDate(strStopDate) Then
                                                If DateDiff("d", strStopDate, txtChargeDate.text) >= 0 Then
'                                                    MsgBox "It is not possible to generate fees after the stop date.:" & szPropertySelection1, vbInformation, "Warning"
                                                End If
                                        Else
                                                 MsgBox "Stop date date format is not correct.:" & szPropertySelection1, vbInformation, "Warning"
                                        End If
                                        GoTo EndOfChargeType
                                    End If
                                 rsGlobalData.Open "Select NoOfDaysToSendMFB4Due from globaldata where PropertyID='" & szPropertySelection1 & "'", adoconn, adOpenStatic, adLockReadOnly
                                 If Not rsGlobalData.EOF Then
                                        dblNoOfDaysToSendMFB4Due = IIf(IsNull(rsGlobalData!NoOfDaysToSendMFB4Due), 0, rsGlobalData!NoOfDaysToSendMFB4Due)
                                  End If
                                 rsGlobalData.Close

                                 If DateDiff("d", Date, rsCharge("NtDueDate").Value) > dblNoOfDaysToSendMFB4Due Then
                                             GoTo EndOfChargeType
                                  End If
                                   lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
'                            szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
'                            adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                            lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
'                            adoPIHeader.Close
'                             j = 1
'
'                             szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
'                             rstSet.Open szSQL5, adoConn, adOpenStatic, adLockReadOnly
'                             lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
'                             rstSet.Close
'                             Set rstSet = Nothing
                             If lT_ID = 495 Then
                                Debug.Print ""
                             End If
                                  adoConnTransactions.Open getConnectionString
                                  adoConnTransactions.BeginTrans
                                  Debug.Print "icounterTemp:" & icounterTemp
                                  icounterTemp = icounterTemp + 1
                                    'moved trxID making code here to sovle that error
                                    
                                szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
                                adoPIHeader.Open szSQL, adoConnTransactions, adOpenStatic, adLockReadOnly
                                lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
                                adoPIHeader.Close
                                 j = 1
                                 
                                 szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
                                 rstSet.Open szSQL5, adoConnTransactions, adOpenStatic, adLockReadOnly
                                 lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
                                 rstSet.Close
                                 Set rstSet = Nothing
                                    
                                     
                                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then 'Fixed Basis  'when working on the fixed procedure only 1 line of setup is done then
                                                                If rsCharge.Fields.Item("agreementEndDate").Value < rsCharge.Fields.Item("FDD").Value Then
                                                                        rsCharge.Close
                                                                        Set rsCharge = Nothing
                                                                        MsgBox "agreement End Date is greatar than following due date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                                        adoConnTransactions.RollbackTrans
                                                                        GoTo EndOfChargeType
                                                                End If
                                                                If DateDiff("d", txtChargeDate.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                                        rsCharge.Close
                                                                        Set rsCharge = Nothing
                                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                                         adoConnTransactions.RollbackTrans
                                                                        GoTo EndOfChargeType
                                                                End If
                                                                rsGlobalData.Open "Select NoOfDaysToSendMFB4Due,* from GlobalData where PropertyID='" & szPropertySelection1 & "'", adoconn, adOpenStatic, adLockReadOnly
                                                                If Not rsGlobalData.EOF Then
                                                                    If Not IsNull(rsGlobalData("NoOfDaysToSendMFB4Due").Value) Then
                                                                        intNoOfDaysToSendMFB4Due = rsGlobalData("NoOfDaysToSendMFB4Due").Value
                                                                    End If
                                                                End If
                                                                rsGlobalData.Clone
                                                                Set rsGlobalData = Nothing
                                                                If DateDiff("d", Date, rsCharge("NtDueDate").Value) > intNoOfDaysToSendMFB4Due Then
                                                                     adoConnTransactions.RollbackTrans
                                                                    GoTo EndOfChargeType
                                                                End If
                                                                
                                                                            dblTotalAmount = rsCharge.Fields.Item("EachPeriod").Value
                                                                            dblFundId = rsCharge.Fields.Item("Fund").Value
                                                                            If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                  dblTotalAmount = dblTotalAmount * (percentageOramount / 100)
                                                                                  dblTotalAmount = Round(dblTotalAmount, 2)
                                                                            End If
                                                                            
                                                                            If dblCapAmount > 0 Then
                                                                            
                                                                                   If dblTotalAmount > dblCapAmount Then
                                                                                        dblTotalAmount = dblCapAmount
                                                                                   End If
                                                                            End If
                                                                            szSQL = "SELECT * FROM tblPurInvSRec"
                                                                            adoPISplit.Open szSQL, adoConnTransactions, adOpenDynamic, adLockOptimistic
                                                                            'Add New Records. At least there is only one split line
                                                                               With adoPISplit
                                                                                   .AddNew
                                                                                   .Fields.Item("MY_ID").Value = UniqueID()
                                                                                   .Fields.Item("ParentID").Value = szMYID
                                                                                   .Fields.Item("TRAN_ID").Value = j
                                                                                   If chkAssignProperty.Value = 0 Then
                                                                                        .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                   Else
                                                                                        .Fields.Item("TRANS").Value = ""
                                                                                   End If
                                                                                    Dim dtNDD1 As Date
                                                                                    Dim dtFDD1 As Date
                                                                                    txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                                                    dtNDD1 = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                                    txtComparenextDueDate1 = DateAdd("d", 1, dtNDD1)
                                                                                    dtFDD1 = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                                    strFromDate = Format(dtNextDue, "dd/MM/yyyy")
                                                                                    strToDate = Format(DateAdd("d", -1, dtNDD1), "dd/MM/yyyy")
                                                                
                                                                
                                                                                   .Fields.Item("UNIT_ID").Value = ""
                                                                                   .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                   .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                  ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                   .Fields.Item("RecoverablePt").Value = 0
                                                                                   '.Fields.Item("description").Value = "Management Fee for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                   .Fields.Item("description").Value = "Management Fee for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
'                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
'                                                                                   .Fields.Item("TAX_CODE").Value = VAT_CODE
'                                                                                   .Fields.Item("VAT").Value = Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
'                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
'                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
'Exit Sub
                                                                                     If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                                            .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                            .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                            .Fields.Item("VAT").Value = dblTotalAmount * Round((VAT_RATE / 100), 2) 'VAT_RATE
                                                                                            .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                             dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                       ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then
                                                                                                                                                                                            .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                            .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                            .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                             dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                      ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then
                                                                                            rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                                               strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                                            If Not rsGlobalData.EOF Then
                                                                                                     VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                                                     VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                                                     VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                                            Else
                                                                                                     VAT_ID = -1
                                                                                                     VAT_RATE = 0
                                                                                                     VAT_CODE = ""
                                                                                            End If
                                                                                            rsGlobalData.Close

                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                      ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
'                                                                                            .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
'                                                                                            .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
'                                                                                            .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
'                                                                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
'                                                                                             dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                                 .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                       
                                                                                       End If
                                                                                   .Update
                                                                               End With
                                                                            adoPISplit.Close
                                                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                                                  'Exit Sub
                                                                                  
                                                                           '*************************************Write tlbPaymentSplit **************************************
                                                                          ' adoPIHeader.Close
                                                                           Set adoPIHeader = Nothing
                                                                              szSQL = "SELECT * FROM tlbPaymentSplit;"
                                                                           adoPISplit.Open szSQL, adoConnTransactions, adOpenDynamic, adLockPessimistic
                                                                        'Add New Records. At least there is one split line.
                                                                                 With adoPISplit
                                                                                    .AddNew
                                                                                    .Fields.Item("TransactionID").Value = UniqueID()
                                                                                    .Fields.Item("PayHeader").Value = lT_ID
                                                                                    .Fields.Item("FundID").Value = dblFundId
                                                                                    .Fields.Item("Amount").Value = dblTotalAmount
                                                                                    .Fields.Item("OSAmount").Value = dblTotalAmount
                                                                                    .Fields.Item("SplitID").Value = j
                                                                                    .Fields.Item("DueDate").Value = Format(txtChargeDate.text, "DD/MMMM/YYYY")
                                                                                    .Fields.Item("Description").Value = "Management Fee"
                                                                                    '.Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                    '.Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                   .Fields.Item("NOMINAL_CODE").Value = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", szSelectedClient)
                                                                                    .Fields.Item("TRANS").Value = "" 'Put property ID  Later
                                                                                    .Fields.Item("UNIT_ID").Value = "" '
                                                                                    .Fields.Item("ScheduleID").Value = Null
                                                                                    .Update
                                                                                 End With
                                                                                 adoPISplit.Close
                                                                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then
                                                                                     'Updating FDDand next due date
                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                                                        dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                                        
                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
                                                                                        dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                                        
                                                                                       strFromDate = Format(dtNextDue, "dd/MM/yyyy")
                                                                                       strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy") ' Format(DateAdd("d", -1, dtFDD), "dd/MMM/yyyy")
                                                                                      szSQL = "Update tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
                                                                                     " Set NtDueDate=#" & Format(dtNDD, "dd/MMM/yyyy") & "# ,FDD=#" & Format(dtFDD, "dd/MMM/yyyy") & "#,lastchargeDate=#" & Format(txtChargeDate.text, "dd/MMM/yyyy") & "# " & _
                                                                                     "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
                                                                                     "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                                                                                     "CPA.PropertyID = '" & szPropertySelection1 & "' and AGREEMENT_ID=" & rsCharge.Fields.Item("AGREEMENT_ID").Value & ";"
                                                                                     adoConnTransactions.Execute szSQL
                                                                                     
                                                                                     'adoConnTransactions.Execute szSQL
                                                                                End If
'                                                  SQLforInsert = "'" & szMYID & "','" & Format(txtChargeDate.text, "dd MMM yyyy") & "'," & szPropertySelection1 & "','Fixed Basis'," & dblTotalAmount & ""
'                                                  adoconn.Execute "Insert into ManagementFeePreview(PI_ActualID,ChargeDate,PropertyID,ChargingMethod," & _
'                                                "ReceiptAmount)" & _
'                                                SQLforInsert
                                                    SQLforInsert = "'" & szMYID & "','" & Format(txtChargeDate.text, "dd MMM yyyy") & "','" & szPropertySelection1 & "','Fixed Basis',1,1," & dblFundId & "," & dblTotalAmount & ")"
                                                   adoconn.Execute "Insert into ManagementFee(PI_ActualID,ChargeDate,PropertyID,ChargingMethod,ReceiptSplitID,ReceiptType,FundID," & _
                                                "MgtFeeAmtTotal) values (" & _
                                                SQLforInsert
                                                
                                                
                                                End If 'end of for Receipt BASIS
                                         If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then 'received basis
                                                If rsCharge.Fields.Item("agreementEndDate").Value < rsCharge.Fields.Item("FDD").Value Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        adoConnTransactions.RollbackTrans
                                                        MsgBox "agreement End Date is greatar than following due date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                      
                                                End If
                                                If DateDiff("d", txtChargeDate.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        adoConnTransactions.RollbackTrans
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                        
                                                End If

                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
                                                    "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                                    "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                    "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                                    szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""

                
'                                                szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                                "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where R1.DemandRef=DS.DemandID and AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND AL.deleteflag=false AND " & _
'                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
'                                                "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                szPropertySelection1 & "' and RS.ISMGTFEES=false  AND Rs.FundID=" & dblFundId & " "
                                                'AND  isnull(RS.ClientStatementID)
                                                      'need to consider the selected property in where clause
'                                                                      'Here I am comparing . need to rem this peice of SQL
'                                                    szSQL = "Update  tlbReceipt R,tlbReceiptsplit RS,Units,tlbReceipt R1, rptTransactionsSPlit AL, DemandSplitRecords DS,Units U SET ISMGTFEES=true,R1.PIREFMGTFEE='" & _
'                                                    szMYID & "',R.Chargedate=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "#  where " & _
'                                                    "R1.DemandRef=DS.DemandID AND AL.TOTRAN=R1.TransactionID AND AL.TransactionID= RS.RptTransactionsIDSplit AND AL.deleteflag=false AND RS.SPLITID=DS.SPLITID " & _
'                                                    "AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "#  " & _
'                                                   "And R.Type in (3,4,23) AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                    szPropertySelection1 & "'  AND ISMGTFEES=false and DS.TypeOfDemand in (" & rsCharge("DEMAND_TYPE").Value & ")"
                                    

                                                    '  rsfixedMethod.Close
                                                    If rsfixedMethod.State = 1 Then
                                                        rsfixedMethod.Close
                                                    End If
                                            rsfixedMethod.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
                            'Here type 3 is for recipt type . I have not written for the credit yet need to understand the principle
                                                If IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value) = 0 Then
                                                    rsfixedMethod.Close
                                                     adoConnTransactions.RollbackTrans
                                                    GoTo EndOfChargeType
                                                End If
                                            percentageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                            
                                            '      ************************************Write tblPurInvSRec **************************************
'
                                            'You cannot write tlbReceiptSplitID equal to DemandRecordsSplitID . From Many of the reasons one is your tlbReceiptSplit
                                            'one demand can have many demand type in it. so if you create a relationship via main demand table that will be wrong as well
                                            'What does the following SQL Means need a proper explanation/ documentation
                                            ' Selecting all the reciepts from the tlbReceiptDetails table//Here why I am using demand table???? filter for demand type in other side of allocation
                                            ' In the following SQL I dont need DemandSplitRecords
                                              szSQL = "Select  (SWITCH(R.TYPE =3,AL.NetAmount,R.TYPE =4,AL.NetAmount,R.TYPE =23,-AL.NetAmount)) as Amt,rs.FundID,(AL.NetAmount)as NETAMT from tlbReceipt R," & _
                                                "tlbReceiptsplit RS, rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID  AND U.PropertyID='" & _
                                                 szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & "  "
                                            'need to insert in Managementfee table here     'Receipt BASIS
                                            SQLforInsert = "Select " & szMYID & " as MYID , '" & Format(txtChargeDate.text, "dd MMM yyyy") & _
                                                "' as ChargeDate,R.SlNumber, R.Type,R.SageAccountNumber,R.Ref,U.PropertyID,RS.FundID,RptAmtType,ExtRef,Rdate,R.TransactionID,RS.SplitID, " & _
                                                "(SWITCH(R.TYPE =3,R.Amount,R.TYPE =4,R.Amount,R.TYPE =23,-R.Amount)) as Amt from tlbReceipt R," & _
                                                "tlbReceiptsplit RS, rptTransactionsSplit AL, Units U where AL.deleteflag=false AND " & _
                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID  AND U.PropertyID='" & _
                                                 szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                 
                                            adoconn.Execute "Insert into ManagementFee(PI_ActualID,ChargeDate,SRSlNumber,ReceiptType,SageAccountNumber,ReceiptTypeDescription," & _
                                                "PropertyID,FundID,RptAmtType,ExtRef,ReceiptDate,ReceiptTransactionID,ReceiptSplitID," & _
                                                "ReceiptAmount)" & _
                                                SQLforInsert
                                                ',AgrPercentage,MgtFeeAmt,VATPercentage,VAT ,MgtFeeAmtTotal this things are null for now. need to update them in the next Update
                                            'RS.ChargeDateS,AgrPercentage,MgtFeeAmt,VATPercentage,VAT ,MgtFeeAmtTotal shall update later with a in where clause
                                            ' Update this ManagementFee on eache loop of charging
                                            
                                            'AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "#   'Receipt BASIS
                                            If rsfixedMethodDetails.State = 1 Then
                                                rsfixedMethodDetails.Close
                                            End If
                                            Dim rsFromandToDate As New ADODB.Recordset
                                            Dim szSQLFrom As String
                                            
                                            'modified on 20211103    'Receipt BASIS
'                                              szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX  from tlbReceipt R,tlbReceiptsplit RS,tlbReceipt R1, " & _
'                                                "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where  AL.TOTRAN=R1.TransactionID AND RS.SPLITID=DS.SPLITID AND R1.DemandRef=DS.demandID  " & _
'                                                "AND AL.deleteflag=false AND AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
'                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                szPropertySelection1 & "' and RS.ISMGTFEES=false  AND Rs.FundID=" & dblFundId & ""
      'SQL Modification done on 07-08-2023
                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from  tlbReceipt R, tlbReceipt R1,tlbReceiptsplit RS,  " & _
                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U,DemandRecords DR where  DS.DemandID=DR.DemandID and DS.DemandID= R1.DemandRef AND R1.transactionID=AL.ToTran AND AL.deleteflag=false AND " & _
                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=DR.UnitNumber AND U.PropertyID='" & _
                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                        
                                                rsFromandToDate.Open szSQLFrom, adoconn, adOpenStatic, adLockReadOnly
                                                ' AL.TOTRAN=R1.TransactionID means connecting to SI Side' for type 23 at sales side->  demand receipt  min max date will not be present. so remove 23
                                                'AND R.RDate>#" & Format(strLastChargeDate, "dd MMM yyyy") & "#   'Receipt BASIS
                                                'AND  isnull(RS.ClientStatementID)
                                                 If Not rsFromandToDate.EOF Then
                                                        If IsNull(rsFromandToDate("DateFromMin").Value) Then
'                                                             strFromDate = Null
'                                                             strToDate = ""
                                                        Else
                                                            strFromDate = Format(rsFromandToDate("DateFromMin").Value, "dd/MM/yyyy")
                                                            strToDate = Format(rsFromandToDate("DateTOMAX").Value, "dd/MM/yyyy")
                                                         End If
                                                 Else
                                                         strFromDate = Null
                                                         strToDate = Null
                                                 End If
                                                 rsFromandToDate.Close
                                            Dim dblOriginalAmount As Double
                                            j = 1
                                            rsfixedMethodDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                             While Not rsfixedMethodDetails.EOF  'we rare using whiles because itr
                                                            'dblTotalAmount = rsfixedMethodDetails.Fields.Item("Amt").Value  'Receipt BASIS
                                                    dblTotalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Amt").Value), 0, rsfixedMethodDetails.Fields.Item("Amt").Value)
                                                    dblOriginalAmount = rsfixedMethodDetails.Fields.Item("NETAMT").Value
        '                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                    dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                    If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                          dblTotalAmount = dblTotalAmount * (percentageOramount / 100)
                                                           dblTotalAmount = Round(dblTotalAmount, 2)
                                                    End If
                                                    If dblTotalAmount <= 0 Then
                                                        adoConnTransactions.RollbackTrans
                                                        GoTo EndOfChargeType
                                                    End If
                                                    'dblGrandTotal = 50
                                                    If dblCapAmount > 0 Then
                                                           'make a condition in the split as well so that amount doesnot exeed cap amount  'Receipt BASIS
                                                           If dblTotalAmount > dblCapAmount Then
                                                                dblTotalAmount = dblCapAmount
                                                           End If
                                                    End If
                                                    szSQL = "SELECT * FROM tblPurInvSRec"
                                                    adoPISplit.Open szSQL, adoConnTransactions, adOpenDynamic, adLockOptimistic
                                                    'Add New Records. At least there is only one split line
                                                       With adoPISplit
                                                           .AddNew
                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                           .Fields.Item("ParentID").Value = szMYID
                                                           .Fields.Item("TRAN_ID").Value = j
                                                          
                                                           If chkAssignProperty.Value = 0 Then
                                                                .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here  'Receipt BASIS
                                                           Else
                                                                .Fields.Item("TRANS").Value = ""
                                                           End If
                                                           .Fields.Item("UNIT_ID").Value = ""
                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                           .Fields.Item("RecoverablePt").Value = 0
                                                          ' .Fields.Item("description").Value = "Management Fee for " & strFundName & " " & DateAdd("d", 1, CDate(strLastChargeDate)) & " - " & txtChargeDate.text & ""
                                                           .Fields.Item("description").Value = "Management Fee for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           descriptionAndDate = "Management Fee for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           '.Fields.Item("NominalCode").Value = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", szSelectedClient)
'                                                                                   .Fields.Item("TAX_CODE").Value = VAT_CODE
'                                                                                   .Fields.Item("VAT").Value = Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
'                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
'                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value    'Receipt BASIS
                                                             If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                     'dblTotalAmount = Round(dblTotalAmount * (100 / (100 + VAT_RATE)), 2)
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                    .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                     adoconn.Execute "Update ManagementFee set AgrPercentage=" & percentageOramount & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                                     ' 'RS.ChargeDateS,AgrPercentage,MgtFeeAmt,VATPercentage,VAT ,MgtFeeAmtTotal shall update later with a in where clause
                                                             ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then
                                                                     ' dblTotalAmount = Round(dblTotalAmount * (100 / (100 + VAT_RATE)), 2)
                                                                     .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                     .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                     .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE    'Receipt BASIS
                                                                     .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                      dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                      adoconn.Execute "Update ManagementFee set AgrPercentage=" & percentageOramount & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                            ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then
                                                                  rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                     strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                  If Not rsGlobalData.EOF Then
                                                                           VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                           VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                           VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                  Else
                                                                           VAT_ID = -1
                                                                           VAT_RATE = 0
                                                                           VAT_CODE = ""
                                                                  End If
                                                                   rsGlobalData.Close
                                                                  .Fields.Item("NET_AMOUNT").Value = dblTotalAmount                                 ''Receipt BASIS
                                                                  .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                  .Fields.Item("VAT").Value = 0 'blTotalAmount * (VAT_RATE / 100) ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                  .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                  .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                   dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                   adoconn.Execute "Update ManagementFee set AgrPercentage=" & percentageOramount & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                            ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                  .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                  .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                  .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                  .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                   dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                   adoconn.Execute "Update ManagementFee set AgrPercentage=" & percentageOramount & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                            End If
                                                            adoconn.Execute "Update ManagementFee set MgtFeeAmt=round(AgrPercentage*ReceiptAmount/100,3),ChargingMethod='Receipt Basis' where  ChargingMethod is null AND MgtFeeAmtTotal is null"
                                                            adoconn.Execute "Update ManagementFee set MgtFeeAmtTotal=round(MgtFeeAmt*(1+VATPercentage/100),3),VAT=round(MgtFeeAmt*(VATPercentage/100),3),ReceiptAmount=" & dblOriginalAmount & " where ChargingMethod='Receipt Basis' AND MgtFeeAmtTotal is null"
                                                           .Update
                                                       End With
                                                            adoPISplit.Close
                                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                           '*************************************Write tlbPaymentSplit **************************************    'Receipt BASIS
                                                          
                                                           Set adoPIHeader = Nothing
                                                              szSQL = "SELECT * FROM tlbPaymentSplit;"
                                                           adoPISplit.Open szSQL, adoConnTransactions, adOpenDynamic, adLockPessimistic
                                                        'Add New Records. At least there is one split line.
                                                          With adoPISplit
                                                             .AddNew
                                                             .Fields.Item("TransactionID").Value = UniqueID()
                                                             .Fields.Item("PayHeader").Value = lT_ID
                                                             .Fields.Item("FundID").Value = dblFundId
                                                             .Fields.Item("Amount").Value = dblTotalAmount
                                                             .Fields.Item("OSAmount").Value = dblTotalAmount
                                                             .Fields.Item("SplitID").Value = j
                                                              j = j + 1
                                                             .Fields.Item("DueDate").Value = Format(txtChargeDate.text, "DD/MMMM/YYYY")
                                                             'Modified by anol 2021-10-08
                                                             .Fields.Item("Description").Value = "Management Fee " & percentageOramount & "% of £" & Round(dblOriginalAmount, 2)
                                                             '.Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                             .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                             .Fields.Item("TRANS").Value = "" 'Put property ID  Later
                                                             .Fields.Item("UNIT_ID").Value = "" '
                                                             .Fields.Item("ScheduleID").Value = Null
                                                             .Update
                                                          End With
                                                          adoPISplit.Close
                                                         If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then
                                                              'Updating FDD and next due date
                                                               txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                               dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                 
                                                               szSQL = "Update tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
                                                              " Set NtDueDate=#" & Format(dtNextDue, "dd/MMM/yyyy") & "# ,FDD=#" & Format(dtFDD, "dd/MMM/yyyy") & "# " & _
                                                              "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
                                                              "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                                                              "CPA.PropertyID = '" & szPropertySelection1 & "';"
                                                              adoConnTransactions.Execute szSQL
                                                         End If
                                                         rsfixedMethodDetails.MoveNext
                                                    Wend
                                                    'NO need to consider fund when you update main table flagging
                                                    rsfixedMethod.Close
                                              End If 'end of for Receipt BASIS
'
                                                    'ISMGTFEES=true  means you have accounted management fees for the receipt and you are markting them as true
                                                   
                                                   
                                                    'SPLITID is not updated for  RS.SPLITID=DS.SPLITID found 2022-11-25 fix it
                                                    ' I need to write demand type in the tlbReceiptSplit while we are allocating receipts then we can use demandtypes Filter in where clause for
                                                    ' producing this management Fees
                                                    'RS split ID need not to be equal to DS split ID
                                                     'Mark 1
                                                    szSQL = "Update  tlbReceipt R,tlbReceiptsplit RS, rptTransactionsSplit AL,Units U SET ISMGTFEES=true,RS.PIREFMGTFEES='" & _
                                                    szMYID & "',RS.ChargedateS=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "#  where " & _
                                                    "AL.TransactionID= RS.RptTransactionsIDSplit AND AL.deleteflag=false  " & _
                                                    "AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "#  " & _
                                                   "And R.Type in (3,4,23) AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                                    szPropertySelection1 & "'  AND ISMGTFEES=false and RS.FundID in (" & dblFundId & ")"
                                                    adoConnTransactions.Execute szSQL
                                                    'AND  isnull(RS.ClientStatementID)
'                                                    Now question is AND  isnull(RS.ClientStatementID) should I use for updating or not ISMGTFEES=true
'                                                    now question is do you calculate MGT fee before CS generation
'                                                    my ans is before cs generation so you always update ismgtfee where isnull(RS.ClientStatementID)

                                                    szSQL = "Update tlbAgreement agr Set lastchargeDate=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                        "where AGREEMENT_ID=" & rsCharge.Fields.Item("AGREEMENT_ID").Value & ";"
                                                        adoConnTransactions.Execute szSQL
                                                                
                                                    dblTotalAmount = dblGrandTotal
                                                    If dblTotalAmount = 0 Then
                                                           adoConnTransactions.RollbackTrans
                                                           adoConnTransactions.Close
                                                           'GoTo EndOfAgreement
                                                           GoTo EndOfOneManagingAgentforOneAgreement
                                                    End If
                                                 
                                                  szSQL = "SELECT * FROM tblPurInv"
                                                  lSlNumber = SlNumber("PI", "tblPurInv", adoConnTransactions)
                                                  'adoConn.Execute "Update ManagementFee set SRslNUmber=" & lSlNumber & " where SRslNUmber is null"
                                                  'this is not getting a new sl number I need to increase it
                                                  With adoPIHeader
                                                    .Open szSQL, adoConnTransactions, adOpenDynamic, adLockPessimistic
                                                    .AddNew
                                                    .Fields.Item("MY_ID").Value = szMYID
                                                    .Fields.Item("CreatedBy").Value = User
                                                    .Fields.Item("CreatedDate").Value = Now
                                                    .Fields.Item("SlNumber").Value = lSlNumber
                                                    .Fields.Item("SUPP_AC").Value = Trim(szManagingAgent(iManagingAgentCount)) 'szSelectedClient
                                                    .Fields.Item("TRAN_DATE").Value = Format(txtChargeDate.text, "dd MMM YYYY")
                                                    .Fields.Item("TransactionType").Value = 6
                                                    .Fields.Item("INV_NO").Value = szPropertySelection1 + "-" + "MFee" + "-" + CStr(lngMgtFeeSL)
                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                    .Fields.Item("History").Value = False
                                                    .Fields.Item("TrfPayment").Value = False
                                                    .Fields.Item("PropertyID").Value = szPropertySelection1
                                                    .Fields.Item("CL_ID").Value = szSelectedClient
                                                    .Fields.Item("NLPost").Value = False
                                                    .Fields.Item("DueDate").Value = Format(txtChargeDate.text, "DD MMM YYYY")
                                                    .Fields.Item("PostingDate").Value = Format(txtPostingDate.text, "DD MMM YYYY")
                                                    .Fields.Item("isManagementFee").Value = True
                                                    .Fields.Item("ReportFromDate").Value = strFromDate
                                                    .Fields.Item("ReportToDate").Value = strToDate
                                                    .Fields.Item("LastModifiedBy").Value = User
                                                    .Fields.Item("LastModifiedDate").Value = Now
                                                    '.Fields.Item("DescriptionANDDates").Value = "Management Fee for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                    .Update
                                                  End With
                                                  adoPIHeader.Close
                                                  'lSlNumber = lSlNumber + 1
                                                   '*************************************Write tlbPayment **************************************
                                                   szSQL = "SELECT * FROM tlbPayment where 1=2"
                                                   With adoPIHeader
                                                      .Open szSQL, adoConnTransactions, adOpenDynamic, adLockOptimistic 'Add New Mode
                                                      .AddNew
                                                      !TransactionID = lT_ID
                                                      !Type = 6  'PP - Purchase Invoice, look in the tlbTransactionType
                                                      !SageAccountNumber = Trim(szManagingAgent(iManagingAgentCount)) ' szSelectedClient
                                                      !Pi = szMYID
                                                      !PDate = Format(txtChargeDate.text, "DD/MMMM/YYYY")
                                                      !dDate = Format(txtChargeDate.text, "DD/MMMM/YYYY")
                                                      !ref = "Management Fee"
                                                      !ExtRef = "Management Fee"
                                                      !amount = dblTotalAmount
                                                      !OSAmount = !amount
                                                      !PaymentView = True
                                                      !Details = "Management Fee"
                                                      !unitid = szPropertySelection1
                                                      '!nominalCode = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", szSelectedClient)
                                                      !SlNumber = lSlNumber
                                                      !fundID = dblFundId
                                        '              !AdjTag = IIf(bAdjustment, "Y", "N")
                                                      !Recoverable = 0
                                                      !postingDate = Format(txtPostingDate.text, "DD/MMMM/YYYY")
                                                      !ClientID = szSelectedClient
                                                      .Update
                                                      .Close
                                                 End With
                                                 lSlNumber = lSlNumber + 1
''                                                 rsCharge.Close
'                                                 Set rsCharge = Nothing
                                                 dblGrandTotal = 0
                                                 Dim szTran2Fix As String
                                                 Dim postResult As Boolean
                                                 If PI_Check(adoConnTransactions, szTran2Fix) = False Then
                                                      adoConnTransactions.RollbackTrans
                                                               ' adoConnTransactions.CommitTrans
                                                      adoConnTransactions.Close
                                                      MsgBox "An error occurred while saving, transaction rollbacked. Transactions: " & szTran2Fix, vbInformation, "Transaction rollbacked"
                                                      Exit Sub
                                                 Else
                                                      postResult = Export_PInPC_2_NL_ForAGENT(adoConnTransactions)
                                                      If postResult = False Then 'this part modified by anol 2021-01-13
                                                                 adoConnTransactions.RollbackTrans
                                                                 adoConnTransactions.Close
                                                                 MsgBox "There was a problem saving this transaction. It has therefore been rolled back", vbInformation, "Transaction rolled back"
                                                      Else
                                                                 adoConnTransactions.CommitTrans
                                                                 szSQL = "UPDATE tblPurInv AS P, tblPurInvSRec AS S " & _
                                                                     "SET P.NLPost = TRUE " & _
                                                                     "WHERE  P.MY_ID = S.ParentID AND NOT P.NLPost AND " & _
                                                                         "(P.TransactionType = 6 OR P.TransactionType = 7);"
                                                                 adoConnTransactions.Execute szSQL
                                                                 iCountPI = iCountPI + 1
                                                                 'MsgBox "PI has been saved", vbInformation, "Saved"
                                                                ' Unload Me
                                                                
                                                      End If
                                                End If
EndOfChargeType:
                                 If adoConnTransactions.State = 1 Then                                                     '
                                      adoConnTransactions.Close
                                 End If
                                                 
                            rsCharge.MoveNext
                            j = j + 1
                            
                      Wend
                      'Write main table here
'                      Exit Sub
                    rsCharge.Close
                    Set rsCharge = Nothing
                            If IsDate(strLastChargeDate) = False Then
                                    GoTo EndOfAgreement
                            End If
                                      
EndOfOneManagingAgentforOneAgreement:
                       Next iManagingAgentCount
                    End If 'end if for 'X' in grid seletion
EndOfAgreement:
            Next iPropertyCount
            End If
            Next iClientCount
           If adoconn.State = 1 Then
                 adoconn.Close
           End If
           
           Dim StrPropertyCol As String
           If dicWarningProp.Count > 0 Then
                Dim oProp
                For Each oProp In dicWarningProp.Items
                    StrPropertyCol = StrPropertyCol & oProp & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid Client Global Settings setup for the property: " & StrPropertyCol, vbInformation, "Client Global Data setup"
           End If
           
           StrPropertyCol = ""
           If dicWarningAgreement.Count > 0 Then
                Dim oWar
                For Each oWar In dicWarningAgreement.Items
                    StrPropertyCol = StrPropertyCol & oWar & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid setup for the property: " & vbCrLf & StrPropertyCol, vbInformation, "Client agreement Fees and charges setup"
           End If
           StrPropertyCol = ""
           If dicWarningFinPeriod.Count > 0 Then
                Dim oWarning
                For Each oWarning In dicWarningFinPeriod.Items
                    StrPropertyCol = StrPropertyCol & oWarning & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                 'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                MsgBox "The posting date does not fall in any existing financial period for following Client(s): " & vbCrLf & StrPropertyCol, vbInformation, "Warning"
           End If
           
           
            MsgBox iCountPI & " Management Fee invoices generated", vbInformation, "Generated"
            If iCountPI > 0 Then
                adoconn.Open getConnectionString
                Call frmManagementFees.LoadFlxPurchase(adoconn)
                adoconn.Close
            End If
            frmManagementFees.fmeLoading.Visible = False
            frmManagementFees.fmeLoading.Refresh
                                  
'
                   
End Sub

Public Function NextDueDate1(FrequencyID As Integer, dtStartDate As Date, PROPERTY_ID As String) As Date
   If PROPERTY_ID = "" Then
       MsgBox "You must select a property!", vbOKOnly + vbCritical, "No property Selected"
       Exit Function
   End If

   Dim adoconn As New ADODB.Connection

   adoconn.Open getConnectionString

   If Not GetFNCGlobalDataPropertyWise(PROPERTY_ID, adoconn) Then Exit Function

   adoconn.Close
   Set adoconn = Nothing

  
      Select Case FrequencyID
            Case 1:                              'Weekly in advance
               NextDueDate1 = dtStartDate
            Case 2:                              'Weekly in arrears
              NextDueDate1 = DateAdd("d", 7, dtStartDate)
            Case 3:                              'Fortnightly in advance
               NextDueDate1 = dtStartDate
            Case 4:                              'Fortnightly in arrears
               NextDueDate1 = DateAdd("d", 14, dtStartDate)
            Case 5:                              'Monthly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Monthly)
            Case 6:                              'Monthly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Monthly)
            Case 7:                              'Quarterly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Quarterly)
            Case 8:                              'Quarterly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Quarterly)
            Case 9:                              'Half yearly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Half_Yearly)
            Case 10:                              'Half yearly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Half_Yearly)
            Case 11:                             'yearly in advance
               NextDueDate1 = NextPayingDate(dtStartDate, InAdv, Pay_Yearly)
            Case 12:                             'yearly in arrears
               NextDueDate1 = NextPayingDate(dtStartDate, InArr, Pay_Yearly)
            Case 13:                             'Daily
               NextDueDate1 = ""
            Case 14:                             '4 Weekly in advance
               NextDueDate1 = dtStartDate
            Case 15:                             '4 Weekly in arrears
               NextDueDate1 = DateAdd("d", 28, dtStartDate)
            Case 16:                             '4 Monrhly in advance
               NextDueDate1 = dtStartDate
            Case 17:                             '4 Monrhly in arrears
               NextDueDate1 = DateAdd("m", 4, dtStartDate)
      End Select

     
End Function
Public Function GetFNCGlobalDataPropertyWise(szPropertyID As String, Conn As ADODB.Connection) As Boolean
   'gets the global data from the global data table and puts the payment dates and
   'VAT rate, base rate, number of days to send demands before due, and price per sq foot for
   'service charge and puts then to global variables for when needed by program later.

   'This procedure will be called when program is opened and when the global data is
   'changed.

   Dim i As Integer, iDateSet As Integer
   Dim Rst As ADODB.Recordset
   Dim SQLStr As String
   
   Set Rst = New ADODB.Recordset
   
   SQLStr = "SELECT * FROM GlobalData WHERE PropertyID = '" & szPropertyID & "';"
   Rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If Rst.EOF Then
        MsgBox "Please enter global data for this property", vbInformation, "Warning"
        Rst.Close
        Exit Function
   End If
   Rst.Close
   SQLStr = "SELECT PropertyID, MDueDate1,MDueDate2,MDueDate3,MDueDate4,MDueDate5,MDueDate6,MDueDate7," & _
            "MDueDate8,MDueDate9,MDueDate10,MDueDate11,MDueDate12,QDueDate1,QDueDate2,QDueDate3,QDueDate4,HYDueDate1,HYDueDate2,YDueDate " & _
            "FROM GlobalData WHERE GlobalData.PropertyID = '" & szPropertyID & "';"
   
   Rst.Open SQLStr, Conn, adOpenDynamic, adLockPessimistic

   If Rst.EOF Then
       ShowMsgInTaskBar "Please Enter your Global Data.", , "N"
       Rst.Close
       Set Rst = Nothing
       GetFNCGlobalDataPropertyWise = False
       Exit Function
   End If
   If IsNull(Rst!YDueDate) Then
        MsgBox "Please Enter your Yearly Due Date in Global Data", , "Warning"
        Exit Function
   End If
   szGDYearly = Rst!YDueDate
   szGDHalfYearly1 = Rst!HYDueDate1
   szGDHalfYearly2 = Rst!HYDueDate2
   szGDQuarterly1 = Rst!QDueDate1
   szGDQuarterly2 = Rst!QDueDate2
   szGDQuarterly3 = Rst!QDueDate3
   szGDQuarterly4 = Rst!QDueDate4

   szaMonthlyGD(0) = Rst!MDueDate1
   szaMonthlyGD(1) = Rst!MDueDate2
   szaMonthlyGD(2) = Rst!MDueDate3
   szaMonthlyGD(3) = Rst!MDueDate4
   szaMonthlyGD(4) = Rst!MDueDate5
   szaMonthlyGD(5) = Rst!MDueDate6
   szaMonthlyGD(6) = Rst!MDueDate7
   szaMonthlyGD(7) = Rst!MDueDate8
   szaMonthlyGD(8) = Rst!MDueDate9
   szaMonthlyGD(9) = Rst!MDueDate10
   szaMonthlyGD(10) = Rst!MDueDate11
   szaMonthlyGD(11) = Rst!MDueDate12

   Rst.Close
   Set Rst = Nothing
   GetFNCGlobalDataPropertyWise = True
End Function
Private Function PI_Check(adoconn As ADODB.Connection, ByRef szTran2Fix As String) As Boolean
       Dim szSQL      As String
       Dim adoRst     As New ADODB.Recordset
      
        
    
       
          szSQL = "SELECT  P.TransactionID " & _
                   "FROM tlbPayment AS P, (" & _
                         "SELECT PayHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                         "From tlbPaymentSplit " & _
                         "Group by PayHeader " & _
                         ") AS Q " & _
                   "WHERE P.TransactionID = Q.PayHeader AND P.Amount <> P.OSAmount AND " & _
                         "ROUND(P.Amount - P.OSAmount, 2) <> Q.T;"
    'Debug.Print szSQL
          adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)
    
             adoRst.MoveNext
          Wend
    
          adoRst.Close
      
        szSQL = "SELECT P.SlNumber, P.TRAN_DATE, P.TransactionType, P.INV_NO, P.PropertyID, P.SlNumber, P.CL_ID, P.PostingDate, P.TOTAL_AMOUNT, P.MY_ID,  " & _
        "tblPurInvSRec.TRAN_ID, tblPurInvSRec.NOMINAL_CODE, tblPurInvSRec.DESCRIPTION " & _
        "FROM tblPurInv AS P LEFT JOIN tblPurInvSRec ON P.MY_ID = tblPurInvSRec.ParentID " & _
        "WHERE (((P.TransactionType)=6 Or (P.TransactionType)=7) AND ((P.NLPost)=False))  AND P.TOTAL_AMOUNT<>0 AND  tran_ID is null; "
    'Debug.Print szSQL
    'This means there is PI without splits
          adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
            'if you found this problem exists then there is some transaction in tblPurInv which are inconsitence
            'Now by comparing with the NLposting this transaction will be zerorized
            'To prevent happening this again there is check I have implemented  in PI
             adoRst.MoveNext
          Wend
    
          adoRst.Close
          szSQL = "SELECT  P.SlNumber,P.TOTAL_AMOUNT , Q.T FROM tblPurinv AS P, (SELECT ParentID, Sum(ROUND(TOTAL_AMOUNT, 2)) AS T From tblPurInvSRec Group by ParentID ) AS Q  " & _
            "WHERE P.MY_ID = Q.ParentID AND round(P.TOTAL_AMOUNT,2) <> Q.T;"
           adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
             adoRst.MoveNext
          Wend
          
          adoRst.Close
          Set adoRst = Nothing
          'This part shall check for the dupplicate serial number in the tblPurInv
          szSQL = "SELECT slnumber,transactiontype, COUNT(*) From tblPurInv GROUP BY slnumber, transactiontype HAVING COUNT(*) > 1"
          adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
          While Not adoRst.EOF
             szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("SlNumber").Value)
             MsgBox "Duplicate SlNumber in tlbpurinv while creating PI", vbInformation, "Warning-" & adoRst.Fields.Item("SlNumber").Value
             adoRst.MoveNext
          Wend
          
          adoRst.Close
          Set adoRst = Nothing
     
       If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)
    
       If Len(szTran2Fix) > 0 Then
            PI_Check = False
       Else
            PI_Check = True
            'MsgBox "HI"
       End If
      
End Function
Private Sub GeneratePIPreview()
        Dim percentageOramount1 As Double
        Dim lSlNumber As Long
        Dim adoconn As New ADODB.Connection
        Dim adoPIHeader As New ADODB.Recordset
        Dim adoPISplit As New ADODB.Recordset
        Dim szSQL As String
        Dim szSQLManagingAgent  As String
        Dim szMYID As String
        Dim szFundID As String
        Dim szSelectedPayableTypeID As String
        Dim szSelectedClient As String
        Dim szPropertySelection1 As String
        Dim szSQL1 As String
        Dim rsfixedMethod As New ADODB.Recordset
        Dim rsfixedMethodDetails As New ADODB.Recordset
        Dim j As Integer
        Dim vatFixedBasis As Double
        'Dim percnetageOramount As Double
        Dim dblGrandTotal As Double
        Dim dtNextDue As Date
        Dim dtFDD As Date
        Dim dblFeqID As Integer
        Dim dblTotalAmount As Double
        Dim lngMgtFeeSL As Long
        Dim iCountPI As Integer
        Dim iClientCount As Integer
        Dim iPropertyCount As Integer
        Dim strLastChargeDate  As String
        Dim strFundName As String
        Dim strStopDate  As String
        Dim dblCapAmount As Double
        Dim rsManagingAgent As New ADODB.Recordset
        Dim rsGlobalData As New ADODB.Recordset
        Dim szManagingAgent() As String
        Dim iManagingAgentCount As Integer
        Dim bControlACForPayable As Boolean
        Dim FinalControlACForPayable As String
        Dim szTemp
        Dim dblNoOfDaysToSendMFB4Due As Integer
        Dim dtNDDInitial As Date
        Dim strFromDate As String
        Dim strToDate As String
        Dim rsFromandToDate As New ADODB.Recordset
        Dim szSQLFrom As String
        Dim percnetageOramount As Double
        Dim SQLforInsert As String
        Dim dtNDD As Date
        
        Dim iCount As Long
        Dim iCount1 As Long
        Dim rCount As Long
        Dim dblFundId As Integer
        Dim dblDemandTypeId As Integer
        Dim i As Integer
        Dim lT_ID As Long
        Dim dicWarningProp As New Dictionary
        Dim dicWarningAgreement As New Dictionary
         Dim dicWarningFinPeriod As New Dictionary
      '  Dim strSelectedChargeType As String
         Dim strSelectedFundID As String
         Dim szSQL5 As String
         Dim rstSet As New ADODB.Recordset
         Dim szPropertySelectionALL As String
         Dim warning1 As String
         Dim warning2 As String
         Dim warning3 As String
         Dim descriptionAndDate As String
         
        If txtChargeDate.text = "" Then
            MsgBox "Please enter invoice date", vbInformation, "Warning "
            FocusControl txtChargeDate
            Exit Sub
        End If
        strSelectedFundID = ""
        For iCount = 1 To flxDemandTypes.Rows - 1
            If flxDemandTypes.TextMatrix(iCount, 2) <> "" Then
                   iCount1 = iCount1 + 1
                   If flxDemandTypes.TextMatrix(iCount, 0) = "X" Then
                        strSelectedFundID = strSelectedFundID + "" + flxDemandTypes.TextMatrix(iCount, 2) + ","
                   End If
            End If
        Next
        If Len(strSelectedFundID) > 0 Then
                strSelectedFundID = Left(strSelectedFundID, Len(strSelectedFundID) - 1)
        End If

        
        For rCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(rCount, 0) = "X" Then
               szSelectedClient = flxClients.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
       
'        For rCount = 1 To flxProperties.Rows - 1
'                If flxProperties.TextMatrix(rCount, 0) = "X" Then
'                    szPropertySelectionALL = szPropertySelectionALL + "," + flxProperties.TextMatrix(rCount, 1)
'
'                End If
'        Next
      For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    If szPropertySelectionALL = "" Then
                        szPropertySelectionALL = "'" + flxProperties.TextMatrix(rCount, 1) + "'"
                    Else
                        szPropertySelectionALL = szPropertySelectionALL + ",'" + flxProperties.TextMatrix(rCount, 1) & "'"
                    End If

                End If
        Next

        For rCount = 1 To flxProperties.Rows - 1
            If flxProperties.TextMatrix(rCount, 0) = "X" Then
               szPropertySelection1 = flxProperties.TextMatrix(rCount, 1)
               Exit For
            End If
        Next
         Debug.Print time & " 1 start"
        For rCount = 1 To flxProperties.Rows - 1
                If flxProperties.TextMatrix(rCount, 0) = "X" Then
                    szPropertySelection1 = flxProperties.TextMatrix(rCount, 1)
                    If checksetupDONE(szPropertySelection1) = True Then
                        ' MsgBox "Please enter a valid setup for the property: " & szPropertySelection1, vbInformation, "Client agreement Fees and charges setup"
                          dicWarningAgreement.Add rCount, szPropertySelection1
                    End If
'                    If hasChargeTypeinlist(szPropertySelection1) = False Then
'                              MsgBox "Please setup a demand type for the selected property: " & szPropertySelection1, vbInformation, "Warning"
'                    End If
                End If
        Next
         Debug.Print time & " 2 End"
        
        If szPropertySelection1 = "" Then
            MsgBox "Please select a Property.", vbInformation, "Warning"
            Exit Sub
        End If

        If strSelectedFundID = "" Then
            MsgBox "Please select a fund", vbInformation, "Warning"
            Exit Sub
        End If
        If chkAssignProperty.Value = 1 Then
            If MsgBox("Are you sure you wish to generate your management fees without assigning a property", vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
        
        
 '      ************************************Write tblPurInv **************************************
        Picture1.Visible = True 'loading msg
        Picture1.Refresh
        adoconn.Open getConnectionString
        adoconn.Execute "Delete from tblPurInvPreview"
        adoconn.Execute "Delete from tblPurInvSRecPreview"
        adoconn.Execute "Delete from ManagementFeePreview"
       lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
        adoconn.Close
  
    For iClientCount = 1 To flxClients.Rows - 1
            If flxClients.TextMatrix(iClientCount, 0) = "X" Then 'also check financial year is correct period and control account is there and it is set in a variable before looping
                        szSelectedClient = flxClients.TextMatrix(iClientCount, 1)
                        If szSelectedClient = "" Then Exit Sub
                        adoconn.Open getConnectionString
                        
                        If FinalControlACForPayable = "" Then
                                  FinalControlACForPayable = GetNominalCodeForControlAccount(adoconn, "Management Fee Payable (P&L)", szSelectedClient)
                        End If
                            'if still control acount is empty generate a warning message and exit this sub procedure
                        If FinalControlACForPayable = "" Then
                                MsgBox "Control Account is not set for this Payable type for Client:" & szSelectedClient, vbInformation, "Warning"
                                GoTo EndOfAgreement
                        End If
                        If IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoconn) = 0 Then
                            MsgBox "The posting date cannot fall within a closed financial period for the client :" & szSelectedClient, vbInformation, "Warning"
                            adoconn.Close
                            'FocusControl txtChargeDate
                            GoTo EndOfAgreement
                        ElseIf IsPeriodStatus(txtPostingDate.text, szSelectedClient, adoconn) = 9 Then
                            'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                            dicWarningFinPeriod.Add iClientCount, szSelectedClient
                            adoconn.Close
                            'FocusControl txtChargeDate
                            GoTo EndOfAgreement
                        End If
                        adoconn.Close
            
            For iPropertyCount = 1 To flxProperties.Rows - 1
                    If iPropertyCount >= flxProperties.Rows Then Exit For
                    If flxProperties.TextMatrix(iPropertyCount, 0) = "X" And szSelectedClient = flxProperties.TextMatrix(iPropertyCount, 3) Then
                            Debug.Print iPropertyCount
                            szPropertySelection1 = flxProperties.TextMatrix(iPropertyCount, 1)
                            If checkGlobalDataEntered(szPropertySelection1) = True Then
                                dicWarningProp.Add iPropertyCount, szPropertySelection1
                                 'MsgBox "Please enter a valid Client Global Settings setup for the property: " & szPropertySelection1, vbInformation, "Client Global Data setup"
                                 'Exit For
                                GoTo EndOfAgreement
                            End If
                            adoconn.Open getConnectionString
'                            adoconn.BeginTrans
                            
                            
                            
            Dim rsCharge As New ADODB.Recordset
            szTemp = ""
            szSQLManagingAgent = "SELECT DISTINCT agr.ManagingAgentID " & _
              "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
              "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
              "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
              "CPA.PropertyID = '" & szPropertySelection1 & "' ANd F.FundID IN(" & strSelectedFundID & ")"
              rsManagingAgent.Open szSQLManagingAgent, adoconn, adOpenDynamic, adLockOptimistic
              szTemp = SQL2String(rsManagingAgent, 0)
              rsManagingAgent.Close
              If Len(szTemp) > 0 Then
                    szManagingAgent = Split(szTemp, ",")
              Else
                    adoconn.Close
                    GoTo EndOfAgreement
              End If
              
              
             szSQL5 = "SELECT MAX(ManagementFeeSL) AS x FROM tblPurInv;"
             rstSet.Open szSQL5, adoconn, adOpenStatic, adLockReadOnly
             lngMgtFeeSL = Val(IIf(IsNull(rstSet.Fields.Item(0).Value), 0, rstSet.Fields.Item(0).Value)) + 1
             rstSet.Close
             Set rstSet = Nothing
             
             adoconn.Close
             
            For iManagingAgentCount = 0 To UBound(szManagingAgent) 'this for shall end after the creation of the PI preview
              
            adoconn.Open getConnectionString
            'For each managing agent I am creating PI
            'lSlNumber = SlNumber("PI", "tblPurInv", adoconn)
             
            szSQL = "SELECT agr.EachPeriod,agr.Capamount, agr.StopDate, CPA.agreementStartDate,CPA.agreementEndDate,agr.CHARGE_METHOD," & _
                "cpa.agreementEndDate as agreementEndD,agr.LastChargeDate,agr.TotalAmount,agr.Amount,agr.Fund,F.FundName, " & _
                "agr.NtDueDate,agr.FDD,agr.Frequency as FrequencyID,(Select FC.Frequency from Frequencies FC where  FC.ID=agr.Frequency) as FrequencyName,ManagingAgentID " & _
                "FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC  " & _
                "WHERE agr.CPA_ID = CPA.CPA_ID AND  F.FundID=agr.fund AND SC.CODE=agr.CHARGE_METHOD AND " & _
                "CPA.ClientID = '" & szSelectedClient & "' And C.ID = agr.CHARGE_TYPE And " & _
                "CPA.PropertyID = '" & szPropertySelection1 & "' AND F.FundID in (" & strSelectedFundID & ") AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
              
              
            '    szSQL = "SELECT agr.TotalAmount,agr.Fund,agr.Frequency ,agr.NtDueDate,agr.FDD " & _
            '            "FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
            '            "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
            '            "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
            '            "CPA.PropertyID = '" & szPropertySelection1 & "'"
            rsCharge.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
            If rsCharge.EOF Then
                rsCharge.Close
                Set rsCharge = Nothing
                'MsgBox "Please setup an agreement for this property:" & szPropertySelection1, vbInformation, "Client agreement Fees and charge setup"
                dicWarningProp.Add iPropertyCount, szPropertySelection1
                GoTo EndOfAgreement
            End If

                            
            i = 1
                'Dim rsGlobalData As New ADODB.Recordset
                Dim VAT_ID As String
                Dim VAT_CODE As String
                Dim VAT_RATE As Double
                
                 rsGlobalData.Open "Select vatOptionEnabled,V.VAT_ID,V.VAT_CODE,V.VAT_RATE from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' AND vatOptionEnabled=true", adoconn, adOpenStatic, adLockReadOnly
                 If Not rsGlobalData.EOF Then
                          VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                          VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                          VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                 Else
                          VAT_ID = -1
                          VAT_RATE = 0
                          VAT_CODE = ""
                 End If
                 rsGlobalData.Close
                 Set rsGlobalData = Nothing
                 
                 Dim rsGlobalData1 As New ADODB.Recordset
                 Dim bolVatOptionEnabled As Boolean
                 Dim bolOptedTotax As String
                 Dim strManagingAgentID As String
                 rsGlobalData1.Open "Select vatOptionEnabled from (Globaldata G LEFT JOIN tlbVatCode V ON G.vatRate=V.VAT_ID) where PropertyID='" & _
                                    szPropertySelection1 & "' ", adoconn, adOpenStatic, adLockReadOnly
                                    
                If Not rsGlobalData1.EOF Then
                        bolVatOptionEnabled = IIf(IsNull(rsGlobalData1("vatOptionEnabled").Value), False, rsGlobalData1("vatOptionEnabled").Value)
                        strManagingAgentID = rsCharge("ManagingAgentID").Value
                End If
                rsGlobalData1.Close
                
              
              
           
      
            szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
            adoPIHeader.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
            adoPIHeader.Close
             j = 1
                 rsGlobalData1.Open "SELECT optedTotax,* FROM Supplier where supplierID='" & strManagingAgentID & "'", adoconn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData1.EOF Then
                        bolOptedTotax = rsGlobalData1("optedTotax").Value
                Else
                        bolOptedTotax = False
                End If
                rsGlobalData1.Close
           
   
            'Dim percnetageOramount As Double
            While Not rsCharge.EOF
                  dblTotalAmount = rsCharge.Fields.Item("TotalAmount").Value
                  dblCapAmount = rsCharge.Fields.Item("CapAmount").Value
                  dblFundId = rsCharge.Fields.Item("Fund").Value
                  'dblDemandTypeId = rsCharge.Fields.Item("DEMAND_TYPE").Value
                  
                  strFundName = rsCharge.Fields.Item("FundName").Value
                  If Not IsNull(rsCharge.Fields.Item("NtDueDate").Value) Then
                      dtNextDue = rsCharge.Fields.Item("NtDueDate").Value
                      dtNDDInitial = rsCharge.Fields.Item("NtDueDate").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FDD").Value) Then
                      dtFDD = rsCharge.Fields.Item("FDD").Value
                  End If
                  If Not IsNull(rsCharge.Fields.Item("FrequencyID").Value) Then
                    If rsCharge.Fields.Item("FrequencyID").Value <> "" Then
                            dblFeqID = rsCharge.Fields.Item("FrequencyID").Value
                      End If
                  End If
                szMYID = UniqueID()
                strLastChargeDate = IIf(IsNull(rsCharge("LastChargeDate").Value), "", rsCharge("LastChargeDate").Value)
                If strLastChargeDate = "" Then
                       rsCharge.Close
                       MsgBox "Please enter a last charge date for Property:" & szPropertySelection1, vbInformation, "Warning"
                       GoTo EndOfAgreement
                End If
                rsGlobalData.Open "Select NoOfDaysToSendMFB4Due from globaldata where PropertyID='" & szPropertySelection1 & "'", adoconn, adOpenStatic, adLockReadOnly
                If Not rsGlobalData.EOF Then
                    dblNoOfDaysToSendMFB4Due = IIf(IsNull(rsGlobalData!NoOfDaysToSendMFB4Due), 0, rsGlobalData!NoOfDaysToSendMFB4Due)
                End If
                rsGlobalData.Close
               
                If DateDiff("d", Date, rsCharge("NtDueDate").Value) > dblNoOfDaysToSendMFB4Due Then
                        GoTo EndOfChargeType
                End If
                 strStopDate = rsCharge("StopDate").Value
                If strStopDate = "" Then
                   ' Exit Sub
                Else
                    If IsDate(strStopDate) Then
                            If DateDiff("d", strStopDate, txtChargeDate.text) >= 0 Then
'                                MsgBox "It is not possible to generate fees after the stop date.:" & szPropertySelection1, vbInformation, "Warning"
                            End If
                    Else
                            MsgBox "Stop date date format is not correct.:" & szPropertySelection1, vbInformation, "Warning"
                    End If
                    GoTo EndOfChargeType
                End If
            
                           
                                'validations
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then   'when working on the fixed procedure only 1 line of setup is done then (fixed basis)
                                                    'This validation is not valid for fixed method
'                                                dblTotalAmount = IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value)
'                                                If dblTotalAmount = 0 Then
'                                                        MsgBox "There are no fees to generate for property:" & szPropertySelection1, vbInformation, "Warning"
'                                                        GoTo EndOfAgreement
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtChargeDate.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                End If
'                                                 percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                            '      ************************************Write tblPurInvSRec **************************************
                                                                                                                                   
                                                    dblTotalAmount = IIf(IsNull(rsCharge("EachPeriod").Value), 0, rsCharge("EachPeriod").Value) 'rsfixedMethodDetails.Fields.Item("Amt").Value
        '                                            dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                    dblFundId = rsCharge.Fields.Item("Fund").Value
'                                                                                            If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
'                                                                                                  dblTotalAmount = dblTotalAmount * percnetageOramount / 100
'                                                                                            End If
                                                    'dblGrandTotal = 50
                                                    If dblCapAmount > 0 Then
                                                           'make a condition in the split as well so that amount doesnot exeed cap amount
                                                           If dblTotalAmount > dblCapAmount Then
                                                                dblTotalAmount = dblCapAmount
                                                           End If
                                                    End If
'                                                    Dim rsNextdueDate As New ADODB.Recordset
'                                                    szSQL = "SELECT min(agr.NtDueDate),agr.FDD,agr.Frequency as FrequencyID,ManagingAgentID,agr.DEMAND_TYPE  " & _
'                                                            "FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
'                                                            "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND  SC.CODE=agr.CHARGE_METHOD AND " & _
'                                                            "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'                                                            "CPA.PropertyID = '" & szPropertySelection1 & "' AND DEMAND_TYPE IN(" & strSelectedFundID & ") " & _
'                                                            "AND agr.ManagingAgentID='" & Trim(szManagingAgent(iManagingAgentCount)) & "'"
'                                                    rsNextdueDate.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                                                    If Not rsNextdueDate.EOF Then
                                                                 txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
                                                                dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
                                                                dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
                                                                strFromDate = Format(dtNextDue, "dd/MM/yyyy")
                                                                strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy")
'                                                    End If
'                                                    rsNextdueDate.Close
                                                    
                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                   ' adoPISplit.Close
                                                    adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                    'Add New Records. At least there is only one split line
                                                       With adoPISplit
                                                           .AddNew
                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                           .Fields.Item("ParentID").Value = szMYID
                                                           .Fields.Item("TRAN_ID").Value = j
                                                          ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                          If chkAssignProperty.Value = 0 Then
                                                                 .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                            Else
                                                                 .Fields.Item("TRANS").Value = ""
                                                            End If
                                                           .Fields.Item("UNIT_ID").Value = ""
                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                           .Fields.Item("RecoverablePt").Value = 0
                                                           '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                
                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                 vatFixedBasis = .Fields.Item("VAT").Value
                                                                .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                           ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled means global data and bolOptedTotax is supplier table
                                                                rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where  (S.VATCode)=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                   strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                If Not rsGlobalData.EOF Then
                                                                         VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                         VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                         VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                Else
                                                                         VAT_ID = -1
                                                                         VAT_RATE = 0
                                                                         VAT_CODE = ""
                                                                End If
                                                                rsGlobalData.Close

                                                                 'done modification on 15-10-2021
                                                                    '.Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                    .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                    .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                    .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                     dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                     
                                                          ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                 dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                          End If
                                                          .Update
                                                       End With
                                                    adoPISplit.Close
                                                   If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then  'In preview mode we are not updating any FDD
                                                                                     'Updating FDDand next due date
'                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNextDue)
'                                                                                        dtNDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                                        txtComparenextDueDate1 = DateAdd("d", 1, dtNDD)
'                                                                                        dtFDD = NextDueDate1(CInt(dblFeqID), txtComparenextDueDate1, szPropertySelection1)
'                                                                                        strFromDate = Format(dtNextDue, "dd/MM/yyyy")
'                                                                                        strToDate = Format(DateAdd("d", -1, dtNDD), "dd/MM/yyyy") ' Format(DateAdd("d", -1, dtFDD), "dd/MMM/yyyy")
'                                                                                      szSQL = "Update tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
'                                                                                     " Set NtDueDate=#" & dtNDD & "# ,FDD=#" & dtFDD & "#,lastchargeDate=#" & txtChargeDate & "# " & _
'                                                                                     "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
'                                                                                     "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'                                                                                     "CPA.PropertyID = '" & szPropertySelection1 & "' and AGREEMENT_ID=" & rsCharge.Fields.Item("AGREEMENT_ID").Value & ";"
'                                                                                     adoConnTransactions.Execute szSQL
                                                                                     
                                                                                     'adoConnTransactions.Execute szSQL
                                                   End If
                                                   dblGrandTotal = dblGrandTotal + dblTotalAmount
'

                                                 SQLforInsert = "'" & szMYID & "','" & Format(txtChargeDate.text, "dd MMM yyyy") & "','" & szPropertySelection1 & "','Fixed Basis'," & dblTotalAmount - vatFixedBasis & "," & vatFixedBasis & "," & dblTotalAmount & ")"
                                                 adoconn.Execute "Insert into ManagementFeePreview(PI_ActualID,ChargeDate,PropertyID,ChargingMethod," & _
                                                "MgtFeeAmt,VAT,MgtFeeAmtTotal) values (" & _
                                                SQLforInsert
                                                vatFixedBasis = 0
                                                                       
                                                                       
                                End If 'end of rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX"
                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then 'RE_ABL what does it mean?????
'                                                If IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value) = 0 Then
'                                                        rsfixedMethod.Close
'                                                        GoTo EndOfChargeType
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtChargeDate.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                        
                                                End If
                                                 szSQL1 = "Select sum(SWITCH(TYPE =1,RS.Amount,TYPE =2,-RS.Amount,TYPE =24,RS.Amount)) as Amt from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                              "R.TransactionID=RS.RptHeader AND RDate<=#" & Format(txtChargeDate.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  and Type in (1,2,24) " & _
                                                              "AND U.UnitNumber=R.UnitID AND U.PropertyID='" & szPropertySelection1 & "' and ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                                              'need to consider the selected property in where clause
                                                            '  rsfixedMethod.Close
                                                    rsfixedMethod.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for recipt type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    
                            '      ************************************Write tblPurInvSRec **************************************
                            
                                                                 szSQL = "Select sum(SWITCH(TYPE =3,RS.Amount,TYPE =4,-RS.Amount,TYPE =23,RS.Amount)) as Amt,rs.FundID from tlbReceipt R,tlbReceiptsplit RS,Units U where " & _
                                                                         "R.TransactionID=RS.RptHeader and Type in (3,4,23) AND RDate<=#" & Format(txtChargeDate.text, " dd MMM yyyy") & "# AND RDate>#" & Format(strLastChargeDate, " dd MMM yyyy") & "#  AND  " & _
                                                                         "U.UnitNumber=R.UnitID and ISMGTFEES=false AND U.PropertyID='" & szPropertySelection1 & "' AND RS.FundID=" & dblFundId & " group by RS.FundID"
                                                                 'need to consider the selected property in where clause
                                                                 If rsfixedMethodDetails.State = 1 Then
                                                                     rsfixedMethodDetails.Close
                                                                 End If
                                                                 rsfixedMethodDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                                                  If Not rsfixedMethodDetails.EOF Then 'we rare using while because itr
                                                                                 dblTotalAmount = rsfixedMethodDetails.Fields.Item("Amt").Value
                                                                                 If dblTotalAmount < 0 Then
                                                                                        GoTo EndOfChargeType
                                                                                 End If
                                                                                 dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                                 If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                                       dblTotalAmount = dblTotalAmount * percnetageOramount / 100
                                                                                 End If
                                                                                 
                                                                                 If dblCapAmount > 0 Then
                                                                                        If dblTotalAmount > dblCapAmount Then
                                                                                             dblTotalAmount = dblCapAmount
                                                                                        End If
                                                                                 End If
                                                                                 szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                                 adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                                                 'Add New Records. At least there is only one split line
                                                                                    With adoPISplit
                                                                                        .AddNew
                                                                                        .Fields.Item("MY_ID").Value = UniqueID()
                                                                                        .Fields.Item("ParentID").Value = szMYID
                                                                                        .Fields.Item("TRAN_ID").Value = j
                                                                                       ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         If chkAssignProperty.Value = 0 Then
                                                                                              .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                                         Else
                                                                                              .Fields.Item("TRANS").Value = ""
                                                                                         End If
                                                                                        .Fields.Item("UNIT_ID").Value = ""
                                                                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                                        .Fields.Item("DEPT_ID").Value = dblFundId
                                                                                       ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                                        .Fields.Item("RecoverablePt").Value = 0
                                                                                        '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                                        
                                                                                        .Fields.Item("description").Value = "Management Fees for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                         descriptionAndDate = "Management Fees for " & strFundName & " " & DateAdd("d", 1, CDate(strFromDate)) & " - " & strToDate & ""
                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                        .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                        .Update
                                                                                    End With
                                                                                 adoPISplit.Close
                                                                                
                                                                                 dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                            End If
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ABL" Then
                                If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then ' receipt basis
'                                                If IIf(IsNull(rsfixedMethod("amt").Value), 0, rsfixedMethod("amt").Value) = 0 Then
'                                                        rsfixedMethod.Close
'                                                        GoTo EndOfChargeType
'                                                End If
                                                If rsCharge("agreementEndD").Value < rsCharge("FDD").Value Then
                                                        MsgBox "agreement End Date is greatar than following due date for the property:" & szPropertySelection1
                                                        GoTo EndOfAgreement
                                                End If
                                                If DateDiff("d", txtChargeDate.text, rsCharge.Fields.Item("agreementStartDate").Value) > 0 Then
                                                        rsCharge.Close
                                                        Set rsCharge = Nothing
                                                        MsgBox "Charge date cannot be before Agreement Start date for property:" & szPropertySelection1, vbInformation, "Warning"
                                                        GoTo EndOfChargeType
                                                        
                                                End If
                                
'                                            add report table and insert data into it
                                                    
                                            szSQL1 = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
                                            "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                            "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                            "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                            szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     
                                                    'here tlbReceipt R is reciopt table and RS,tlbReceipt R1 is invoice table
                                                    'need to consider the selected property in where clause
                                                    'Need to take only allocated transactions
                                                            '  rsfixedMethod.Close
                                                        'newly added this line
                                                        ' "Units U, (Select distinct fromtran from rptTransactions where deleteflag=false) as A  where A.Fromtran= R.TransactionID AND " & _
                                                        'by anol 20211024
                                                        If rsfixedMethod.State = 1 Then
                                                                rsfixedMethod.Close
                                                        End If
                                                    rsfixedMethod.Open szSQL1, adoconn, adOpenStatic, adLockReadOnly
                                                    'Here type 3 is for reciept type . I have not written for the credit yet need to understand the principle
                                                    
                                                    If rsfixedMethod.EOF Then
                                                        rsfixedMethod.Close
                                                        Set rsfixedMethod = Nothing
                                                        GoTo EndOfChargeType
                                                    End If
                                                    percnetageOramount = IIf(IsNull(rsCharge("amount").Value), 0, rsCharge("amount").Value)
                                                    percentageOramount1 = percnetageOramount
                                                    
                            '      ************************************Write tblPurInvSRec **************************************

'                                     szSQL = "Select  sum(SWITCH(R.TYPE =3,RS.Amount,R.TYPE =4,RS.Amount,R.TYPE =23,-RS.Amount))  as Amt from tlbReceipt R,tlbReceiptsplit RS, " & _
'                                     "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
'                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
'                                     "AND  R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                     szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                     'Modification done by anol 08-07-2023
                                     
                                     szSQL = "Select  (SWITCH(R.TYPE =3,AL.NetAmount,R.TYPE =4,AL.NetAmount,R.TYPE =23,-AL.NetAmount))  as Amt," & _
                                     "(SWITCH(R.TYPE =3,R.Amount,R.TYPE =4,R.Amount,R.TYPE =23,-R.Amount)) as Orig from tlbReceipt R,tlbReceiptsplit RS, " & _
                                     "rptTransactionsSPlit AL, Units U where AL.deleteflag=false AND " & _
                                     "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                     "AND  R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
                                     szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & ""
                                                        'need to consider the selected property in where clause
                                                        'modified on 20211103
                                                        SQLforInsert = "Select " & szMYID & " as MYID ,'" & Format(txtChargeDate.text, "dd MMM yyyy") & _
                                                "' as ChargeDate,R.SlNumber, R.Type,R.SageAccountNumber,R.Ref,U.PropertyID,RS.FundID,RptAmtType,ExtRef,Rdate,R.TransactionID,RS.SplitID, " & _
                                                "(SWITCH(R.TYPE =3,AL.NETAmount,R.TYPE =4,AL.NETAmount,R.TYPE =23,-AL.NETAmount)) as Amt from tlbReceipt R," & _
                                                "tlbReceiptsplit RS, rptTransactionsSplit AL, Units U where AL.deleteflag=false AND " & _
                                                "AL.TransactionID= RS.RptTransactionsIDSplit AND R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                "and R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID  AND U.PropertyID='" & _
                                                 szPropertySelection1 & "' and RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                 
                                            adoconn.Execute "Insert into ManagementFeePreview(PI_ActualID,ChargeDate,SRSlNumber,ReceiptType,SageAccountNumber,ReceiptTypeDescription," & _
                                                    "PropertyID,FundID,RptAmtType,ExtRef,ReceiptDate,ReceiptTransactionID,ReceiptSplitID," & _
                                                "ReceiptAmount)" & _
                                                SQLforInsert
                                                
                                                
'                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from tlbReceipt R,tlbReceiptsplit RS, " & _
'                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U where  AL.deleteflag=false AND " & _
'                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
'                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=R.UnitID AND U.PropertyID='" & _
'                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                'SQL Modification done on 07-08-2023
                                                        szSQLFrom = "Select  min(DS.DateFrom) as  DateFromMin ,max(DS.DateTO) as DateTOMAX   from  tlbReceipt R, tlbReceipt R1,tlbReceiptsplit RS,  " & _
                                                        "rptTransactionsSPlit AL, DemandSplitRecords DS, Units U,DemandRecords DR where  DS.DemandID=DR.DemandID and DS.DemandID= R1.DemandRef AND R1.transactionID=AL.ToTran AND AL.deleteflag=false AND " & _
                                                        "AL.TransactionID= RS.RptTransactionsIDSplit AND  R.TransactionID=RS.RptHeader AND R.RDate<=#" & Format(txtChargeDate.text, "dd MMM yyyy") & "# " & _
                                                        "AND R.Type in (3,4,23)  AND U.UnitNumber=DR.UnitNumber AND U.PropertyID='" & _
                                                        szPropertySelection1 & "' AND RS.ISMGTFEES=false AND Rs.FundID=" & dblFundId & " "
                                                        
                                                        
                                                        rsFromandToDate.Open szSQLFrom, adoconn, adOpenStatic, adLockReadOnly
                                                        If Not rsFromandToDate.EOF Then
                                                                strFromDate = Format(rsFromandToDate("DateFromMin").Value, "dd/MM/yyyy")
                                                                strToDate = Format(rsFromandToDate("DateTOMAX").Value, "dd/MM/yyyy")
                                                        Else
                                                                strFromDate = Null
                                                                strToDate = Null
                                                        End If
                                                        rsFromandToDate.Close
                                                        
                                                        Dim dblOriginalAmount As Double
                                                        'Need to take only allocated transactions
                                                        j = 1
                                                        If rsfixedMethodDetails.State = 1 Then
                                                            rsfixedMethodDetails.Close
                                                        End If
                                                        rsfixedMethodDetails.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                                                              While Not rsfixedMethodDetails.EOF  'we rare using while because itr
                                                                    dblOriginalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Orig").Value), 0, rsfixedMethodDetails.Fields.Item("Orig").Value)
                                                                    dblTotalAmount = IIf(IsNull(rsfixedMethodDetails.Fields.Item("Amt").Value), 0, rsfixedMethodDetails.Fields.Item("Amt").Value)
                                                                    'dblOriginalAmount = rsfixedMethodDetails.Fields.Item("NETAMT").Value
                                                                    If dblTotalAmount <= 0 Then
                                                                           rsfixedMethod.Clone
                                                                           GoTo EndOfChargeType
                                                                    End If
                                                                    'dblFundId = rsfixedMethodDetails.Fields.Item("FundID").Value
                                                                    If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                                                                           dblTotalAmount = dblTotalAmount * (percnetageOramount / 100)
                                                                           dblTotalAmount = Round(dblTotalAmount, 2)
                                                                    End If
                                                                    
                                                                    If dblCapAmount > 0 Then
                                                                           If dblTotalAmount > dblCapAmount Then
                                                                                dblTotalAmount = dblCapAmount
                                                                           End If
                                                                    End If
                                                                    szSQL = "SELECT * FROM tblPurInvSRecPreview"
                                                                    adoPISplit.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic
                                                                    'Add New Records. At least there is only one split line
                                                                       With adoPISplit
                                                                           .AddNew
                                                                           .Fields.Item("MY_ID").Value = UniqueID()
                                                                           .Fields.Item("ParentID").Value = szMYID
                                                                           .Fields.Item("TRAN_ID").Value = j
                                                                           j = j + 1
                                                                          ' .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                            If chkAssignProperty.Value = 0 Then
                                                                                 .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
                                                                            Else
                                                                                 .Fields.Item("TRANS").Value = ""
                                                                            End If
                                                                           .Fields.Item("UNIT_ID").Value = ""
                                                                           .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
                                                                           .Fields.Item("DEPT_ID").Value = dblFundId
                                                                          ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
                                                                           .Fields.Item("RecoverablePt").Value = 0
                                                                           '' (Current Charge date)" '"MFee" + szPropertySelection1 + Format(lngMgtFeeSL, "0000") '"Management Fee"
                                                                           .Fields.Item("description").Value = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                                           descriptionAndDate = "Management Fees for " & strFundName & " (" & strFromDate & " - " & strToDate & ")"
                                                                            '.Fields.Item("Description").Value = "Management Fee " & percentageOramount & "% of £" & Round(dblOriginalAmount, 2)
                                                                           'Original Part
'                                                                                        .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
'                                                                                        .Fields.Item("TAX_CODE").Value = VAT_CODE
'                                                                                        .Fields.Item("VAT").Value = Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
'                                                                                        .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
'                                                                                         dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                           If bolVatOptionEnabled = True And bolOptedTotax = True Then
                                                                                   'dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = VAT_CODE
                                                                                   .Fields.Item("VAT").Value = Round(dblTotalAmount * (VAT_RATE / 100), 2) 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = .Fields.Item("VAT").Value + dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoconn.Execute "Update ManagementFeePreview set AgrPercentage=" & percnetageOramount & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                                              ElseIf bolVatOptionEnabled = True And bolOptedTotax = False Then 'bolVatOptionEnabled=global data
                                                                                   'Modified by anol 2021-10-15
                                                                                    dblTotalAmount = dblTotalAmount * Round((100 / (100 + VAT_RATE)), 2)
                                                                                    .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoconn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                                             ElseIf bolVatOptionEnabled = False And bolOptedTotax = True Then 'bolVatOptionEnabled=global data
                                                                                   rsGlobalData.Open "Select V.VAT_ID,V.VAT_CODE,V.VAT_RATE from  Supplier S,tlbVatCode V where S.VATCode=cstr(V.VAT_ID)  AND  SupplierID='" & _
                                                                                                      strManagingAgentID & "' ", adoconn, adOpenStatic, adLockReadOnly
                                                                                   If Not rsGlobalData.EOF Then
                                                                                            VAT_ID = IIf(IsNull(rsGlobalData.Fields("VAT_ID").Value), "-1", rsGlobalData.Fields("VAT_ID").Value)
                                                                                            VAT_RATE = IIf(IsNull(rsGlobalData.Fields("VAT_RATE").Value), "0.00", rsGlobalData.Fields("VAT_RATE").Value)
                                                                                            VAT_CODE = IIf(IsNull(rsGlobalData.Fields("VAT_CODE").Value), "", rsGlobalData.Fields("VAT_CODE").Value)
                                                                                   Else
                                                                                            VAT_ID = -1
                                                                                            VAT_RATE = 0
                                                                                            VAT_CODE = ""
                                                                                   End If
                                                                                   rsGlobalData.Close
                                                                                   'done modification on 15-10-2021
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null ' VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") ' Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE' + Format(dblTotalAmount * (VAT_RATE / 100), "0.00")
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount + Round(dblTotalAmount * (VAT_RATE / 100), 2)
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoconn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=" & VAT_RATE & " where MgtFeeAmtTotal is null"
                                                                             ElseIf bolVatOptionEnabled = False And bolOptedTotax = False Then
                                                                                   .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
                                                                                   .Fields.Item("TAX_CODE").Value = Null 'VAT_CODE
                                                                                   .Fields.Item("VAT").Value = 0 'Format(dblTotalAmount * (VAT_RATE / 100), "0.00") 'VAT_RATE
                                                                                   .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                                                                    dblTotalAmount = .Fields.Item("TOTAL_AMOUNT").Value
                                                                                    adoconn.Execute "Update ManagementFeePreview set AgrPercentage=" & percentageOramount1 & ",VATPercentage=0 where MgtFeeAmtTotal is null"
                                                                              End If
                                                                              adoconn.Execute "Update ManagementFeePreview set MgtFeeAmt=round(AgrPercentage*ReceiptAmount/100,3),ChargingMethod='Receipt Basis' where ChargingMethod is null AND MgtFeeAmtTotal is null"
                                                                              adoconn.Execute "Update ManagementFeePreview set MgtFeeAmtTotal=round(MgtFeeAmt*(1+VATPercentage/100),3),VAT=round(MgtFeeAmt*(VATPercentage/100),3),ReceiptAmount=" & dblOriginalAmount & " where ChargingMethod='Receipt Basis' AND MgtFeeAmtTotal is null"
                                                                              .Update
                                                                       End With
                                                                    adoPISplit.Close
                                                                    dblGrandTotal = dblGrandTotal + dblTotalAmount
                                                                    rsfixedMethodDetails.MoveNext 'Looping in all receipt
                                                            Wend
                                                            rsfixedMethod.Close
                                End If 'end of  If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_ED" Then
                               
                               ' dblTotalAmount = dblGrandTotal
                                szSQL = "SELECT * FROM tblPurInvPreview"
                                    'dblTotalAmount = Format(dblGrandTotal, "00000.00")
                                    dblTotalAmount = Round(dblTotalAmount, 2)
                                    If dblTotalAmount = 0 Then
'                                           If adoConn.State = 1 Then
'                                                adoConn.Close
'                                           End If
                                           GoTo EndOfChargeType
                                    End If
                                    
                                    With adoPIHeader
                                            .Open szSQL, adoconn, adOpenDynamic, adLockPessimistic
                                            .AddNew
                                            .Fields.Item("MY_ID").Value = szMYID
                                            .Fields.Item("SlNumber").Value = lSlNumber
                                            .Fields.Item("SUPP_AC").Value = Trim(szManagingAgent(iManagingAgentCount))
                                            .Fields.Item("TRAN_DATE").Value = Format(txtChargeDate.text, "DD MMMM YYYY")
                                            .Fields.Item("TransactionType").Value = 6
                                            .Fields.Item("INV_NO").Value = szPropertySelection1 + "-" + "MFee" + "-" + CStr(lngMgtFeeSL)
                                             lngMgtFeeSL = lngMgtFeeSL + 1
                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
                                            .Fields.Item("History").Value = False
                                            .Fields.Item("TrfPayment").Value = False
                                            .Fields.Item("PropertyID").Value = ""
                                            .Fields.Item("CL_ID").Value = szSelectedClient
                                            .Fields.Item("NLPost").Value = False
                                            .Fields.Item("DueDate").Value = Format(dtNDDInitial, "DD MMMM YYYY")
                                            .Fields.Item("PostingDate").Value = Format(txtPostingDate.text, "DD MMMM YYYY")
                                            .Fields.Item("ReportFromDatePreview").Value = IIf(strFromDate = "", Null, strFromDate)
                                            .Fields.Item("ReportToDatePreview").Value = IIf(strToDate = "", Null, strToDate)
                                            .Fields.Item("DescriptionANDDates").Value = descriptionAndDate
                                            
                                            .Update
                                            iCountPI = iCountPI + 1
                                            lSlNumber = lSlNumber + 1
                                   End With
                                   adoPIHeader.Close
                                    
                                
EndOfChargeType:
            rsCharge.MoveNext
            j = j + 1
            
      Wend


                        rsCharge.Close
                        Set rsCharge = Nothing
                        'adoConn.CommitTrans
                        Dim rsManagementFeePreview As New ADODB.Recordset
                        Dim previousSL As String
                        Dim SL As Integer
                        SL = 1
                        'Exit Sub
                        rsManagementFeePreview.Open "select * from ManagementFeePreview order by PI_ActualID", adoconn, adOpenDynamic, adLockOptimistic
                        While Not rsManagementFeePreview.EOF
                            If previousSL <> rsManagementFeePreview("PI_ActualID").Value Then
                                    previousSL = rsManagementFeePreview("PI_ActualID").Value
                                    SL = 1
                            End If
                            rsManagementFeePreview!ReceiptSLNumber = SL
                            rsManagementFeePreview.Update
                            'adoConn.Execute "Update ManagementFeePreview set ReceiptSLNumber=" & SL & " where PI_ActualID='" & rsManagementFeePreview("PI_ActualID").Value & "'"
                            SL = SL + 1
                            rsManagementFeePreview.MoveNext
                        Wend
                        adoconn.Close
                        Set adoconn = Nothing
EndOfOneManagingAgentforOneAgreement:
           Next iManagingAgentCount
        End If 'end if for 'X' in grid selection
EndOfAgreement:
            Next iPropertyCount
            End If
            Next iClientCount
           Dim StrPropertyCol As String
           If dicWarningProp.Count > 0 Then
                Dim oProp
                For Each oProp In dicWarningProp.Items
                    StrPropertyCol = StrPropertyCol & oProp & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid Client Global Settings setup for the property: " & vbCrLf & StrPropertyCol, vbInformation, "Client Global Data setup"
                warning1 = "Please enter a valid Client Global Settings setup for the property: " & vbCrLf & StrPropertyCol
           End If
           '
           
           StrPropertyCol = ""
           If dicWarningAgreement.Count > 0 Then
                Dim oWar
                For Each oWar In dicWarningAgreement.Items
                    StrPropertyCol = StrPropertyCol & oWar & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                MsgBox "Please enter a valid setup for the property: " & vbCrLf & StrPropertyCol, vbInformation, "Client agreement Fees and charges setup"
                warning2 = "Please enter a valid setup for the property: " & vbCrLf & StrPropertyCol
           End If
           StrPropertyCol = ""
           If dicWarningFinPeriod.Count > 0 Then
                Dim oWarning
                For Each oWarning In dicWarningFinPeriod.Items
                    StrPropertyCol = StrPropertyCol & oWarning & ", "
                Next
           End If
           If Len(StrPropertyCol) > 0 Then
                StrPropertyCol = Left(StrPropertyCol, Len(StrPropertyCol) - 2)
                 'MsgBox "The posting date does not fall in any existing financial period :" & szSelectedClient, vbInformation, "Warning"
                MsgBox "The posting date does not fall in any existing financial period for following Client(s): " & vbCrLf & StrPropertyCol, vbInformation, "Warning"
                warning3 = "The posting date does not fall in any existing financial period for following Client(s): " & vbCrLf & StrPropertyCol
           End If
          
           
            MsgBox iCountPI & " Management Fee Invoice(s) Preview generated.", vbInformation, "Generate Fee Preview"
'            If iCountPI > 0 Then
'                If adoConn.State = 0 Then
'                    adoConn.Open getConnectionString
'                 End If
'                Call frmManagementFees.LoadFlxPurchase(adoConn)
'                adoConn.Close
'            End If
'            frmManagementFees.fmeLoading.Visible = False
'            frmManagementFees.fmeLoading.Refresh
            
            
            
            
            
          If iCountPI > 0 Then
                        Sleep (100)
                        Dim reportApp As New CRAXDRT.Application
                        Dim Report As CRAXDRT.Report
                        Dim rep As frmReport
                        Set Report = reportApp.OpenReport(App.Path & szReportPath & "\" & "ManagementFeePreview.rpt")
                        Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
                        Report.DiscardSavedData
                        Report.EnableParameterPrompting = False
                        
                        Report.ParameterFields(1).AddCurrentValue szSelectedClient
                         Report.ParameterFields(2).AddCurrentValue Replace(szPropertySelectionALL, ",", "   ")
                         Report.ParameterFields(3).AddCurrentValue warning1
                         Report.ParameterFields(4).AddCurrentValue warning2
                         Report.ParameterFields(5).AddCurrentValue warning3
                        
                        Set rep = New frmReport
                        Load rep
                        rep.LoadReportViewer Report
                        
        End If
   
End Sub

Private Function AllSelProperties() As String
   Dim iRow As Integer

   AllSelProperties = ""

   For iRow = 1 To flxProperties.Rows - 1
      If flxProperties.TextMatrix(iRow, 0) = "X" Then
         AllSelProperties = "'" & flxProperties.TextMatrix(iRow, 1) & "', " & AllSelProperties
      End If
   Next iRow
   If Len(AllSelProperties) > 0 Then AllSelProperties = Left(AllSelProperties, Len(AllSelProperties) - 2)
End Function










Private Function SendDemandByE_Mail(adoconn As ADODB.Connection) As Boolean
   Dim i As Integer
   Dim szSub As String, szBody As String

   For i = 0 To iLes - 1
      szSub = "Rent and Service Charge demands from " & uLessee(i).szClient
      szBody = "Please find attachment your rent and service charge demands for payment." & (Chr(13) + Chr(10)) & _
               (Chr(13) + Chr(10)) & _
               "Yours sincerely," & (Chr(13) + Chr(10)) & _
               (Chr(13) + Chr(10)) & _
               uLessee(i).szClient

      SendDemandByE_Mail = SendEmail(szFromEmail, Trim(uLessee(i).szLesseeEmail), _
                                     szSub, _
                                     szBody, , , _
                                     uLessee(i).colAtt, uLessee(i).szLesseeID, "SI")

      adoconn.Execute "UPDATE DemandRecords " & _
                      "SET SentByEmail =" & IIf(SendDemandByE_Mail, "1", "0") & " " & _
                      "WHERE DemandID IN (" & ListDemandID(i) & ");"
   Next i
End Function

Private Function ListDemandID(i As Integer) As String
   Dim j As Integer

   For j = 1 To uLessee(i).colURN.Count
      ListDemandID = ListDemandID & CStr(uLessee(i).colURN(j)) & ", "
   Next j

   ListDemandID = Left(ListDemandID, Len(ListDemandID) - 2)
End Function

Private Function ReceiverEmailList(szLessee As String, szEmail As String, lURN As Long, Optional szClient As String) As Integer
   Dim i As Integer

   ReceiverEmailList = 0
   If iLes = 0 Then
      iLes = 1
      ReDim uLessee(0) As SendDemandByEmail
      Set uLessee(0).colURN = New Collection

      uLessee(0).szLesseeID = szLessee
      uLessee(0).szLesseeEmail = szEmail
      uLessee(0).szClient = szClient
      uLessee(0).colURN.Add lURN
      Exit Function
   End If

   For i = 0 To UBound(uLessee)
      If uLessee(i).szLesseeID = szLessee Then
         uLessee(i).colURN.Add lURN
         Exit For
      End If
   Next i

   If i > UBound(uLessee) Then
      ReDim Preserve uLessee(iLes) As SendDemandByEmail
      Set uLessee(iLes).colURN = New Collection

      uLessee(iLes).szLesseeID = szLessee
      uLessee(iLes).szLesseeEmail = szEmail
      uLessee(iLes).szClient = szClient
      uLessee(iLes).colURN.Add lURN
      ReceiverEmailList = iLes
      iLes = iLes + 1
   End If
End Function

Private Sub SaveAttachment(szFile As String, szLessee As String)
   Dim i As Integer

   On Error GoTo DeclareArray
   For i = 0 To iLes - 1
      If uLessee(i).szLesseeID = szLessee Then
         uLessee(i).colAtt.Add szFile
         Exit For
      End If
   Next i
   Exit Sub

DeclareArray:
   Set uLessee(i).colAtt = New Collection
   uLessee(i).colAtt.Add szFile
End Sub

Private Sub flxClients_RowColChange()
'   SelectFlxGridRow 0, flxClients, flxClients.row
'        txtChargeDate.text = Format(Date, "dd/MM/yyyy")
   Exit Sub
'   SelectOnly1RowFlxGrid flxClients, flxClients.row, 0
'   iSelClient = 1
'
'   FilterProperties
'   flxDemandTypes.Clear
'
'   Dim rCount As Integer
'   Dim adoConn1 As New ADODB.Connection
'   Dim szSelectedClient As String
'   Dim szSelectedClientName As String
'   Dim PurchaseLedgerControl As String
'   For rCount = 1 To flxClients.Rows - 1
'         If flxClients.TextMatrix(rCount, 0) = "X" Then
'            szSelectedClient = flxClients.TextMatrix(rCount, 1)
'            szSelectedClientName = flxClients.TextMatrix(rCount, 2)
'            Exit For
'         End If
'   Next
'
'
''   If szSelectedClient = "" Then Exit Sub
'
'   If adoConn1.State = 0 Then
'        adoConn1.Open getConnectionString
'   End If
'
'    PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Management Fee Payable (P&L)", szSelectedClient)
'    If (PurchaseLedgerControl = "") Then
'        MsgBox "Please set up Management Fees control accounts for '" & szSelectedClientName & "' "
'        Exit Sub
'    End If
'
'
'    PurchaseLedgerControl = GetNominalCodeForControlAccount(adoConn1, "Management Fees Control Account (B/S)", szSelectedClient)
'    If (PurchaseLedgerControl = "") Then
'        MsgBox "Please set up Management Fees control accounts for '" & szSelectedClientName & "' "
'        Exit Sub
'    End If
    
End Sub

Private Sub flxDemandTypes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   flxDemandTypes.MousePointer = vbArrow
End Sub

Private Sub ConfigFlxGrids()
   Dim szHeader As String

   flxClients.RowHeight(0) = 0
   flxClients.ColWidth(0) = 300
   flxClients.ColWidth(1) = 1300
    flxClients.ColAlignment(1) = vbLeftJustify
   flxClients.ColWidth(2) = 2350
   flxClients.ColAlignment(2) = vbLeftJustify

   szHeader$ = "<|<|<|<"
   flxProperties.FormatString = szHeader
   flxProperties.Cols = 4
   flxProperties.RowHeight(0) = 0
   flxProperties.ColWidth(0) = 300                   '"X"
   flxProperties.ColWidth(1) = 1200                'Property ID
   flxProperties.ColAlignment(1) = vbLeftJustify
   flxProperties.ColWidth(2) = 2350                'Property Name
   flxProperties.ColAlignment(2) = vbLeftJustify
   flxProperties.ColWidth(3) = 0                   'Client ID

   flxDemandTypes.Cols = 5
   flxDemandTypes.RowHeight(0) = 0
   flxDemandTypes.ColWidth(0) = 300                  '"X"
   flxDemandTypes.ColWidth(1) = 900                  'Property ID
   flxDemandTypes.ColWidth(2) = 0               'Demand Type ID
    flxDemandTypes.ColAlignment(1) = vbLeftJustify
   flxDemandTypes.ColAlignment(2) = vbLeftJustify
   flxDemandTypes.ColWidth(3) = 4000               'Demand Type Name
   flxDemandTypes.ColWidth(4) = 0                  'Demand Category

   flxCategory.RowHeight(0) = 0
   flxCategory.ColWidth(0) = 0
   flxCategory.ColWidth(1) = 0
   flxCategory.ColWidth(2) = flxCategory.Width - 250
End Sub

Private Sub flxDemandTypes_RowColChange()
'   SelectFlxGridRow 0, flxDemandTypes, flxDemandTypes.row
'   If flxDemandTypes.TextMatrix(flxDemandTypes.row, 0) = "X" Then
'      iSelDemandTypes = iSelDemandTypes + 1
'   Else
'      iSelDemandTypes = iSelDemandTypes - 1
'   End If
'Debug.Print iSelDemandTypes
End Sub

Private Sub flxProperties_RowColChange()
'   SelectFlxGridRow 0, flxProperties, flxProperties.row
'   If flxProperties.TextMatrix(flxProperties.row, 0) = "X" Then
'      iSelProperties = iSelProperties + 1
'   Else
'      iSelProperties = iSelProperties - 1
'   End If
'
'   FilterDemandTypes
End Sub

Private Sub Form_Load()
   Dim szChoice As String

'   Me.Height = 8550
'   Me.Width = 7710
'   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   
   chkProp.BackColor = Me.BackColor
   ChkAllClient.BackColor = Me.BackColor
   chkDT.BackColor = Me.BackColor
   chkAssignProperty.BackColor = Me.BackColor
   
   
   

   ConfigFlxGrids          'flxClients, flxProperties, flxDemandTypes, flxCategory
   LoadProperyIDAndClientID ''flxClients, flxProperties,
   'LoadFlxGrids            ' flxDemandTypes, flxCategory, cboGDPFreq

   iSelClient = 0

   txtChargeDate.text = Format(Date, "dd/mm/yyyy")
   txtPostingDate.text = txtChargeDate.text
   If Not frmMMain.IsRibbonVersion Then
      txtPostingDate.Visible = False
      Label19(0).Visible = False
   End If

   Call WheelHook(Me.hWnd)
End Sub
Private Sub LoadProperyIDAndClientID()
    Dim szSQL As String, r As Integer
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

'   connect to database
   adoconn.Open getConnectionString

   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT;"
   adoRst.Open szSQL, adoconn, adOpenDynamic, adLockPessimistic

   r = 1
   flxClients.Rows = 1

   While Not adoRst.EOF
      flxClients.AddItem ""
      flxClients.TextMatrix(r, 1) = adoRst.Fields.Item("CLIENTID").Value
      flxClients.TextMatrix(r, 2) = adoRst.Fields.Item("CLIENTNAME").Value
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
'------------------------------------------------------------------------------------------
   szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY " & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   r = 0
   ReDim szaProp(adoRst.RecordCount) As String

   While Not adoRst.EOF
      szaProp(r) = adoRst.Fields.Item("PROPERTYID").Value & " ## " & _
                   adoRst.Fields.Item("PROPERTYNAME").Value & " ## " & _
                   adoRst.Fields.Item("ClientID").Value
      r = r + 1
      adoRst.MoveNext
   Wend
   iProperty = r

   adoRst.Close
   adoconn.Close
   
End Sub
Private Sub LoadFlxGrids()
   Dim szSQL As String, r As Integer
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

'   connect to database
   adoconn.Open getConnectionString

'------------------------------------------------------------------------------------------
'Put demand type here

    Dim i As Integer
    Dim strSelProperties  As String
    For i = 1 To flxProperties.Rows - 1
        If flxProperties.TextMatrix(i, 0) = "X" Then
            strSelProperties = strSelProperties & "'" & flxProperties.TextMatrix(i, 1) & "',"
        End If
    Next i
    If Len(strSelProperties) > 0 Then
        strSelProperties = Left(strSelProperties, Len(strSelProperties) - 1)
    End If
    
'   szSQL = "SELECT PropertyID, ID, Type, CategoryCode " & _
'           "FROM   DemandTypes where PropertyID in (" & strSelProperties & ")" & _
'           "ORDER BY ID;"
' Distinct added on 20221128
    szSQL = "SELECT Distinct F.FundCode, F.FundID, F.FundName,F.CategoryCode FROM tlbAgreement agr, ClientProAgr CPA, ChargeTypes C,Fund  F,SECONDARYCODE SC " & _
            "WHERE agr.CPA_ID = CPA.CPA_ID And F.FundID=agr.fund AND SC.CODE=agr.CHARGE_METHOD AND " & _
            "C.ID = agr.CHARGE_TYPE And " & _
            "CPA.PropertyID in (" & strSelProperties & ");"
'
'    szSQL = "SELECT D.PropertyID, D.ID, D.Type, D.CategoryCode  FROM tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,SECONDARYCODE SC,tlbPayable P " & _
'            "WHERE P.ClientID=CPA.ClientID AND P.PAY_FUND=agr.Fund AND P.CPA_ID=CPA.CPA_ID AND agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND SC.CODE=agr.CHARGE_METHOD AND " & _
'            "D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'            "CPA.PropertyID in (" & strSelProperties & ");"
            
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   r = 0
   ReDim szaDT(adoRst.RecordCount) As String

   While Not adoRst.EOF
      szaDT(r) = adoRst.Fields.Item(0).Value & " ## " & _
                 adoRst.Fields.Item(1).Value & " ## " & _
                 adoRst.Fields.Item(2).Value & " ## " & _
                 adoRst.Fields.Item(3).Value
      r = r + 1
      adoRst.MoveNext
   Wend
   iDT = r

   adoRst.Close
   
'Put charge type here
'   szSQL = "SELECT  PropertyID,ID, FeeType, CategoryCode " & _
'           "FROM   ChargeTypes " & _
'           "ORDER BY ID;"
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   r = 0
'   ReDim szaDT(adoRst.RecordCount) As String
'
'   While Not adoRst.EOF
'      szaDT(r) = adoRst.Fields.Item(0).Value & " ## " & _
'                 adoRst.Fields.Item(1).Value & " ## " & _
'                 adoRst.Fields.Item(2).Value & " ## " & _
'                 adoRst.Fields.Item(3).Value
'      r = r + 1
'      adoRst.MoveNext
'   Wend
'   iDT = r
'
'   adoRst.Close
'------------------------------------------------------------------------------------------
   szSQL = "SELECT Code, Value " & _
           "FROM SecondaryCode " & _
           "WHERE PrimaryCode = 'DCTG' " & _
           "ORDER BY Code;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   r = 1
   flxCategory.Rows = 1

   While Not adoRst.EOF
      flxCategory.AddItem ""
      flxCategory.TextMatrix(r, 1) = adoRst.Fields.Item(0).Value
      flxCategory.TextMatrix(r, 2) = adoRst.Fields.Item(1).Value
      r = r + 1
      adoRst.MoveNext
   Wend

   adoRst.Close
   For r = 1 To flxCategory.Rows - 1
      SelectFlxGridRow 0, flxCategory, r
   Next r
   iSelDemandCategory = r - 1
'------------------------------------------------------------------------------------------
'   FillFrequency adoConn

   Set adoRst = Nothing
   adoconn.Close
   Set adoconn = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub FilterProperties()
   Dim i As Integer, j As Integer, r As Integer
   Dim szaTemp() As String

   flxProperties.Clear
   iSelProperties = 0
   chkProp.Value = 0
   chkDT.Value = 0

   flxProperties.Rows = 1
   r = 1
   For i = 0 To iProperty - 1
      szaTemp = Split(szaProp(i), " ## ")

      For j = 0 To flxClients.Rows - 1
         If flxClients.TextMatrix(j, 0) = "X" And flxClients.TextMatrix(j, 1) = szaTemp(2) Then
            flxProperties.AddItem ""
            flxProperties.TextMatrix(r, 1) = szaTemp(0)
            flxProperties.TextMatrix(r, 2) = szaTemp(1)
            flxProperties.TextMatrix(r, 3) = szaTemp(2)
            'added by anol 16 Aug 2016
            flxProperties.row = 1
            r = r + 1
         End If
      Next j
   Next i
End Sub

Private Sub flxClients_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxClients.MousePointer = vbArrow
End Sub

Private Sub FilterDemandTypes()
   Dim i As Integer, j As Integer, r As Integer
   Dim szaTemp() As String

   flxDemandTypes.Clear
   chkDT.Value = 0

   flxDemandTypes.Rows = 1
   r = 1
   For i = 0 To iDT - 1
      szaTemp = Split(szaDT(i), " ## ")

      'For j = 1 To flxProperties.Rows - 1
         'If flxProperties.TextMatrix(j, 0) = "X" And (flxProperties.TextMatrix(j, 1) = szaTemp(0) Or szaTemp(0) = "ALL") Then
         'If flxProperties.TextMatrix(j, 0) = "X" Then
            If DemandCategorySelected(CInt(szaTemp(3))) Then 'This will be fund category
               flxDemandTypes.AddItem ""
               flxDemandTypes.TextMatrix(r, 1) = szaTemp(0) 'Property ID
               flxDemandTypes.TextMatrix(r, 2) = szaTemp(1)
               flxDemandTypes.TextMatrix(r, 3) = szaTemp(2)
               flxDemandTypes.TextMatrix(r, 4) = szaTemp(3)
               'added by anol 16 Aug 2016
               flxDemandTypes.row = 1
               r = r + 1
               If szaTemp(0) = "ALL" Then Exit For
            End If
         'End If
     ' Next j
   Next i
End Sub

Private Function DemandCategorySelected(iCat As Integer) As Boolean
   Dim i As Integer

   DemandCategorySelected = False

   For i = 1 To flxCategory.Rows - 1
      If flxCategory.TextMatrix(i, 0) = "X" And flxCategory.TextMatrix(i, 1) = iCat Then
         DemandCategorySelected = True
         Exit For
      End If
   Next i
End Function

'Public Sub FillFrequency(Conn1 As ADODB.Connection)
'   Dim i As Integer, SQLStr1 As String
'   Dim Rst1 As New ADODB.Recordset
'   Dim Data() As String
'
'   SQLStr1 = "SELECT * FROM Frequencies"
'   Rst1.Open SQLStr1, Conn1, adOpenStatic, adLockReadOnly
'
'   If Not Rst1.EOF Then
'      ReDim Preserve Data(1, Rst1.RecordCount) As String
'      Data(0, 0) = 0
'      Data(1, 0) = "ALL Frequencies"
'      i = 1
'      While Rst1.EOF = False
'         Data(0, i) = Rst1!Id
'         Data(1, i) = Rst1!Frequency
'         i = i + 1
'         Rst1.MoveNext
'      Wend
'
'
'      cboGDPFreq.Column() = Data()
'      cboGDPFreq.ListIndex = 0
'   End If
'
'   Rst1.Close
'   Set Rst1 = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmDemands3.Enabled = True
   UnLoadForm Me
'   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtPostingDate_Change()
   TextBoxChangeDate txtPostingDate
End Sub

Private Sub txtPostingDate_GotFocus()
   SelTxtInCtrl txtPostingDate
End Sub

Private Sub txtPostingDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtPostingDate, KeyAscii
End Sub

Private Sub txtPostingDate_LostFocus()
   TextBoxFormatDate txtPostingDate
End Sub
'************
Private Sub txtLastChargeDate_Change()
   TextBoxChangeDate txtLastChargeDate
End Sub

Private Sub txtLastChargeDate_GotFocus()
   SelTxtInCtrl txtLastChargeDate
End Sub

Private Sub txtLastChargeDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtLastChargeDate, KeyAscii
End Sub

Private Sub txtLastChargeDate_LostFocus()
   TextBoxFormatDate txtLastChargeDate
End Sub
'**************
Private Sub txtSCDateFrom_Change()
'   TextBoxChangeDate txtSCDateFrom
End Sub

Private Sub txtSCDateFrom_GotFocus()
'   SelTxtInCtrl txtSCDateFrom
End Sub

Private Sub txtSCDateFrom_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtSCDateFrom, KeyAscii
End Sub

Private Sub txtSCDateFrom_LostFocus()
'   TextBoxFormatDate txtSCDateFrom
End Sub

Private Sub txtSCDateTo_Change()
'   TextBoxChangeDate txtSCDateTo
End Sub

Private Sub txtSCDateTo_GotFocus()
'   SelTxtInCtrl txtSCDateTo
End Sub

Private Sub txtSCDateTo_KeyPress(KeyAscii As Integer)
'   TextBoxKeyPrsDate txtSCDateTo, KeyAscii
End Sub

Private Sub txtSCDateTo_LostFocus()
'   TextBoxFormatDate txtSCDateTo
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

'            If dblTotalAmount = 0 Then
'                    rsCharge.Close
'                    Set rsCharge = Nothing
'                    adoConn.Close
'                    Set adoConn = Nothing
'                    MsgBox "Total Amount cannot be zero"
'            End If
'            If rsCharge.Fields.Item("agreementEndDate").Value < rsCharge.Fields.Item("FDD").Value Then
'                    rsCharge.Close
'                    Set rsCharge = Nothing
'                    adoConn.Close
'                    Set adoConn = Nothing
'                    MsgBox "agreement End Date is greatar than following due date"
'                    Exit Sub
'            End If
'            adoConn.Open getConnectionString
'            adoConn.BeginTrans
'            If rsCharge.Fields.Item("CHARGE_METHOD").Value = "RE_FIX" Then   'when working on the fixed procedure only 1 line of setup is done then
'                    If dblTotalAmount = 0 Then
'                            rsCharge.Close
'                            Set rsCharge = Nothing
'                            adoConn.Close
'                            Set adoConn = Nothing
'                            MsgBox "Total Amount cannot be zero"
'                    End If
'                    If rsCharge.Fields.Item("agreementEndDate").Value < rsCharge.Fields.Item("FDD").Value Then
'                            rsCharge.Close
'                            Set rsCharge = Nothing
'                            adoConn.Close
'                            Set adoConn = Nothing
'                            MsgBox "agreement End Date is greatar than following due date"
'                            Exit Sub
'                    End If
'                                  szMYID = UniqueID()
'
'                                  szSQL = "SELECT * FROM tblPurInv"
'
'                                  lSlNumber = SlNumber("PI", "tblPurInv", adoConn)
'                                  With adoPIHeader
'                                            .Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'                                            .AddNew
'                                            .Fields.Item("MY_ID").Value = szMYID
'                                            .Fields.Item("SlNumber").Value = lSlNumber
'                                            .Fields.Item("SUPP_AC").Value = szSelectedClient
'                                            .Fields.Item("TRAN_DATE").Value = Format(txtChargeDate.text, "DD MMMM YYYY")
'                                            .Fields.Item("TransactionType").Value = 6
'                                            .Fields.Item("INV_NO").Value = "Mngment Fee"
'                                            .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
'                                            .Fields.Item("History").Value = False
'                                            .Fields.Item("TrfPayment").Value = False
'                                            .Fields.Item("PropertyID").Value = ""
'                                            .Fields.Item("CL_ID").Value = szSelectedClient
'                                            .Fields.Item("NLPost").Value = False
'                                            .Fields.Item("DueDate").Value = Format(txtChargeDate.text, "DD MMMM YYYY")
'                                            .Fields.Item("PostingDate").Value = Format(txtPostingDate.text, "DD MMMM YYYY")
'                                            .Fields.Item("isManagementFee").Value = True
'                                            .Update
'                                            'Exit Sub
'                                  End With
'
'
'                            '      ************************************Write tblPurInvSRec **************************************
'
'                                  szSQL = "SELECT * FROM tblPurInvSRec"
'                                  adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'                                    'Add New Records. At least there is only one split line
'                                     With adoPISplit
'                                         .AddNew
'                                         .Fields.Item("MY_ID").Value = UniqueID()
'                                         .Fields.Item("ParentID").Value = szMYID
'                                         .Fields.Item("TRAN_ID").Value = "1"
'                                         .Fields.Item("TRANS").Value = szPropertySelection1  ' If you select One property then you can write a value here
'                                         .Fields.Item("UNIT_ID").Value = ""
'                                         .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
'                                         .Fields.Item("DEPT_ID").Value = dblFundId
'                                        ' .Fields.Item("JOB_ID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
'                                         .Fields.Item("RecoverablePt").Value = 0
'                                         .Fields.Item("description").Value = "Mngment Fee"
'                                         .Fields.Item("NET_AMOUNT").Value = dblTotalAmount
'                                         .Fields.Item("TAX_CODE").Value = Null
'                                         .Fields.Item("VAT").Value = 0
'                                         .Fields.Item("TOTAL_AMOUNT").Value = dblTotalAmount
'                                         .Update
'                                     End With
'                                  adoPIHeader.Close
'                                  'Exit Sub
'                                   '*************************************Write tlbPayment **************************************
'                                  Dim lT_ID As Long
'                                  szSQL = "SELECT MAX(TRANSACTIONID) AS TID FROM tlbPayment;"
'                                  adoPIHeader.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'                                  lT_ID = CLng(IIf(IsNull(adoPIHeader!TID), 1, adoPIHeader!TID + 1))
'                                  adoPIHeader.Close
'
'
'
'                                   szSQL = "SELECT * FROM tlbPayment where 1=2"
'                                   With adoPIHeader
'                                          .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic 'Add New Mode
'                                          .AddNew
'                                          !transactionID = lT_ID
'                                          !Type = 6  'PP - Purchase Invoice, look in the tlbTransactionType
'                                          !SageAccountNumber = szSelectedClient
'                                          !Pi = szMYID
'                                          !PDate = Format(txtChargeDate.text, "DD MMMM YYYY")
'                                          !dDate = Format(txtChargeDate.text, "DD MMMM YYYY")
'                                          !ref = "Mngment Fee"
'                                          !ExtRef = "Mngment Fee"
'                                          !amount = dblTotalAmount
'                                          !OSAmount = !amount
'                                          !PaymentView = True
'                                          !Details = "Mngment Fee"
'                                          '!unitid = txtProperty.text
'                                          !SlNumber = lSlNumber
'                                          !fundID = dblFundId
'                            '              !AdjTag = IIf(bAdjustment, "Y", "N")
'                                          !Recoverable = 0
'                                          !postingDate = Format(txtPostingDate.text, "DD MMMM YYYY")
'                                          !clientID = szSelectedClient
'                                          .Update
'                                          .Close
'                                 End With
'                               '*************************************Write tlbPaymentSplit **************************************
'                               adoPISplit.Close
'                               Set adoPISplit = Nothing
'                               szSQL = "SELECT * FROM tlbPaymentSplit;"
'                               adoPISplit.Open szSQL, adoConn, adOpenDynamic, adLockPessimistic
'                            'Add New Records. At least there is one split line.
'                                     With adoPISplit
'                                        .AddNew
'                                        .Fields.Item("TransactionID").Value = UniqueID()
'                                        .Fields.Item("PayHeader").Value = lT_ID
'                                        .Fields.Item("FundID").Value = dblFundId
'                                        .Fields.Item("Amount").Value = dblTotalAmount
'                                        .Fields.Item("OSAmount").Value = dblTotalAmount
'                                        .Fields.Item("SplitID").Value = "1"
'                                        .Fields.Item("DueDate").Value = Format(txtChargeDate.text, "DD MMMM YYYY")
'                                        .Fields.Item("Description").Value = "Mngment Fee"
'                                        '.Fields.Item("JobID").Value = flxPI.TextMatrix(iRow, 9)            'Job No
'                                        .Fields.Item("NOMINAL_CODE").Value = FinalControlACForPayable
'                                        .Fields.Item("TRANS").Value = "" 'Put property ID  Later
'                                        .Fields.Item("UNIT_ID").Value = "" '
'                                        .Fields.Item("ScheduleID").Value = Null
'                                        .Update
'                                     End With
'
'                                   dtNextDue = NextDueDate1(dblFeqID, dtNextDue, szPropertySelection1)
'                                   dtFDD = NextDueDate1(dblFeqID, dtNextDue, szPropertySelection1)
'                               'Update FDD and Next due date
'                                szSQL = "Update tlbAgreement agr, ClientProAgr CPA, DemandTypes D, ChargeTypes C,Fund  F,FREQUENCIES FC,SECONDARYCODE SC " & _
'                                " Set NtDueDate=#" & dtNextDue & "# ,FDD=#" & dtFDD & "# " & _
'                        "WHERE agr.CPA_ID = CPA.CPA_ID And cstr(F.FundID)=agr.fund AND FC.ID=agr.Frequency AND SC.CODE=agr.CHARGE_METHOD AND " & _
'                        "CPA.ClientID = '" & szSelectedClient & "' And D.ID = agr.DEMAND_TYPE And C.ID = agr.CHARGE_TYPE And " & _
'                        "CPA.PropertyID = '" & szPropertySelection1 & "';"
'
'                         adoConn.Execute szSQL
'            End If 'end of for  fixed method  when only 1 line of setup is done
