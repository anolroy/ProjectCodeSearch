VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLLFeesCharges1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Landlord/Client Fees and Charges"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Myriad Condensed Web"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLLFeesCharges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11655
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   1755
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3096
      _Version        =   393216
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton Command8 
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10320
      TabIndex        =   5
      Top             =   7400
      Width           =   1215
   End
   Begin TabDlg.SSTab tabFees 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   2625
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Recharge"
      TabPicture(0)   =   "frmLLFeesCharges.frx":1982
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRechargeConfirmed"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxRecharge"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtSetRechargeDate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdRechargeGenerate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Agent Fees"
      TabPicture(1)   =   "frmLLFeesCharges.frx":199E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSHFlexGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Management Fee"
      TabPicture(2)   =   "frmLLFeesCharges.frx":19BA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "MSHFlexGrid3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Rent Payable"
      TabPicture(3)   =   "frmLLFeesCharges.frx":19D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSHFlexGrid4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Select an option:"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   6375
         Begin MSForms.TextBox txtRechrgeStDt 
            Height          =   285
            Left            =   3720
            TabIndex        =   18
            Top             =   300
            Width           =   1095
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "1931;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optReAllExp 
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2040
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3598;529"
            Value           =   "1"
            Caption         =   "Recharge All Expenses"
            FontName        =   "Myriad Condensed Web"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optDateRange 
            Height          =   375
            Left            =   2520
            TabIndex        =   17
            Top             =   240
            Width           =   1215
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "2143;661"
            Value           =   "0"
            Caption         =   "Date Range:"
            FontName        =   "Myriad Condensed Web"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtRechrgeEndDt 
            Height          =   285
            Left            =   5160
            TabIndex        =   19
            Top             =   300
            Width           =   1095
            VariousPropertyBits=   679495711
            BorderStyle     =   1
            Size            =   "1931;503"
            SpecialEffect   =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   225
            Left            =   4905
            TabIndex        =   24
            Top             =   300
            Width           =   165
            VariousPropertyBits=   276824091
            Caption         =   "to"
            Size            =   "291;397"
            FontName        =   "Myriad Condensed Web"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdRechargeConfirmed 
         Caption         =   "Confirm Recharge"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   -65400
         TabIndex        =   22
         Top             =   4240
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Generate Letting Fee"
         BeginProperty Font 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74880
         TabIndex        =   10
         Top             =   4080
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRecharge 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   6
         Top             =   1125
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5424
         _Version        =   393216
         Cols            =   15
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
         _Band(0).Cols   =   15
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   15
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   3675
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   15
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   3675
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6482
         _Version        =   393216
         Cols            =   15
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
         _Band(0).Cols   =   15
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.Label Label5 
         Height          =   255
         Left            =   -68640
         TabIndex        =   25
         Top             =   4290
         Width           =   1575
         Caption         =   "Set Recharge Date:"
         Size            =   "2778;450"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSetRechargeDate 
         Height          =   285
         Left            =   -66960
         TabIndex        =   21
         Top             =   4290
         Width           =   1335
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2355;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdRechargeGenerate 
         Height          =   495
         Left            =   -64800
         TabIndex        =   20
         Top             =   480
         Width           =   1095
         Caption         =   "Generate"
         Size            =   "1931;873"
         FontName        =   "Myriad Condensed Web"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.TextBox txtClientID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   3255
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin MSAdodcLib.Adodc adoMain 
      Height          =   330
      Left            =   3840
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Height          =   2655
      Left            =   7320
      ScaleHeight     =   2625
      ScaleWidth      =   5385
      TabIndex        =   13
      Top             =   0
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
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   2175
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Condensed Web"
            Size            =   9.75
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
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   11655
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   855
      Caption         =   "Properties:"
      Size            =   "1508;450"
      FontName        =   "Myriad Condensed Web"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Client/Landlord ID:"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLLFeesCharges1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szProertyID As String
'Private szUnitID As String
Private bData As Boolean

Private Sub cmdClient_Click()
   Call PrepareList

   picClientList.Top = txtClientID.Top + txtClientID.Height + 5
   picClientList.Left = txtClientID.Left + 5
   picClientList.Visible = True
   picClientList.ZOrder 0
End Sub

Private Sub cmdClient_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then picClientList.Visible = False
End Sub

Private Sub cmdGridUnitLookup_Click()
   picClientList.Visible = False
End Sub

Private Sub cmdRechargeGenerate_Click()

Exit Sub

   Dim iRow As Integer, iSelected As Integer, iCol As Integer
   Dim Conn1 As New rdoConnection
   Dim szBackUp As String
   Dim rstBackUp As rdoResultset

   If optReAllExp.Value Then
      szBackUp = "SELECT tlbRechargePre.* " & _
                 "FROM tlbRechargePre, Units " & _
                 "WHERE tlbRechargePre.UNIT_ID = Units.UnitNumber And " & _
                     "Units.PropertyID = '" & szProertyID & "';"
      Set rstBackUp = Conn1.OpenResultset(szBackUp, rdOpenDynamic, rdConcurRowVer)

      If rstBackUp.EOF Then GoTo ErrorHandler

      iRow = 1
'    View all data from tlbRechargePre which saved backup data
      While Not rstBackUp.EOF
         If rstBackUp!Checked = True Then
            flxRecharge.TextMatrix(iRow, 0) = "X"
            iSelected = iSelected + 1                       'counting selected rows
            flxRecharge.Row = iRow
            For iCol = 1 To flxRecharge.Cols - 1
               flxRecharge.Col = iCol
               flxRecharge.CellBackColor = RGB(180, 255, 180)
            Next iCol
         End If
         flxRecharge.TextMatrix(iRow, 1) = rstBackUp!TRAN_ID
         flxRecharge.TextMatrix(iRow, 2) = rstBackUp!TRAN_DATE
         flxRecharge.TextMatrix(iRow, 3) = rstBackUp!TRAN_TYPE
         flxRecharge.TextMatrix(iRow, 4) = rstBackUp!TRANS
         flxRecharge.TextMatrix(iRow, 5) = rstBackUp!INV_NO
         flxRecharge.TextMatrix(iRow, 6) = rstBackUp!NOMINAL_CODE
         flxRecharge.TextMatrix(iRow, 7) = rstBackUp!DEPT_ID
         flxRecharge.TextMatrix(iRow, 8) = rstBackUp!PROJ_REF
         flxRecharge.TextMatrix(iRow, 9) = rstBackUp!COST_CODE
         flxRecharge.TextMatrix(iRow, 10) = rstBackUp!description
         flxRecharge.TextMatrix(iRow, 11) = Format(rstBackUp!NET_AMOUNT, "0.00")
         flxRecharge.TextMatrix(iRow, 12) = Format(rstBackUp!RECHARGE, "0.00")

         rstBackUp.MoveNext
         If Not rstBackUp.EOF Then flxRecharge.AddItem ""
         iRow = iRow + 1
         bData = True
      Wend
      rstBackUp.Close
      Set rstBackUp = Nothing
   Else
      If txtRechrgeStDt.text = "" Or txtRechrgeEndDt.text = "" Then
         MsgBox "Please enter date range for generate charges.", vbCritical + vbOKOnly, "Date missing error"
         Exit Sub
      End If
      If DateDiff("d", CDate(txtRechrgeStDt.text), CDate(txtRechrgeEndDt.text)) < 0 Then
         MsgBox "Please enter valid date sequence.", vbCritical + vbOKOnly, "Date entry error"
         Exit Sub
      End If
   
      
'   TRAN_DATE
   
   End If
   
   Set Conn1 = Nothing
'****************************************************************************
   Exit Sub

ErrorHandler:
   rstBackUp.Close

   Set rstBackUp = Nothing

   Conn1.Close
   Set Conn1 = Nothing
End Sub

Private Sub DrawRechargeGrid()
   Dim szHeader As String

   flxRecharge.Cols = 13
   flxRecharge.Rows = 2
   flxRecharge.Clear
   flxRecharge.RowHeightMin = 285
   szHeader$ = "|<No.|<Date|<Type|<Trans|<Inv No|<N/C|<Dept|<Proj|<C Code|<Details|>Net|>Recharge"
   flxRecharge.FormatString = szHeader$

   flxRecharge.ColWidth(0) = 250                               'Solid column
   flxRecharge.ColWidth(1) = (flxRecharge.Width - 250) * 0.05  'No
   flxRecharge.ColWidth(2) = (flxRecharge.Width - 250) * 0.09  'Date
   flxRecharge.ColWidth(3) = (flxRecharge.Width - 250) * 0.085 'Type
   flxRecharge.ColWidth(4) = (flxRecharge.Width - 250) * 0.085 'Trans
   flxRecharge.ColWidth(5) = (flxRecharge.Width - 250) * 0.07  'Inv No
   flxRecharge.ColWidth(6) = (flxRecharge.Width - 250) * 0.07  'N/C
   flxRecharge.ColWidth(7) = (flxRecharge.Width - 250) * 0.06  'Dept
   flxRecharge.ColWidth(8) = (flxRecharge.Width - 250) * 0.06  'Proj
   flxRecharge.ColWidth(9) = (flxRecharge.Width - 250) * 0.07  'C Code
   flxRecharge.ColWidth(10) = (flxRecharge.Width - 250) * 0.18 'Details
   flxRecharge.ColWidth(11) = (flxRecharge.Width - 250) * 0.09 'Net
   flxRecharge.ColWidth(12) = (flxRecharge.Width - 250) * 0.09 'Recharge
End Sub

Private Sub flxClientList_Click()
   Dim sSQLQuery_ As String, sFilter As String

   txtClientID.text = flxClientList.TextMatrix(flxClientList.Row, 1)

   MousePointer = vbHourglass
   
   adoMain.ConnectionString = "DSN=" & Adsn & ";UID=;PWD="
   sSQLQuery_ = "SELECT * " & _
                "FROM CLIENT " & _
                "WHERE CLIENT.ClientID = '" & flxClientList.TextMatrix(flxClientList.Row, 1) & "';"

   adoMain.RecordSource = sSQLQuery_
   adoMain.CommandType = adCmdText
   adoMain.Refresh

   If Not Fill_Form(Me, adoMain) Then
      MsgBox "Error in Database.", vbExclamation
   Else
      LoadClientProperty
   End If

   MousePointer = vbDefault

   picClientList.Visible = False
End Sub

Private Sub LoadClientProperty()
   Dim conClient As New RDO.rdoConnection
   Dim rstProperty As rdoResultset
   Dim szSQL As String

   On Error GoTo ErrorHandler

   'Set the RDO Connections to the dataset
   conClient.Connect = "DSN=" & Adsn & ";UID=;PWD="
   conClient.CursorDriver = rdUseIfNeeded
   conClient.EstablishConnection rdDriverNoPrompt

   szSQL = "SELECT PropertyID, PropertyName  " & _
           "FROM PROPERTY " & _
           "WHERE CLIENTID = '" & txtClientID.text & "' " & _
           "ORDER BY PropertyName;"

   Set rstProperty = conClient.OpenResultset(szSQL, rdOpenStatic, rdConcurReadOnly)

   DrawPropertyGrid
   
   If rstProperty.EOF Then GoTo NoRes

   Dim iRow As Integer
   iRow = 1

   While Not rstProperty.EOF
      flxProperties.TextMatrix(iRow, 1) = rstProperty!PROPERTYID
      flxProperties.TextMatrix(iRow, 2) = rstProperty!PropertyName

      rstProperty.MoveNext
      If Not rstProperty.EOF Then flxProperties.AddItem ""
      iRow = iRow + 1
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

Private Sub flxProperties_Click()
   szProertyID = flxProperties.TextMatrix(flxProperties.Row, 1)
End Sub

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50
   tabFees.Tab = 0
   
   DrawPropertyGrid
   
'   txtSetRechargeDate.text = Format(Now, "dd/mm/yyyy")
   DrawRechargeGrid
End Sub

Private Sub DrawPropertyGrid()
   Dim szHeader As String

   flxProperties.Cols = 3
   flxProperties.Rows = 2
   flxProperties.Clear
   szHeader$ = "|<ID|<Propety Name"
   flxProperties.FormatString = szHeader$

   flxProperties.ColWidth(0) = 500                                 'Solid column
   flxProperties.ColWidth(1) = (flxProperties.Width - 800) * 0.3   'Property ID
   flxProperties.ColWidth(2) = (flxProperties.Width - 800) * 0.65  'Property Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub PrepareList()
   FlxDemandsConfigure flxClientList
   LoadAllClientFlxGrd
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

Private Sub OptionButton2_Change()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub optDateRange_Change()
   If optDateRange.Value Then
      txtRechrgeStDt.Locked = False
      txtRechrgeEndDt.Locked = False
      txtRechrgeStDt.SetFocus
   Else
      txtRechrgeStDt.Locked = True
      txtRechrgeEndDt.Locked = True
      txtRechrgeStDt.text = ""
      txtRechrgeEndDt.text = ""
   End If
End Sub

Private Sub txtRechrgeEndDt_Change()
   MsTextBoxChangeDate txtRechrgeEndDt
End Sub

Private Sub txtRechrgeEndDt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate txtRechrgeEndDt, KeyAscii
End Sub

Private Sub txtRechrgeEndDt_LostFocus()
   MsTextBoxFormatDate txtRechrgeEndDt
End Sub

Private Sub txtRechrgeStDt_Change()
   MsTextBoxChangeDate txtRechrgeStDt
End Sub

Private Sub txtRechrgeStDt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate txtRechrgeStDt, KeyAscii
End Sub

Private Sub txtRechrgeStDt_LostFocus()
   MsTextBoxFormatDate txtRechrgeStDt
End Sub

Private Sub txtSetRechargeDate_Change()
   MsTextBoxChangeDate txtSetRechargeDate
End Sub

Private Sub txtSetRechargeDate_KeyPress(KeyAscii As MSForms.ReturnInteger)
   MsTextBoxKeyPrsDate txtSetRechargeDate, KeyAscii
End Sub

Private Sub txtSetRechargeDate_LostFocus()
   MsTextBoxFormatDate txtSetRechargeDate
End Sub
