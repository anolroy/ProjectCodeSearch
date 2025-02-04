VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmailSalesInv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sent Email of Sales Invoice"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailSalesInv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraClientMain 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   29
      Top             =   7320
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame fraClient 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   5055
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1200
            Width           =   1380
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Ok"
            Height          =   375
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1200
            Width           =   1380
         End
         Begin VB.Label lblSupplier 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Name"
            Height          =   195
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   2835
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Please select the client: "
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1875
         End
         Begin MSForms.ComboBox cboClientID 
            Height          =   315
            Left            =   720
            TabIndex        =   32
            Top             =   720
            Width           =   4170
            VariousPropertyBits=   1753237531
            DisplayStyle    =   3
            Size            =   "7355;556"
            BoundColumn     =   0
            TextColumn      =   2
            ColumnCount     =   2
            ListRows        =   20
            cColumnInfo     =   1
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   6
            FontName        =   "Myriad Web"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            Object.Width           =   "1763"
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdBACSEmailTemp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Demand Email &Template"
      Height          =   375
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6240
      Width           =   2100
   End
   Begin VB.CommandButton cmdOpenAttch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Open Attachment"
      Height          =   375
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   1620
   End
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Select All"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5900
      Width           =   1215
   End
   Begin VB.OptionButton optEmail 
      Caption         =   "Failed Email"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optEmail 
      Caption         =   "Sent Email"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   26
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdResend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Resend"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6240
      Width           =   1380
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   375
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6240
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Date Range"
      Height          =   975
      Left            =   12015
      TabIndex        =   15
      Top             =   45
      Width           =   3495
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sort"
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Generate Payment later"
         Top             =   240
         Width           =   960
      End
      Begin VB.TextBox txtFromDate 
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
         Height          =   285
         Left            =   720
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "01/01/2000"
         Top             =   240
         Width           =   1575
      End
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
         Height          =   285
         Left            =   720
         MaxLength       =   10
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   180
      End
   End
   Begin VB.PictureBox picDmdLeaseList 
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
      Height          =   3135
      Left            =   5865
      ScaleHeight     =   3105
      ScaleWidth      =   6345
      TabIndex        =   6
      Top             =   6795
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdDmdGridUnitLookup 
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
         Left            =   6080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   20
         Width           =   255
      End
      Begin VB.TextBox txtDmdTenantSearchID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtDmdTenantSearchName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   300
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDmdLeaseList 
         Height          =   2490
         Left            =   45
         TabIndex        =   10
         Top             =   585
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4392
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   12632256
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   75
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   75
         Width           =   585
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   17
         Left            =   45
         Top             =   70
         Width           =   6015
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmailRemitt 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorSel    =   15329508
      ForeColorSel    =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
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
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtSupplierName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   25
      Text            =   "All Lessees"
      Top             =   120
      Width           =   4280
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   120
      Width           =   525
   End
   Begin MSForms.CommandButton cmdDmdSuppLookup 
      Height          =   285
      Left            =   5010
      TabIndex        =   13
      Top             =   120
      Width           =   255
      Caption         =   """"
      Size            =   "450;512"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lessee"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
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
      Index           =   18
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   12240
   End
End
Attribute VB_Name = "frmEmailSalesInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iKeyUp    As Integer
Private iRowNum   As Integer
Private bEmail    As Boolean

Private Sub chkSelectAll_Click()
   Dim i As Integer

   For i = 1 To flxEmailRemitt.Rows - 1
      If flxEmailRemitt.RowHeight(i) = 240 Then
         flxEmailRemitt.TextMatrix(i, 0) = IIf(chkSelectAll.Value = 1, "X", "")
      End If
   Next i
End Sub

Private Sub cmdBACSEmailTemp_Click()
   Load frmEmailTemplate
   frmEmailTemplate.Caption = "Demand Email Template"
   frmEmailTemplate.Show
End Sub

Private Sub cmdCancel_Click()
   flxEmailRemitt.TextMatrix(iRowNum, 8) = "Cancelled"
   
   fraClientMain.Visible = False
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDmdGridUnitLookup_Click()
   picDmdLeaseList.Visible = False
End Sub

Private Sub cmdDmdSuppLookup_Click()
   txtDmdTenantSearchID.text = ""
   txtDmdTenantSearchID.Left = Label2(0).Left
   txtDmdTenantSearchID.Width = Label2(1).Left - Label2(0).Left
   txtDmdTenantSearchName.text = ""
   txtDmdTenantSearchName.Left = Label2(1).Left
   txtDmdTenantSearchName.Width = flxDmdLeaseList.ColWidth(2)

   picDmdLeaseList.Top = txtSupplierName.Top + txtSupplierName.Height + 5
   picDmdLeaseList.Left = txtSupplierName.Left + 5
   picDmdLeaseList.Visible = True
   picDmdLeaseList.ZOrder 0
   flxDmdLeaseList.SetFocus
End Sub

Private Sub ConfigureDmdFlxLeaseList()
   Dim szHeader As String

   flxDmdLeaseList.Clear
   flxDmdLeaseList.Cols = 3
   flxDmdLeaseList.RowHeight(0) = 0
   szHeader$ = "|<Lessee ID|<Lessee Name"
   flxDmdLeaseList.FormatString = szHeader$
   flxDmdLeaseList.ColWidth(0) = Label2(0).Left - flxDmdLeaseList.Left   '240        Solid column
   flxDmdLeaseList.ColWidth(1) = Label2(1).Left - Label2(0).Left - 20  '1400       'Tenant ID
   flxDmdLeaseList.ColWidth(2) = flxDmdLeaseList.Left + flxDmdLeaseList.Width - Label2(1).Left - 300 'Unit Name
   flxDmdLeaseList.Rows = 2
End Sub

Private Sub cmdOpenAttch_Click()
   Dim i As Integer
   
   MousePointer = vbHourglass

   For i = 1 To flxEmailRemitt.Rows - 1
      If flxEmailRemitt.TextMatrix(i, 0) = "X" And flxEmailRemitt.TextMatrix(i, 1) <> "" Then 'And flxEmailRemitt.TextMatrix(i, 1) <> "" has been added by anol 20161103
         If OpenFile(FileName_FilePath(flxEmailRemitt.TextMatrix(i, 5)), _
                     FilePath_FilePath(flxEmailRemitt.TextMatrix(i, 5))) < 32 Then _
            ShowMsgInTaskBar "File has been moved from original location.", "Y", "N"
      End If
   Next i

   MousePointer = vbDefault
End Sub

Private Sub cmdSort_Click()
   Dim i As Integer

   chkSelectAll.Value = 0

   If optEmail(0).Value Then
      optEmail_Click 0
   Else
      optEmail_Click 1
   End If
End Sub

'Private Sub Command1_Click()
'    inspectgrid
'End Sub

Private Sub flxDmdLeaseList_Click()
   txtSupplierName.text = flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) & " / " & flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 2)
   
   chkSelectAll.Value = 0
   
   picDmdLeaseList.Visible = False
   
   If optEmail(0).Value Then
      optEmail_Click 0
   Else
      optEmail_Click 1
   End If
End Sub

Private Sub flxDmdLeaseList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDmdLeaseList_Click
    End If
End Sub

Private Sub flxEmailRemitt_Click()
   ToggleGridRowSelection flxEmailRemitt
End Sub

Private Sub flxEmailRemitt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxEmailRemitt.ToolTipText = flxEmailRemitt.TextMatrix(flxEmailRemitt.MouseRow, 8)
End Sub

Private Sub Form_Load()
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 15855
   Me.Height = 7215
   Me.BackColor = MODULEBACKCOLOR
   fraClient.BackColor = Me.BackColor
   fraClientMain.BackColor = Me.BackColor
   picDmdLeaseList.BackColor = Me.BackColor
   Frame2.BackColor = Me.BackColor
   optEmail(0).BackColor = Me.BackColor
   optEmail(1).BackColor = Me.BackColor
   chkSelectAll.BackColor = Me.BackColor

   txtToDate.text = Format(Now, "dd/mm/yyyy")

   ConfigFlxEmailRemitt

   LoadLessee

   LoadFlxEmailRemitt
   Call WheelHook(Me.hWnd)
End Sub

Private Sub LoadLessee()
   Dim adoConn As New ADODB.Connection
   Dim szSQL As String

   adoConn.Open getConnectionString

   ConfigureDmdFlxLeaseList
      
   szSQL = "SELECT SageAccountNumber, Name " & _
           "FROM   Tenants " & _
           "ORDER BY SageAccountNumber;"

   LoadFlxDmdLeaseList adoConn, szSQL

   LoadClients adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub CheckClientName()
   Dim i          As Integer

   For i = 1 To flxEmailRemitt.Rows - 1
      If flxEmailRemitt.TextMatrix(i, 0) = "X" And _
            flxEmailRemitt.RowHeight(i) = 240 And _
            flxEmailRemitt.TextMatrix(i, 8) = "" Then
         fraClientMain.Width = Me.Width
         fraClientMain.Height = Me.Height
         fraClientMain.Top = 0
         fraClientMain.Left = 0
         fraClient.Top = fraClientMain.Height / 2 - fraClient.Height
         fraClient.Left = fraClientMain.Width / 2 - fraClient.Width / 2
         lblSupplier.Caption = flxEmailRemitt.TextMatrix(i, 3)
         fraClientMain.Visible = True
         iRowNum = i
         bEmail = False

         Exit Sub
      End If
   Next i
   bEmail = True
End Sub

Private Sub cmdResend_Click()
   bEmail = True
   CheckClientName
   If Not bEmail Then Exit Sub

   Dim adoConn       As New ADODB.Connection
   Dim szReceipent   As String
   Dim i             As Integer
   Dim szSubject     As String
   Dim szBody        As String
   Dim szClient      As String
   Dim colAtt        As New Collection
   Dim iKount        As Integer
   Dim k As Integer

   MousePointer = vbHourglass

   adoConn.Open getConnectionString

   iKount = 0
   Call inspectgrid
   For i = flxEmailRemitt.Rows - 1 To 1 Step -1
      If flxEmailRemitt.TextMatrix(i, 0) = "X" And flxEmailRemitt.RowHeight(i) = 240 Then
         szReceipent = ReceipentEmail(flxEmailRemitt.TextMatrix(i, 7), adoConn)

         If szReceipent <> "Not_Found" And flxEmailRemitt.TextMatrix(i, 8) <> "Cancelled" Then
            Demand_EmailText adoConn, szSubject, szBody, flxEmailRemitt.TextMatrix(i, 8), flxEmailRemitt.TextMatrix(i, 3)
            
            Set colAtt = Nothing
            
            'If colAtt.count > 0 Then colAtt.Remove (1)
            k = i
            Do
                If flxEmailRemitt.TextMatrix(k, 0) = "X" And flxEmailRemitt.RowHeight(k) = 240 Then
                     colAtt.Add flxEmailRemitt.TextMatrix(k, 5)
                End If
                k = k - 1
            Loop While (flxEmailRemitt.TextMatrix(k, 10) <> "sendnow" Or flxEmailRemitt.TextMatrix(k, 0) = "") And k > 0
'            colAtt.Add flxEmailRemitt.TextMatrix(i, 5)
            If flxEmailRemitt.TextMatrix(i, 10) = "sendnow" Then
               ' szReceipent = "anolcse@gmail.com"
                If SendEmail(szFromEmail, szReceipent, _
                               szSubject, _
                               szBody, , , _
                               colAtt, flxEmailRemitt.TextMatrix(i, 7), "SI", flxEmailRemitt.TextMatrix(i, 8)) Then

                     flxEmailRemitt.row = i
'               ToggleGridRowSelection flxEmailRemitt
                     iKount = iKount + 1
                End If
            End If
         End If
      End If
   Next i
   For i = flxEmailRemitt.Rows - 1 To 1 Step -1
        flxEmailRemitt.TextMatrix(i, 0) = ""
   Next i
   adoConn.Close
   Set adoConn = Nothing

   MousePointer = vbDefault

   MsgBox iKount & " Email Sent.", vbInformation + vbOKOnly, "Resend Email"
End Sub
Private Sub inspectgrid()
    Dim i As Integer
    Dim szLessee As String
    Dim strDate As String
    Dim strTime As String
    For i = flxEmailRemitt.Rows - 1 To 1 Step -1
      If flxEmailRemitt.TextMatrix(i, 0) = "X" And flxEmailRemitt.RowHeight(i) = 240 Then
            If szLessee = "" And strDate = "" And strTime = "" Then
                flxEmailRemitt.TextMatrix(i, 10) = "sendnow"
                GoTo XX
            End If
            If strDate = flxEmailRemitt.TextMatrix(i, 1) And strTime = flxEmailRemitt.TextMatrix(i, 2) And szLessee = flxEmailRemitt.TextMatrix(i, 3) Then
                 flxEmailRemitt.TextMatrix(i, 10) = ""
                 GoTo XX
            End If
            If strDate = flxEmailRemitt.TextMatrix(i, 1) And strTime = flxEmailRemitt.TextMatrix(i, 2) And szLessee = flxEmailRemitt.TextMatrix(i, 3) Then
                 flxEmailRemitt.TextMatrix(i, 10) = ""
            Else
                 flxEmailRemitt.TextMatrix(i, 10) = "sendnow"
            End If
XX:
            strDate = flxEmailRemitt.TextMatrix(i, 1)
            strTime = flxEmailRemitt.TextMatrix(i, 2)
            szLessee = flxEmailRemitt.TextMatrix(i, 3)
      Else
            flxEmailRemitt.TextMatrix(i, 10) = ""
      End If
    Next i

End Sub
Private Sub cmdOK_Click()
   If cboClientID.text = "" Then
      MsgBox "Please select a client to continue.", vbInformation + vbOKOnly, "Resending Email"
      cboClientID.SetFocus
      Exit Sub
   End If
   
   flxEmailRemitt.TextMatrix(iRowNum, 8) = cboClientID.Column(1)

   fraClientMain.Visible = False

   CheckClientName
End Sub

Public Function ReceipentEmail(szSupp As String, ByVal adoConn As ADODB.Connection) As String
   Dim szSQL As String
   Dim adoRst As New ADODB.Recordset

   ReceipentEmail = "Not_Found"

   szSQL = "SELECT Email1, Email2, InvoiceTo " & _
           "FROM Tenants " & _
           "WHERE SageAccountNumber = '" & szSupp & "' " & _
           "ORDER BY Name;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      If adoRst.Fields.Item("InvoiceTo").Value = "B" Then        'Alternet Address
         ReceipentEmail = IIf(IsNull(adoRst.Fields.Item("Email2").Value), "Not_Found", adoRst.Fields.Item("Email2").Value)
      Else                                                     'Lessee Address
         ReceipentEmail = IIf(IsNull(adoRst.Fields.Item("Email1").Value), "Not_Found", adoRst.Fields.Item("Email1").Value)
      End If
   End If

   adoRst.Close
   Set adoRst = Nothing
End Function

Public Function LoadFlxDmdLeaseList(adoConn As ADODB.Connection, ByVal sSQLQuery_ As String)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open sSQLQuery_, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox vbTab & "Either there are no lessee records entered in the system. " & vbCrLf & _
             "Please enter a lessee in the lessee module.", vbInformation + vbOKOnly, "Email Archive"

      GoTo NoRes
   End If

   Dim iRow As Integer
   iRow = 1

   While Not adoRst.EOF
      flxDmdLeaseList.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
      flxDmdLeaseList.TextMatrix(iRow, 2) = adoRst!Name

      iRow = iRow + 1
      adoRst.MoveNext

      If Not adoRst.EOF Then flxDmdLeaseList.AddItem ""
   Wend

NoRes:
   adoRst.Close
   Set adoRst = Nothing
End Function

Private Sub LoadFlxEmailRemitt()
   LoadSentEmail_SI flxEmailRemitt, flxDmdLeaseList

   optEmail_Click 0
End Sub

Private Sub ConfigFlxEmailRemitt()
   Dim szHeader As String

   flxEmailRemitt.Clear
   
   flxEmailRemitt.Cols = 12
   szHeader$ = "X|<Date|<Time|<Lessee|<Email|AttPath|Result|SageAccountNumber|Client"
   
   flxEmailRemitt.Rows = 2
   flxEmailRemitt.RowHeight(0) = 0

   flxEmailRemitt.FormatString = szHeader$

   flxEmailRemitt.ColAlignment(0) = vbCenter
   flxEmailRemitt.ColWidth(0) = Label1(1).Left - flxEmailRemitt.Left    'X
   flxEmailRemitt.ColWidth(1) = Label1(2).Left - Label1(1).Left         'Date
   flxEmailRemitt.ColWidth(2) = Label1(3).Left - Label1(2).Left         'Time
   flxEmailRemitt.ColWidth(3) = Label1(4).Left - Label1(3).Left         'Lessee
   flxEmailRemitt.ColWidth(4) = 2800 'flxEmailRemitt.Width + flxEmailRemitt.Left - Label1(4).Left - 300 'Email
   flxEmailRemitt.ColWidth(5) = 0                                     'Attch
   flxEmailRemitt.ColWidth(6) = 0                                       'Result
   flxEmailRemitt.ColWidth(7) = 0                                       'SageAccountNumber
   flxEmailRemitt.ColWidth(8) = 0
   flxEmailRemitt.ColWidth(9) = 3700        'Client
   flxEmailRemitt.ColWidth(10) = 0        'Inspec
   flxEmailRemitt.ColWidth(11) = 1200       'DemandID
   flxEmailRemitt.ColAlignment(11) = vbLeftJustify
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub optEmail_Click(Index As Integer)
   Dim i As Integer

   For i = 1 To flxEmailRemitt.Rows - 1
      flxEmailRemitt.RowHeight(i) = 240
      If flxEmailRemitt.TextMatrix(i, 1) <> "" Then
         If optEmail(0).Value And flxEmailRemitt.TextMatrix(i, 6) <> "True" Then    'Sent
            flxEmailRemitt.RowHeight(i) = 0
         End If
         If optEmail(1).Value And flxEmailRemitt.TextMatrix(i, 6) = "True" Then     'Failed
            flxEmailRemitt.RowHeight(i) = 0
         End If
         If CDate(flxEmailRemitt.TextMatrix(i, 1)) < CDate(txtFromDate.text) Or _
               CDate(flxEmailRemitt.TextMatrix(i, 1)) > CDate(txtToDate.text) Then
            flxEmailRemitt.RowHeight(i) = 0
         End If
         If txtSupplierName.text <> "All Lessees" Then
            If flxEmailRemitt.TextMatrix(i, 7) <> flxDmdLeaseList.TextMatrix(flxDmdLeaseList.row, 1) Then
               flxEmailRemitt.RowHeight(i) = 0
            End If
         End If
      End If
   Next i
End Sub

Private Sub txtDmdTenantSearchID_Change()
   Dim i As Integer

   If Len(txtDmdTenantSearchID.text) > 0 Then txtDmdTenantSearchName.text = ""

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 1), Len(txtDmdTenantSearchID.text))) <> UCase(txtDmdTenantSearchID.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtDmdTenantSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDmdTenantSearchName.SetFocus
    End If
End Sub

Private Sub txtDmdTenantSearchName_Change()
   Dim i As Integer

   If Len(txtDmdTenantSearchName.text) > 0 Then txtDmdTenantSearchID.text = ""

   For i = 1 To flxDmdLeaseList.Rows - 1
      flxDmdLeaseList.RowHeight(i) = 240
      If UCase(Left(flxDmdLeaseList.TextMatrix(i, 2), Len(txtDmdTenantSearchName.text))) <> UCase(txtDmdTenantSearchName.text) Then
         flxDmdLeaseList.RowHeight(i) = 0
      End If
   Next i
End Sub

Private Sub txtDmdTenantSearchName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDmdLeaseList.SetFocus
    End If
End Sub

Private Sub txtSupplierName_Change()
   If txtSupplierName.text = "" Then
      txtSupplierName.text = "All Suppliers"

      If optEmail(0).Value Then
         optEmail_Click 0
      Else
         optEmail_Click 1
      End If
   End If
End Sub

Private Sub txtSupplierName_KeyPress(KeyAscii As Integer)
   If iKeyUp <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtSupplierName_KeyUp(KeyCode As Integer, Shift As Integer)
   iKeyUp = KeyCode
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
   SelTxtInCtrl txtFromDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtToDate_GotFocus()
   SelTxtInCtrl txtToDate
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate
End Sub

Private Sub LoadClients(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM   CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboClientID.Column() = Data()

   adoRst.Close
   Set adoRst = Nothing
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
            bHandled = False
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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
