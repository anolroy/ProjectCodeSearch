VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmServiceCharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Charge Budget"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   Icon            =   "frmServiceCharge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   13350
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   10170
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   28
      Top             =   1890
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
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   10
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   7091
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
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   29
         Top             =   135
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Client Name"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   8
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   9
         Top             =   375
         Width           =   4545
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8017;450"
         SpecialEffect   =   0
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
         Left            =   45
         Top             =   75
         Width           =   5850
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   345
      Left            =   9840
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSCBdCancel 
      Caption         =   "&Cancel"
      Height          =   385
      Left            =   13275
      TabIndex        =   12
      Top             =   7110
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6345
      Index           =   1
      Left            =   75
      TabIndex        =   13
      Top             =   915
      Width           =   9360
      Begin VB.CommandButton cmdRunFlag 
         Caption         =   "Manage Year End"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   225
         TabIndex        =   39
         Top             =   5130
         Width           =   1755
      End
      Begin VB.PictureBox fmeLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   4140
         ScaleHeight     =   315
         ScaleWidth      =   2655
         TabIndex        =   36
         Top             =   2385
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label lblLoading 
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait while importing..."
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
            Left            =   90
            TabIndex        =   37
            Top             =   45
            Width           =   2745
         End
      End
      Begin VB.CommandButton cmdSCBdClose 
         Caption         =   "Cl&ose"
         Height          =   385
         Left            =   7995
         TabIndex        =   7
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSCBdDelete 
         Caption         =   "&Delete"
         Height          =   385
         Left            =   2385
         TabIndex        =   5
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdSCBdNew 
         Caption         =   "&New"
         Height          =   385
         Left            =   90
         TabIndex        =   4
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import"
         Height          =   385
         Left            =   4695
         TabIndex        =   6
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtBudgetId 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "txtBudgetId"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMatrixRow 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Text            =   "txtMatrixRow"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   16
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSCTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5055
         Width           =   1935
      End
      Begin VB.TextBox txtSCBudgetTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6300
         TabIndex        =   14
         Top             =   5055
         Width           =   2715
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSCBudgetDetails 
         Height          =   4665
         Left            =   90
         TabIndex        =   3
         Top             =   360
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   8229
         _Version        =   393216
         ForeColor       =   0
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   12632256
         BackColorSel    =   15329508
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   8421504
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
         _Band(0).Cols   =   6
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Code"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   38
         Top             =   90
         Width           =   780
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9360
         Y1              =   5535
         Y2              =   5535
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   3
         Left            =   6285
         TabIndex        =   21
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price/SqFoot"
         Height          =   195
         Index           =   4
         Left            =   7605
         TabIndex        =   20
         Top             =   120
         Width           =   945
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area"
         Height          =   195
         Index           =   2
         Left            =   4605
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund Name"
         Height          =   195
         Index           =   0
         Left            =   2235
         TabIndex        =   18
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   17
         Top             =   5055
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdClientList 
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
      Height          =   300
      Left            =   4680
      TabIndex        =   0
      Top             =   135
      Width           =   300
   End
   Begin VB.CommandButton cmdProperty 
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
      Height          =   300
      Left            =   9090
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton cmdBudgetYears 
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
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   495
      Width           =   300
   End
   Begin MSForms.TextBox txtBudgetYears 
      Height          =   285
      Left            =   1080
      TabIndex        =   35
      Top             =   495
      Width           =   3600
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6350;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   5805
      TabIndex        =   34
      Top             =   90
      Width           =   3285
      VariousPropertyBits=   746604571
      Size            =   "5794;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1080
      TabIndex        =   33
      Top             =   135
      Width           =   3600
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "6350;503"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   3
      Left            =   5160
      TabIndex        =   25
      Top             =   135
      Width           =   630
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Year:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   930
   End
End
Attribute VB_Name = "frmServiceCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public bEDIT               As Boolean
Public bNew                As Boolean
Private bFormLoaded        As Boolean
'Dim adoConn                As New ADODB.Connection
Dim szSQL                  As String

Dim flgChange              As Integer
Dim detailsMatrix(69, 69)  As clsSCDtl
Dim matrixLength           As Integer
Dim bLoadingProperty       As Boolean
Dim sTextBox As String
'Dim bufferMatrix()         As clsSCDtl
'
'Public Function fillBufferMatrix(ByVal bDetails As clsSCDtl, ByVal num As Integer)
'   ReDim Preserve bufferMatrix(num) As clsSCDtl
'
'   Set bufferMatrix(num - 1) = New clsSCDtl
'
'   bufferMatrix(num - 1).setBudgetAmount bDetails.getBudgetAmount
'   bufferMatrix(num - 1).setBudgetDetailId bDetails.getBudgetDetailID
'   bufferMatrix(num - 1).setBudgetId bDetails.getBudgetId
'   bufferMatrix(num - 1).setFlgDel bDetails.getFlgDel
'   bufferMatrix(num - 1).setNCode bDetails.getNCode
'   bufferMatrix(num - 1).setNName bDetails.getNName
'End Function
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
'   lblClientID.Width = 1400
'   lblClientID.Left = 50
'   lblClientName.Width = 2600
'   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
'   flxClient.Width = 5175
   
   
   adoconn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
'           flxClient.TextMatrix(rRow, 0) = ""
'           flxClient.TextMatrix(rRow, 1) = ""
'           flxClient.TextMatrix(rRow, 2) = ""
'           flxClient.RowHeight(rRow) = 240
'           flxClient.AddItem ""
'           rRow = 2
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

Private Sub cmdBudgetYears_Click()
    sTextBox = "3"
    picClient.Left = 915
    picClient.Top = 170
    LoadGridFY
    picClient.Visible = True
    cmdProperty.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1(1).Enabled = True
    cmdClientList.Enabled = True
    cmdProperty.Enabled = True
'    If sTextBox = "1" Then
'        cmdDmdClientList.SetFocus
'    ElseIf sTextBox = "2" Then
'         cmdDmdPropertyList.SetFocus
'    ElseIf sTextBox = "3" Then
'        cmdClientList.SetFocus
'    ElseIf sTextBox = "4" Then
'        cmdPropertyList.SetFocus
'    End If
End Sub

Private Sub cmdproperty_Click()
    sTextBox = "2"
    picClient.Left = 3215
    picClient.Top = 70
    LoadPropertyList
    picClient.Visible = True
    cmdClientList.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdRunFlag_Click()
    LoadForm frmServiceChargeRunFlag
End Sub

Private Sub flxClient_Click()
        Frame1(1).Enabled = True
        cmdClientList.Enabled = True
        cmdProperty.Enabled = True
        
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        If sTextBox = "1" Then
               
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.text = ""
                txtPropertyName.Tag = ""
               
                Dim adoRst As New ADODB.Recordset
                Dim szSQL As String
                
                szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
                adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields(1).Value
                        txtPropertyName.Tag = adoRst.Fields(0).Value
                        Call cboProperty_Change
                        'added by anol 20161013
                        LoadFlxSCBudgetDetails adoconn
                        
                        LoadFY adoconn
                        cboBudgetYears_Change
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
                End If
                Dim i       As Integer
                Dim Rst     As New ADODB.Recordset
                szSQL = "SELECT g.*, f.FundName,f.FundCode,P.clientID " & _
                        "FROM GlobalSC g, Fund f,Property P " & _
                        "WHERE CInt(g.Fund)=f.FundId AND P.propertyID=G.propertyID and F.CategoryCode=5 and P.ClientID='" & txtClientList.Tag & "';"
                        
                Rst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                If Not Rst.EOF Then
                        cmdRunFlag.Enabled = True
                Else
                         cmdRunFlag.Enabled = False
                End If
                Rst.Close
   
        
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
               
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                
                'added by anol 20161013
                LoadFlxSCBudgetDetails adoconn
                
                Call cboProperty_Change
                LoadFY adoconn
                cboBudgetYears_Change
                cmdBudgetYears.SetFocus
        End If
        If sTextBox = "3" Then
                txtBudgetYears.text = flxClient.TextMatrix(flxClient.row, 1)
                txtBudgetYears.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                Call cboBudgetYears_Change
                cmdBudgetYears.SetFocus
        End If
        adoconn.Close
       
        picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub



Private Sub txtSearchClientID_Change()
    'Updated by anol 22 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
              flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
              flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyDown Then
           flxClient.SetFocus
    End If
    If KeyCode = 13 Then
           txtSearchClientName.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
         picClient.Visible = False
          Frame1(1).Enabled = True
          'Frame2.Enabled = True
          If sTextBox = "1" Then
                 cmdClientList.SetFocus
           ElseIf sTextBox = "2" Then
'                cmdproperty.SetFocus
           End If
    End If
End Sub

Private Sub txtSearchClientName_Change()
   'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
        flxClient.RowHeight(i) = 240
        If InStr(1, UCase(flxClient.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
        End If
        If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
        End If
   Next i
End Sub

Private Sub txtSearchClientName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = 13 Then
         flxClient.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            flxClient.SetFocus
        End If
    End If
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1500
   flxClient.ColWidth(2) = 4500
   
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'   picClient.Height = 4095
'   flxClient.Height = 3345
  ' flxClient.Width = 5175
   
   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
               flxClient.RowHeight(rRow) = 240
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
            Wend
      
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

End Sub
Private Sub LoadCmbClient(adoconn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
        txtClientList.text = adoRst.Fields("CLIENTNAME").Value
        txtClientList.Tag = adoRst.Fields("CLIENTID").Value
                adoRst.Close
'                szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
            szSQL = "SELECT FYrID, FinancialYear, FY_Description, P.PropertyID, P.PropertyName " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.ClientID = '" & txtClientList.Tag & "' ORDER BY P.PropertyID;"
                adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields("PropertyName").Value
                        txtPropertyName.Tag = adoRst.Fields("PropertyID").Value
                        txtBudgetYears.text = adoRst.Fields("FinancialYear").Value
                        txtBudgetYears.Tag = adoRst.Fields("FYrID").Value
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
                        txtBudgetYears.text = ""
                        txtBudgetYears.Tag = ""
                End If
   Else
            adoRst.Close
   End If
               
End Sub



Private Sub cmdClientList_Click()
    sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 70
    
    LoadflxClient
    picClient.Visible = True
    cmdProperty.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub
Private Sub cboBudgetYears_Change()
   Dim iRow As Integer

   For iRow = 1 To flxSCBudgetDetails.Rows - 1
   'cboProperty.Column(4)
      If txtBudgetYears.Tag = flxSCBudgetDetails.TextMatrix(iRow, 9) And txtPropertyName.Tag = flxSCBudgetDetails.TextMatrix(iRow, 1) Then
         flxSCBudgetDetails.RowHeight(iRow) = 240
      Else
         flxSCBudgetDetails.RowHeight(iRow) = 0
      End If
   Next iRow

   SCSumTotal
End Sub

'Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'   Dim errsmg As String
'   On Error GoTo ErrorHandler
'
''*************************************** CLIENT COMBO ******************************************
'   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
'               "LandLordSageCustAC, LandLordSageSuppAC " & _
'           "FROM CLIENT " & _
'           "ORDER BY CLIENTID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim Data() As String
'   Dim i As Integer, j As Integer
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'
'   cboClient.Column() = Data()
'   cboClient.ListIndex = 0
'   adoRst.Close
'   errsmg = "Client load completed."
''*************************************** PROPERTY ******************************************
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode, CBY " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY PropertyID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'      For j = 0 To TotalCol
'         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'      Next j
'      adoRst.MoveNext
'      If adoRst.EOF Then Exit For
'   Next i
'   cboProperty.Column() = Data()
'   cboProperty.ListIndex = 0
'    errsmg = "Client and property load completed."
'NoRes:
'   'Resolved by BOSL
'   'Modified by Anol 09 Sep 2014
'   'Issue number 471 note 6
'       If adoRst.State = 1 Then
'            adoRst.Close
'            Set adoRst = Nothing
'       End If
'
'   Exit Sub
'   Exit Sub
'
'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number & "-Error on Preparelist" & errsmg
'
'   'Resolved by BOSL
'   'Modified by Anol 09 Sep 2014
'   'Issue number 471 note 6
'       If adoRst.State = 1 Then
'            adoRst.Close
'            Set adoRst = Nothing
'       End If
'End Sub

Private Sub cboClient_Change(adoconn As ADODB.Connection)
   If Not bFormLoaded Then Exit Sub
   If txtClientList.text = "" Then Exit Sub
   bLoadingProperty = True

'   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   On Error GoTo ErrorHandler

'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode, CBY " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY PropertyID;"
'
'   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then GoTo NoRes
'
'   TotalRow = adoRst.RecordCount - 1
'   TotalCol = adoRst.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
'       Next j
'       adoRst.MoveNext
'       If adoRst.EOF Then Exit For
'   Next i
'   cboProperty.Column() = Data()
'   cboProperty.ListIndex = 0
'   bLoadingProperty = False

   LoadFY adoconn

NoRes:
   adoRst.Close
   adoconn.Close
   Set adoRst = Nothing
'   Set adoConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   'Modified by anol 14 Sep 2014
   'issue 471 note 6  production
    If adoconn.State = 1 Then
        adoconn.Close
        Set adoRst = Nothing
    End If
   ' End of modification
'   Set adoConn = Nothing
End Sub

Public Function initialiseMatrix() As Integer
   Dim i As Integer, j As Integer

   For i = 0 To 59
      For j = 0 To 59
         Set detailsMatrix(i, j) = New clsSCDtl
      Next j
   Next i
End Function

Public Function getDetailsFromMatrix(ByVal row As Integer, ByVal col As Integer) As clsSCDtl
   Set getDetailsFromMatrix = detailsMatrix(row, col)
End Function

Private Sub cboProperty_Change()
    'Created by Anol 29 Oct 2014
    'issue 471 note 724
    If bLoadingProperty = False Then Exit Sub
    Dim strCheck As String
    Dim rsSQL As New ADODB.Recordset
    Dim adoConn1 As New ADODB.Connection
    If adoConn1.State = 0 Then
        adoConn1.Open getConnectionString
    End If
    szSQL = "SELECT P.CBY " & _
        "FROM Property P " & _
        "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
    rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
    If Not rsSQL.EOF Then
        strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
    End If
    If strCheck = "" Then
        ShowMsgInTaskBar "A service charge budget year has not been set for this property." & vbCrLf & "Please set a service charge budget year in the global data screen.", "Y", "N"
    End If
    If adoConn1.State = 1 Then
        adoConn1.Close
    End If
End Sub

Private Sub cboProperty_Click()
   'issue 471 note 724
   'added by anol 05 Nov 2014
   If bFormLoaded Then
      Dim strCheck As String
      Dim rsSQL As New ADODB.Recordset
      Dim adoConn1 As New ADODB.Connection
      If adoConn1.State = 0 Then
         adoConn1.Open getConnectionString
      End If
      szSQL = "SELECT P.CBY " & _
         "FROM Property P " & _
         "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
      rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
      If Not rsSQL.EOF Then
      strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
      End If
      If strCheck = "" Then
         ShowMsgInTaskBar "A service charge budget year has not been set for this property." & vbCrLf & "Please set a service charge budget year in the global data screen.", "Y", "N"
      End If
      If adoConn1.State = 1 Then
         adoConn1.Close
      End If
   End If
   'End of modification
   If bLoadingProperty Then Exit Sub
   'txtBudgetYears.Tag = cboProperty.Column(4)
   'Grid was not refresing
   'issue 471
   'Modified By anol 11 Sep 2014
    Call cboBudgetYears_Change
    'End of modification
      'Created by Anol 29 Oct 2014
      'issue 471 note 724
      'If bLoadingProperty = False Then Exit Sub
      
End Sub

Private Function CreateReportSCImport(adoconn As ADODB.Connection, oWS As Worksheet, szNC As String, szFund As String) As Boolean
    ' Get the list of nominal code
        Dim iRow As Integer
        
        Dim szNC_NotExist    As String
        Dim szFund_NotExist  As String
        Dim adoRst As New Recordset
        Dim rsReportSCImport As New ADODB.Recordset
        
        adoRst.Open "SELECT Code & ' # '  & Name FROM NominalLedger WHERE ClientID = '" & txtClientList.Tag & "';", adoconn, adOpenStatic, adLockReadOnly
        szNC = SQL2String(adoRst, 0)
        adoRst.Close
        'Modified by anol 18 04 2016
        adoRst.Open "SELECT FundCode & ' # '  & FundID FROM Fund where (CategoryCode=2 OR CategoryCode=5);", adoconn, adOpenStatic, adLockReadOnly
        szFund = SQL2String(adoRst, 0)
        adoRst.Close
   
   
         Dim reportingDate As String
         reportingDate = Format(DateValue(Now), "dd mmmm yyyy")
         Dim sessionID As String
         sessionID = GetTimeStamp
        ' adoConn.Execute "DELETE FROM ReportSCImport WHERE ReportingDate < #" & reportingDate & "# ;"
         adoconn.Execute "DELETE FROM ReportSCImport;"
         rsReportSCImport.Open "Select * from ReportSCImport where 1=2", adoconn, adOpenStatic, adLockOptimistic
         
         iRow = 2
         While oWS.Range("A" & CStr(iRow)).Value <> vbEmpty
             
                
            If oWS.Range("B" & iRow).Value = vbEmpty Then
                MsgBox " A Fund Code is missing in row no: " & iRow & ". Please check your import file.", vbInformation, "Import failed"
                
                CreateReportSCImport = True
                Exit Function
            End If
            If oWS.Range("C" & iRow).Value = vbEmpty Then
                 MsgBox "A Nominal Code is missing in row no: " & iRow & ". Please check your import file.", vbInformation, "Import failed"
                 CreateReportSCImport = True
                 Exit Function
            End If
            With rsReportSCImport
                .AddNew
                !reportingDate = reportingDate
                !sessionID = sessionID
                !FUNDORNOMINAL = "N"
                !NominalName = oWS.Range("A" & iRow).Value
                !FundCode = oWS.Range("B" & iRow).Value
                !nominalCode = oWS.Range("C" & iRow).Value
                !budgetAmount = Val(IIf(IsNull(oWS.Range("D" & iRow).Value), 0, oWS.Range("D" & iRow).Value))
                .Update
            End With
             
             
            iRow = iRow + 1
         Wend
         If rsReportSCImport.State = 1 Then
            rsReportSCImport.Close
         End If
'         rsReportSCImport.Open "Select count(NominalCode) from ReportSCImport  group by NominalCode having Count(NominalCode)>1", adoConn, adOpenStatic, adLockOptimistic
         rsReportSCImport.Open "Select count(NominalCode+fundcode) from ReportSCImport  group by NominalCode+fundcode having Count(NominalCode+fundcode)>1", adoconn, adOpenStatic, adLockOptimistic
         If Not rsReportSCImport.EOF Then
              MsgBox "There are Duplicate Nominal Codes in your import file. Please check your import file and try again.", vbInformation, "Import failed"
             CreateReportSCImport = True
             Exit Function
         End If
         If rsReportSCImport.State = 1 Then
            rsReportSCImport.Close
         End If
         adoconn.Execute "DELETE FROM ReportSCImport;"
         rsReportSCImport.Open "Select * from ReportSCImport where 1=2", adoconn, adOpenStatic, adLockOptimistic
         iRow = 2
         
         
   While oWS.Range("A" & CStr(iRow)).Value <> vbEmpty 'loop with excel row
      
      If InStr(UCase(szFund), UCase(oWS.Range("B" & iRow).Value)) <= 0 Then
            szFund_NotExist = szFund_NotExist & oWS.Range("B" & iRow).Value & ", "
            'insert into report table when fund not valid
            With rsReportSCImport
                .AddNew
                !reportingDate = reportingDate
                !sessionID = sessionID
                !FUNDORNOMINAL = "F"
                !NominalName = oWS.Range("A" & iRow).Value
                !FundCode = oWS.Range("B" & iRow).Value
                !nominalCode = oWS.Range("C" & iRow).Value
                !budgetAmount = Val(IIf(IsNull(oWS.Range("D" & iRow).Value), 0, oWS.Range("D" & iRow).Value))
                .Update
            End With
      End If
      
      If InStr(szNC, oWS.Range("C" & iRow).Value) <= 0 Then
            szNC_NotExist = szNC_NotExist & oWS.Range("C" & iRow).Value & ", "
             'insert into report table when NC not valid
            With rsReportSCImport
                .AddNew
                !reportingDate = reportingDate
                !sessionID = sessionID
                !FUNDORNOMINAL = "N"
                !NominalName = oWS.Range("A" & iRow).Value
                !FundCode = oWS.Range("B" & iRow).Value
                !nominalCode = oWS.Range("C" & iRow).Value
                !budgetAmount = Val(IIf(IsNull(oWS.Range("D" & iRow).Value), 0, oWS.Range("D" & iRow).Value))
                .Update
            End With
       End If
       iRow = iRow + 1
'       If iRow > 100 Then
'
'                Debug.Print ""
'       End If
    Wend
    If szNC_NotExist = "" And szFund_NotExist = "" Then
          'all are matched
          'so no report needs to come for missmatched Fund and nominal ,so exit this function
          Exit Function
    Else
          If szNC_NotExist <> "" Then
             ShowMsgInTaskBar "Import Failed due to Invalid Nominal Code", "Y", "N"
             szNC_NotExist = Mid(Trim(szNC_NotExist), 1, Len(Trim(szNC_NotExist)) - 1)
          Else
             szNC_NotExist = "NULL"
          End If
          If szFund_NotExist <> "" Then
             ShowMsgInTaskBar "Import Failed due to Invalid Fund Code", "Y", "N"
             szFund_NotExist = Mid(Trim(szFund_NotExist), 1, Len(Trim(szFund_NotExist)) - 1)
          Else
             szFund_NotExist = "NULL"
          End If
    
          Dim reportApp As New CRAXDRT.Application
          Dim Report As CRAXDRT.Report
          Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SC_Imp_Report.rpt")
          Report.EnableParameterPrompting = False
          Report.DiscardSavedData
          Report.ParameterFields(1).AddCurrentValue szNC_NotExist
          Report.ParameterFields(2).AddCurrentValue szFund_NotExist
          Load frmReport
          frmReport.LoadReportViewer Report
          CreateReportSCImport = True
   End If
End Function
Private Sub CreateTableSCImport(adoconn As ADODB.Connection)
        'creating temp table
        On Error GoTo CreateReportSCImport
            Dim adoRst As New ADODB.Recordset
           adoRst.Open "SELECT * FROM ReportSCImport;", adoconn, adOpenStatic, adLockReadOnly
           adoRst.Close
           Exit Sub
           
        
CreateReportSCImport:
           adoconn.Execute _
              "CREATE TABLE ReportSCImport " & _
                 "(" & _
                    "ReportingDate DateTime  NOT NULL, " & _
                    "SessionID     TEXT(100) NOT NULL, " & _
                    "FUNDORNOMINAL      TEXT(100), " & _
                    "NOMINALNAME      TEXT(100), " & _
                    "FUNDCODE        TEXT(100), " & _
                    "NominalCode     TEXT(200), " & _
                    "BUDGETAMOUNT   TEXT(15) NOT NULL " & _
                 ");"
        

End Sub
Private Function ExcelHeaderIsvalid(adoconn As ADODB.Connection, oWS As Worksheet) As Boolean
   If UCase(oWS.Range("A1").Value) <> "NOMINAL NAME" Or _
         UCase(oWS.Range("B1").Value) <> "FUND CODE" Or _
         (UCase(oWS.Range("C1").Value) <> "NOMINAL CODE" And UCase(oWS.Range("C1").Value) <> "NOMINAL_CODE") Or _
         UCase(oWS.Range("D1").Value) <> "BUDGET AMOUNT" Then
                    
          ShowMsgInTaskBar "The file format is wrong", "Y", "N"
   Else
       ExcelHeaderIsvalid = True
   End If
End Function
Private Sub cmdImport_Click()
   If txtClientList.text = "" Then
      ShowMsgInTaskBar "Please select a client", "Y", "N"
      cmdClientList.SetFocus
      Exit Sub
   End If
   If txtBudgetYears.Tag = "" Then
      ShowMsgInTaskBar "Please select a budget year", "Y", "N"
      cmdBudgetYears.SetFocus
      Exit Sub
   End If
   If txtPropertyName.text = "" Then
      ShowMsgInTaskBar "Please select a property", "Y", "N"
      cmdProperty.SetFocus
      Exit Sub
   End If
    'Resolved by BOSL
    'issue number 471 Note 724
    'Added by anol 05 Nov 2014
    Dim strCheck As String
    Dim rsSQL As New ADODB.Recordset
    Dim adoConn1 As New ADODB.Connection
    adoConn1.Open getConnectionString
     
    szSQL = "SELECT P.CBY FROM Property P WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
    rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
    If Not rsSQL.EOF Then
          strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
    End If
    If strCheck = "" Then
          MsgBox "A service charge budget year has not been set for this property." & vbCrLf & "Please set a service charge budget year in the global data screen."
          Exit Sub
    End If
    'End of Modification
      
   If MsgBox("Do you wish to import a Service Charge budget?", vbQuestion + vbYesNo, "Import Service Charge Budget") = vbNo Then
        adoConn1.Close
        Exit Sub
   End If
   Dim ofn                    As OPENFILENAME
   Dim lHwnd                  As Long
   Const HKEY_LOCAL_MACHINE   As Long = &H80000002
   Dim szOldFile_PathName     As String
   Dim szNewFile_Path         As String
   Dim szNewFile_Name         As String
   Dim szNewFile_PathName     As String
   Dim fso                    As Object
   Dim szImportFile           As String
  

   On Error GoTo FileError

   ofn.lStructSize = Len(ofn)
   ofn.hwndOwner = lHwnd
   ofn.hInstance = App.hInstance
   ofn.lpstrFilter = "MS Office Excel Workbook 2007-2010 (*.xlsx)" + Chr$(0) + "*.xlsx" + Chr$(0) + _
                     "MS Office Excel Workbook 97-2003 (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + _
                     "CSV Files (*.csv)" + Chr$(0) + "*.csv" + Chr$(0)

   ofn.lpstrFile = Space$(254)
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.lpstrInitialDir = CurDir
   ofn.lpstrTitle = "Select File to Save"
   ofn.Flags = 0

   If GetOpenFileName(ofn) = 0 Then Exit Sub

   szImportFile = ofn.lpstrFile

   If szImportFile = "" Then
      ShowMsgInTaskBar "Please select an input file to import", "Y", "N"
      cmdImport.SetFocus
      Exit Sub
   End If

       Dim oXL As New Excel.Application
       Dim oWB As Workbook
       Dim oWS As Worksheet
    
       Set oWB = oXL.Workbooks.Open(szImportFile)
       Set oWS = oWB.Worksheets(1) 'Specify your worksheet name
       
'        'New checking for dupplicate Nominal Code by anol 20161013
'        Dim adoExcelConn  As New ADODB.Connection
'        Dim rsCheck As New ADODB.Recordset
'        Dim szaFile()     As String
'        Dim iFileNameLen  As Integer
'
'        szaFile = Split(szImportFile, "\")
'        iFileNameLen = Len(szaFile(UBound(szaFile)))
'
'        adoExcelConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                   "Data Source=" & Left(szImportFile, Len(szImportFile) - iFileNameLen) & ";" & _
'                   "Extended Properties=""text;HDR=YES;FMT=Delimited"""
'
'        'Query from the file assuming it as a database table
'         strCheck = ""
'        strCheck = "SELECT * FROM [" & szaFile(UBound(szaFile)) & "];"
'       Does not work on below line

'        rsCheck.Open "SELECT * FROM [" & szaFile(UBound(szaFile)) & "];", _
'                adoExcelConn, adOpenStatic, adLockOptimistic, adCmdText
'        strCheck = ""
'
'        'End of checking dupplicate
             
             
             
             
       Dim iRow             As Integer
       Dim adoRst As New ADODB.Recordset
       Dim szBudgetID       As String
       Dim szNC  As String
       Dim szFund  As String
    'validating the column names are correct in the excel file
    If ExcelHeaderIsvalid(adoConn1, oWS) = False Then Exit Sub
    
    'seperated this procedure on 09 Sep 2016
    'this shall create temporary table for the report
    Call CreateTableSCImport(adoConn1)
    'Written by anol 09 Sep 2016
    ' If NC code and  fund code not in the master table then exit procedure
    If CreateReportSCImport(adoConn1, oWS, szNC, szFund) = True Then
        oWB.Close
        oXL.Quit
        Exit Sub
    End If
    iRow = 2
    lblLoading.Caption = "Please wait while importing..."
    fmeLoading.Visible = True
    cmdImport.Enabled = False
    
    While oWS.Range("B" & CStr(iRow)).Value <> vbEmpty 'loop with excel row
         szBudgetID = getBudgetId(txtPropertyName.Tag, oWS.Range("B" & iRow).Value, txtBudgetYears.Tag, adoConn1)
            If szBudgetID <> "NULL" Then
               szSQL = "SELECT C.* " & _
                       "FROM GlobalSCDtls AS C INNER JOIN GlobalSC AS P ON P.BudgetID = C.BudgetID " & _
                       "WHERE C.BudgetID = '" & szBudgetID & "' AND " & _
                             "P.FinancialYear = '" & txtBudgetYears.Tag & "' AND " & _
                             "C.NC = '" & oWS.Range("C" & iRow).Value & "';"
               adoRst.Open szSQL, adoConn1, adOpenDynamic, adLockOptimistic

               If Not adoRst.EOF Then
                   'Val function was added by anol 22 Apr 2015
                  adoRst.Fields.Item("BudgetAmt").Value = Val(oWS.Range("D" & iRow).Value)
               Else
                  adoRst.AddNew
                  adoRst.Fields.Item("BudgetDtlID").Value = UniqueID()
                  adoRst.Fields.Item("BudgetID").Value = szBudgetID
                  adoRst.Fields.Item("NC").Value = oWS.Range("C" & iRow).Value
                  adoRst.Fields.Item("NN").Value = GetNN(szNC, oWS.Range("C" & iRow).Value)
                  'Val function was added by anol 22 Apr 2015
                  adoRst.Fields.Item("BudgetAmt").Value = Val(oWS.Range("D" & iRow).Value)
               End If
               adoRst.Update
               adoRst.Close
            Else
               szBudgetID = UniqueID()
               adoRst.Open "SELECT * FROM GlobalSC;", adoConn1, adOpenDynamic, adLockOptimistic
               adoRst.AddNew
               adoRst.Fields.Item("BudgetID").Value = szBudgetID
               adoRst.Fields.Item("PropertyID").Value = txtPropertyName.Tag
               adoRst.Fields.Item("Fund").Value = GetFundID(szFund, oWS.Range("B" & iRow).Value)
               adoRst.Fields.Item("SCArea").Value = 1
               adoRst.Fields.Item("FinancialYear").Value = txtBudgetYears.Tag
               adoRst.Update
               adoRst.Close

               adoRst.Open "SELECT * FROM GlobalSCDtls;", adoConn1, adOpenDynamic, adLockOptimistic
               adoRst.AddNew
               adoRst.Fields.Item("BudgetDtlID").Value = UniqueID()
               adoRst.Fields.Item("BudgetID").Value = szBudgetID
               adoRst.Fields.Item("NC").Value = oWS.Range("C" & iRow).Value
               adoRst.Fields.Item("NN").Value = GetNN(szNC, oWS.Range("C" & iRow).Value)
               'Val function was added by anol 22 Apr 2015
               adoRst.Fields.Item("BudgetAmt").Value = Val(oWS.Range("D" & iRow).Value)
               adoRst.Update
               adoRst.Close
           End If
      iRow = iRow + 1
   Wend

   oWB.Close
   oXL.Quit

   Set oWS = Nothing
   Set oWB = Nothing
   Set oXL = Nothing

   Dim adoChild      As New ADODB.Recordset

   adoRst.Open "SELECT * FROM GlobalSC;", adoConn1, adOpenDynamic, adLockOptimistic

   While Not adoRst.EOF
      szSQL = "SELECT SUM(C.BudgetAmt) " & _
              "FROM GlobalSCDtls AS C " & _
              "WHERE C.BudgetID = '" & adoRst.Fields.Item("BudgetID").Value & "';"

      adoChild.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
        'Samrat 10/09/2014  VAL was crashing due to NULL value.
      If adoRst.Fields.Item("TotalBudget").Value <> Val(IIf(IsNull(adoChild.Fields.Item(0).Value), 0, adoChild.Fields.Item(0).Value)) Then
         adoRst.Fields.Item("TotalBudget").Value = Val(adoChild.Fields.Item(0).Value)
         adoRst.Fields.Item("PPSF").Value = adoRst.Fields.Item("TotalBudget").Value
         adoRst.Update
      End If
      
      adoChild.Close
      adoRst.MoveNext
   Wend


    'added by anol 06 Aug 2015 issue 571
   Update_SC_Lease adoConn1
   Set adoRst = Nothing
   Set adoChild = Nothing
'    adoConn1.Close
'    adoConn1.Open getConnectionString
    ConfigFlxSCBudgetDetails
    LoadFlxSCBudgetDetails adoConn1
    LoadMatrix adoConn1

    adoConn1.Close
 
'    SCSumTotal


   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If txtBudgetYears.Tag = flxSCBudgetDetails.TextMatrix(iRow, 9) And txtPropertyName.Tag = flxSCBudgetDetails.TextMatrix(iRow, 1) Then
            flxSCBudgetDetails.RowHeight(iRow) = 240 'added by anol 18 Jan 2016
            txtSCBudgetTotal.text = Val(txtSCBudgetTotal.text) + flxSCBudgetDetails.TextMatrix(iRow, 6)
            txtSCBudgetTotal.text = Format(txtSCBudgetTotal.text, "0.00")
      Else
         flxSCBudgetDetails.RowHeight(iRow) = 0
      End If
   Next iRow
   SCSumTotal
   fmeLoading.Visible = False
   cmdImport.Enabled = True
   MsgBox "Service Charge budget successfully imported.", vbInformation, "Import successful"
   Exit Sub
FileError:
   ShowMsgInTaskBar "File not does not exists", "Y", "N"
   fmeLoading.Visible = False
    cmdImport.Enabled = True
End Sub

Private Function GetNN(szNCList As String, szNC As String) As String
   Dim szaTemp() As String
   Dim szaTem1() As String
   Dim i         As Integer

   szaTemp = Split(szNCList, ", ")

   For i = 0 To UBound(szaTemp)
'      If InStr(szaTemp(i), szNC) > 0 Then
'         szaTem1 = Split(szaTemp(i), " # ")
'         GetNN = szaTem1(1)
'         Exit Function
'      End If
'Fixed by anol 15 Oct 2015
'While upload it was picking the wrong fund
      szaTem1 = Split(szaTemp(i), " # ")
      If Trim(szaTem1(0)) = Trim(szNC) Then
         GetNN = szaTem1(1)
         Exit Function
      End If
   Next i
   GetNN = "NULL"
End Function

Private Function GetFundID(szFundList As String, szFundName As String) As Single
   Dim szaTemp() As String
   Dim szaTem1() As String
   Dim i         As Integer

   szaTemp = Split(szFundList, ", ")

   For i = 0 To UBound(szaTemp)
'      If InStr(1, szaTemp(i), szFundName, vbBinaryCompare) > 0 Then
'         szaTem1 = Split(szaTemp(i), " # ")
'         GetFundID = Val(szaTem1(1))
'         Exit Function
'      End If
'Fixed by anol 15 Oct 2015
'While upload it was picking the wrong fund
    szaTem1 = Split(szaTemp(i), " # ")
    'Debug.Print szaTem1(0)
    If Trim(szaTem1(0)) = Trim(szFundName) Then
         GetFundID = Val(szaTem1(1))
         Exit Function
      End If
   Next i
   GetFundID = 0
End Function

Private Function getBudgetId(szPropID As String, szFundName As String, szFY As String, adoconn As ADODB.Connection) As String
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

   szSQL = "SELECT BudgetID " & _
           "FROM   GlobalSC AS G, Fund AS F " & _
           "WHERE  G.Fund = F.FundID AND " & _
                  "G.PropertyID = '" & szPropID & "' AND " & _
                  "F.FundCode = '" & szFundName & "' AND " & _
                  "G.FinancialYear = '" & szFY & "';"
                  
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      getBudgetId = adoRst.Fields.Item(0).Value
   Else
      getBudgetId = "NULL"
   End If

   adoRst.Close
   Set adoRst = Nothing
End Function


Private Sub rollBackServiceChargeDetails()
'This is an unused procedure
   Dim i As Integer

   For i = 0 To 59
      If detailsMatrix(txtMatrixRow.text, i).getBudgetId <> "" Then
         Set detailsMatrix(txtMatrixRow.text, i) = New clsSCDtl
      Else
         Exit For
      End If
   Next i
End Sub

Private Sub cmdSCBdClose_Click()
   Me.Hide
   Unload Me
End Sub

Private Sub cmdSCBdDelete_Click()
   If MsgBox("Do you want to delete the selected Service Charge Budget Detail?", vbQuestion + vbYesNo, "Saving") = vbNo Then
         Exit Sub
   End If
    lblLoading.Caption = "Please wait while deleting..."
    fmeLoading.Visible = True
    fmeLoading.Refresh
    cmdSCBdDelete.Enabled = False
   If Trim(flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)) <> "" Then
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X"
    'Modified by anol 11 Sep 2014
   'issue 471 production
      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = "X"
     'End of modification
      flxSCBudgetDetails.RowHeight(flxSCBudgetDetails.row) = 0
   Else
      flxSCBudgetDetails.RemoveItem (flxSCBudgetDetails.row)
   End If
   flgChange = 1
    'Modified by anol 11 Sep 2014
   'issue 471
   Call cmdSCBDSave_Click
   SCSumTotal
   'End of modification
   ControlsModeRentBudgetDetails DefaultMode
   fmeLoading.Visible = False
   cmdSCBdDelete.Enabled = True
   MsgBox "Selected Service Charge Budget has been deleted", vbInformation, "Deleted successfully"
   
End Sub

Private Sub cmdSCBDEdit_Click()
   Dim col As Integer

      updateGrid

      ReDim bufferMatrix(0) As clsSCDtl
      flgChange = 1

      SCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
'   End If
End Sub

Private Sub cmdSCBdNew_Click()
   If txtBudgetYears.text = "" Then
      ShowMsgInTaskBar "Please select a financial year", "Y", "N"
      cmdBudgetYears.SetFocus
      Exit Sub
   End If
      'Resolved by BOSL
      'issue number 471 Note 724
      'Added by anol 05 Nov 2014
      Dim strCheck As String
      Dim rsSQL As New ADODB.Recordset
      Dim adoConn1 As New ADODB.Connection
      If adoConn1.State = 0 Then
         adoConn1.Open getConnectionString
      End If
      szSQL = "SELECT P.CBY " & _
         "FROM Property P " & _
         "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
      rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
      If Not rsSQL.EOF Then
            strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
      End If
      If strCheck = "" Then
            MsgBox "A service charge budget year has not been set for this property." & vbCrLf & "Please set a service charge budget year in the global data screen."
            Exit Sub
      End If
      If adoConn1.State = 1 Then
            adoConn1.Close
      End If
      'End of Modification
   flgChange = 1

   txtBudgetId.text = UniqueID()
   bNew = True
    With frmServiceChargeDetails
            .szPropertySelection1 = txtPropertyName.Tag
              .szClientID = txtClientList.Tag
     End With
   Load frmServiceChargeDetails
   frmServiceChargeDetails.Show
   Me.Enabled = False
End Sub
Private Sub Update_SC_Lease(adoconn As ADODB.Connection)
   Dim Rst     As New ADODB.Recordset
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   
   
On Error GoTo ErrHandler:

'Resolved By BOSL.
'Modified By Asif. Issue: 0000519. Date: 04-Jan-2015
'Updating the service charge budgets through SQL rather than iteration which is time consuming.

' Charging Method: 2
    szSQL = "UPDATE LServiceCharges " & _
    "SET " & _
    "LServiceCharges.SCTotal = 0, " & _
    "LServiceCharges.SCAmount = 0 " & _
    "Where " & _
    "LServiceCharges.ChargingMethod = 2;"
    
    adoconn.Execute szSQL
    
    '******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 AND FN.CategoryCode<>5 ;"
    
    adoconn.Execute szSQL
    
    '******  fund  category = 5
     szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.TotalBudget * LServiceCharges.CMFigure)/100, " & _
    "LServiceCharges.SCAmount = (GSC.TotalBudget * LServiceCharges.CMFigure / 100) / F.PartOfYear " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 2 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL
    
    
    
    ' Charging Method: 4
    
    szSQL = "UPDATE LServiceCharges " & _
    "SET " & _
    "LServiceCharges.SCTotal = 0, " & _
    "LServiceCharges.SCAmount = 0, " & _
    "LServiceCharges.CMFigure = 0 " & _
    "Where " & _
    "LServiceCharges.ChargingMethod = 4;"
    
    adoconn.Execute szSQL
    
    '******  fund  category not 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P, Fund FN  " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 AND FN.CategoryCode<>5;"
    
    adoconn.Execute szSQL
      '******  fund  category = 5
    szSQL = "UPDATE LServiceCharges, GlobalSC AS GSC, LeaseDetails AS L, " & _
    "Units AS U, Frequencies AS F,  Property AS P , Fund FN " & _
    "SET " & _
    "LServiceCharges.SCTotal = (GSC.PPSF * U.TotalArea), " & _
    "LServiceCharges.SCAmount = (GSC.PPSF * U.TotalArea)/F.PartOfYear, " & _
    "LServiceCharges.CMFigure = (GSC.PPSF * U.TotalArea) " & _
    "Where  FN.FundID=GSC.Fund AND " & _
    "cstr(GSC.Fund) = LServiceCharges.ServiceChargeDept AND L.UnitNumber = U.UnitNumber AND " & _
    "L.LeaseID = LServiceCharges.LeaseID AND " & _
    "LServiceCharges.SCFrequency = F.ID AND U.PropertyID = GSC.PropertyID AND " & _
    "P.CBY = GSC.FinancialYear AND U.PropertyID = P.PropertyID AND " & _
    "LServiceCharges.ChargingMethod = 4 and GSC.SCYearEndGenerated=false AND FN.CategoryCode=5;"
    
    adoconn.Execute szSQL



'''''END OF MODIFICATIONS - ASIF

'   'issue 519
'   'Modified by anol 01  Jan 2015
'   Debug.Print Now
''      szSQL = "Update GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
''                        "Frequencies AS F, LServiceCharges AS LSC, Property AS P " & _
''                        "SET LSC.SCTotal=GSC.TotalBudget*LSC.CMFigure/100, " & _
''                        "LSC.SCAmount=(GSC.TotalBudget*LSC.CMFigure/100)/F.PartOfYear where " & _
''                        "L.UnitNumber = U.UnitNumber AND " & _
''                        "L.LeaseID = LSC.LeaseID AND " & _
''                        "LSC.SCFrequency = F.ID AND " & _
''                        "U.PropertyID = GSC.PropertyID AND " & _
''                        "P.CBY = GSC.FinancialYear AND " & _
''                        "U.PropertyID = P.PropertyID AND LSC.ChargingMethod = 4;"
''           adoConn.Execute szSQL
''
''    szSQL = "Update GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
''                        "Frequencies AS F, LServiceCharges AS LSC, Property AS P " & _
''                        "SET LSC.SCTotal=GSC.PPSF*U.TotalArea, " & _
''                        "LSC.SCAmount=(GSC.PPSF*U.TotalArea/F.PartOfYear), " & _
''                        "LSC.CMFigure=GSC.PPSF*U.TotalArea where " & _
''                        "L.UnitNumber = U.UnitNumber AND " & _
''                        "L.LeaseID = LSC.LeaseID AND " & _
''                        "LSC.SCFrequency = F.ID AND " & _
''                        "U.PropertyID = P.PropertyID AND " & _
''                        "P.CBY = GSC.FinancialYear AND " & _
''                        "U.PropertyID = GSC.PropertyID AND LSC.ChargingMethod = 2;"
''                 adoConn.Execute szSQL
''              Debug.Print Now
'   'End of Code
'
'   szSQL = "SELECT * FROM LServiceCharges WHERE ChargingMethod = 4 OR ChargingMethod = 2;"
''   Debug.Print Now
'   With Rst
'      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'      While Not .EOF
'         If .Fields.Item("ChargingMethod").Value = 2 Then
''If .Fields.Item("LeaseID").Value = "01100303170211771318" Then         'samrat 09/09/2014
''MsgBox ""
''End If
'            szSQL = "SELECT GSC.TotalBudget, F.PartOfYear " & _
'                    "FROM   GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
'                        "Frequencies AS F, LServiceCharges AS LSC, " & _
'                        "Property AS P " & _
'                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
'                        "L.UnitNumber = U.UnitNumber AND " & _
'                        "L.LeaseID = LSC.LeaseID AND " & _
'                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
'                        "LSC.SCFrequency = F.ID AND " & _
'                        "U.PropertyID = GSC.PropertyID AND " & _
'                        "P.CBY = GSC.FinancialYear AND " & _
'                        "U.PropertyID = P.PropertyID;"
'
'
''Debug.Print szSQL
'            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'            If Not adoRst.EOF Then
'               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("TotalBudget").Value * _
'                                                .Fields.Item("CMFigure").Value / 100
'               .Fields.Item("SCAmount").Value = (adoRst.Fields.Item("TotalBudget").Value * _
'                                                (.Fields.Item("CMFigure").Value / 100)) / _
'                                                adoRst.Fields.Item("PartOfYear").Value
'
'            Else
'               .Fields.Item("SCTotal").Value = 0
'               .Fields.Item("SCAmount").Value = 0
'            End If
'            adoRst.Close
'         Else                             '.ChargingMethod = 4
'            szSQL = "SELECT GSC.PPSF, U.TotalArea, F.PartOfYear " & _
'                    "FROM GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
'                        "Frequencies AS F, LServiceCharges AS LSC, Property AS P " & _
'                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
'                        "L.UnitNumber = U.UnitNumber AND " & _
'                        "L.LeaseID = LSC.LeaseID AND " & _
'                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
'                        "LSC.SCFrequency = F.ID AND " & _
'                        "U.PropertyID = P.PropertyID AND " & _
'                        "P.CBY = GSC.FinancialYear AND " & _
'                        "U.PropertyID = GSC.PropertyID;"
'            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
''Debug.Print szSQL
'            If Not adoRst.EOF Then
'               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("PPSF").Value * _
'                                               adoRst.Fields.Item("TotalArea").Value
'               .Fields.Item("SCAmount").Value = adoRst.Fields.Item("PPSF").Value * _
'                                                adoRst.Fields.Item("TotalArea").Value / _
'                                                adoRst.Fields.Item("PartOfYear").Value
'               .Fields.Item("CMFigure").Value = adoRst.Fields.Item("PPSF").Value * _
'                                                adoRst.Fields.Item("TotalArea").Value
'
'            Else
'               .Fields.Item("SCTotal").Value = 0
'               .Fields.Item("SCAmount").Value = 0
'               .Fields.Item("CMFigure").Value = 0
'            End If
'            adoRst.Close
'         End If
'         .Update
'         .MoveNext
'      Wend
'      .Close
'   End With
''Debug.Print Now
'   Set adoRst = Nothing
'   Set Rst = Nothing

Exit Sub
ErrHandler:
   MsgBox ERR.Number & " " & ERR.description, vbExclamation + vbOKOnly, "Could not update Service Charge Budget"
End Sub
Private Sub cmdSCBDSave_Click()
   Dim i          As Integer
   Dim iRow       As Integer
   Dim iRowChild  As Integer
   Dim Rst        As New ADODB.Recordset
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString

   i = 0
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(iRow, 7) <> "X" Then
         SaveUpdateSC iRow, adoconn
         txtBudgetId.text = flxSCBudgetDetails.TextMatrix(iRow, 0)
         saveSCDetails i
      Else
         deleteSC flxSCBudgetDetails.TextMatrix(iRow, 0), adoconn
      End If
   i = i + 1
   Next iRow

   Update_SC_Lease adoconn

   LoadMatrix adoconn

   Set Rst = Nothing
   adoconn.Close
   flgChange = 0

   ControlsModeRentBudgetDetails SavedMode
End Sub

'Private Sub Update_SC_Lease()
'   Dim Rst     As New ADODB.Recordset
'   Dim adoRst  As New ADODB.Recordset
'
'   szSQL = "SELECT * FROM LServiceCharges WHERE ChargingMethod = 4 OR ChargingMethod = 2;"
'
'   With Rst
'      .Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
'
'      While Not .EOF
'         If .Fields.Item("ChargingMethod").Value = 2 Then
'            szSQL = "SELECT GSC.TotalBudget, F.PartOfYear " & _
'                    "FROM   GlobalSC AS GSC, LeaseDetails AS L, Units AS U, " & _
'                        "Frequencies AS F, LServiceCharges AS LSC " & _
'                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
'                        "L.UnitNumber = U.UnitNumber AND " & _
'                        "L.LeaseID = LSC.LeaseID AND " & _
'                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
'                        "LSC.SCFrequency = F.ID AND " & _
'                        "U.PropertyID = GSC.PropertyID;"
''Debug.Print szSQL
'            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'            If Not adoRst.EOF Then
'               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("TotalBudget").Value * _
'                                                .Fields.Item("CMFigure").Value / 100
'               .Fields.Item("SCAmount").Value = (adoRst.Fields.Item("TotalBudget").Value * _
'                                                (.Fields.Item("CMFigure").Value / 100)) / _
'                                                adoRst.Fields.Item("PartOfYear").Value
'            Else
'               .Fields.Item("SCTotal").Value = 0
'               .Fields.Item("SCAmount").Value = 0
'            End If
'            adoRst.Close
'         Else                             '.ChargingMethod = 4
'            szSQL = "SELECT GSC.PPSF, U.TotalArea, F.PartOfYear " & _
'                    "FROM GlobalSC AS GSC, LeaseDetails AS L, Units AS U " & _
'                        "Frequencies AS F, LServiceCharges AS LSC " & _
'                    "WHERE  GSC.Fund = " & .Fields.Item("ServiceChargeDept").Value & " AND " & _
'                        "L.UnitNumber = U.UnitNumber AND " & _
'                        "L.LeaseID = LSC.LeaseID AND " & _
'                        "LSC.ServiceCharge = '" & .Fields.Item("ServiceCharge").Value & "' AND " & _
'                        "LSC.SCFrequency = F.ID AND " & _
'                        "U.PropertyID = GSC.PropertyID;"
''Debug.Print szSQL
'            If Not adoRst.EOF Then
'               .Fields.Item("SCTotal").Value = adoRst.Fields.Item("PPSF").Value * _
'                                               adoRst.Fields.Item("TotalArea").Value
'               .Fields.Item("SCAmount").Value = adoRst.Fields.Item("PPSF").Value * _
'                                                adoRst.Fields.Item("TotalArea").Value / _
'                                                adoRst.Fields.Item("PartOfYear").Value
'               .Fields.Item("CMFigure").Value = adoRst.Fields.Item("PPSF").Value * _
'                                                adoRst.Fields.Item("TotalArea").Value
'
'            Else
'               .Fields.Item("SCTotal").Value = 0
'               .Fields.Item("SCAmount").Value = 0
'               .Fields.Item("CMFigure").Value = 0
'            End If
'            adoRst.Close
'         End If
'         .Update
'         .MoveNext
'      Wend
'      .Close
'   End With
'
'   Set adoRst = Nothing
'   Set Rst = Nothing
'End Sub

Private Function saveSCDetails(ByVal row As Integer)
   Dim col As Integer
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
   For col = 0 To 59
      If frmServiceCharge.getDetailsFromMatrix(row, col).getBudgetDetailID = "" Then
          Exit For
      Else
         If frmServiceCharge.getDetailsFromMatrix(row, col).getFlgDel = "D" Then
            deleteSCD frmServiceCharge.getDetailsFromMatrix(row, col), adoconn
         Else
            SaveUpdateSCD frmServiceCharge.getDetailsFromMatrix(row, col), adoconn
         End If
      End If
   Next col
   adoconn.Close
End Function

Private Function SaveUpdateSCD(ByVal budgetDetails As clsSCDtl, adoconn As ADODB.Connection)
   Dim Rst     As New ADODB.Recordset

   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   Rst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   szSQL = "SELECT * " & _
            "FROM GlobalSCDtls;"
   Rst.Open szSQL, adoconn, adOpenDynamic, adLockOptimistic

   Rst.AddNew
   Rst!BudgetDtlID = budgetDetails.getBudgetDetailID
   Rst!budgetId = budgetDetails.getBudgetId
   Rst!NN = budgetDetails.getNName
   Rst!NC = budgetDetails.getNCode
   Rst!BudgetAmt = budgetDetails.getBudgetAmount

   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSCD(ByVal budgetDetails As clsSCDtl, adoconn As ADODB.Connection)
   szSQL = "DELETE * " & _
            "FROM GlobalSCDtls " & _
            "WHERE BudgetDtlID = '" & budgetDetails.getBudgetDetailID & "' "
   adoconn.Execute szSQL
End Function

Private Function SaveUpdateSC(ByVal Index As Integer, adoconn As ADODB.Connection)
   Dim Rst     As New ADODB.Recordset
   'added by anol 20170310
   If flxSCBudgetDetails.TextMatrix(Index, 0) = "" Then Exit Function
   
   szSQL = "SELECT * " & _
            "FROM GlobalSC where budgetId='" & flxSCBudgetDetails.TextMatrix(Index, 0) & "';"
   Rst.Open szSQL, adoconn, adOpenStatic, adLockOptimistic

   If Rst.RecordCount < 1 Then
      Rst.AddNew
      Rst!budgetId = flxSCBudgetDetails.TextMatrix(Index, 0)
      Rst!FinancialYear = txtBudgetYears.Tag 'If frmMMain.IsRibbonVersion Then
      Rst!propertyID = txtPropertyName.Tag
   Else
      Rst!propertyID = flxSCBudgetDetails.TextMatrix(Index, 1)
   End If

'   If frmMMain.IsRibbonVersion Then
'   Rst!PropertyID = txtPropertyName.tag
   
'   Else
'      Rst!PropertyID = frmGlobal1.txtPropertyName.tag
'   End If
   Rst!Fund = flxSCBudgetDetails.TextMatrix(Index, 2)
   Rst!TotalBudget = flxSCBudgetDetails.TextMatrix(Index, 6)
   Rst!SCArea = flxSCBudgetDetails.TextMatrix(Index, 5)
   Rst!ppsf = flxSCBudgetDetails.TextMatrix(Index, 6)

   Rst.Update
   Rst.Close
   Set Rst = Nothing
End Function

Private Function deleteSC(ByVal bId As String, adoconn As ADODB.Connection)
   szSQL = "DELETE * " & _
            "FROM GlobalSC " & _
            "WHERE BudgetID = '" & bId & "';"
   adoconn.Execute szSQL

'   szSQL = "DELETE GlobalSCDtls.* " & _
'           "FROM   GlobalSCDtls " & _
'           "WHERE  GlobalSCDtls.BudgetID NOT IN (" & _
'               "SELECT GlobalSC.BudgetID " & _
'               "FROM GlobalSC);"
'issue  381 SQL that is taking long time to complete 20170512 fixed by anol
  szSQL = "DELETE GlobalSCDtls.* From GlobalSCDtls LEFT JOIN GlobalSC ON GlobalSCDtls.BudgetID =GlobalSC.BudgetID WHERE  GlobalSC.BudgetID  IS NULL;"
   adoconn.Execute szSQL
End Function

'Private Sub cmdSCClose_Click()
'   initialiseGrid
'   Me.Hide
'End Sub

Private Sub ConfigFlxSCBudgetDetails()
   Dim szFlxHeader As String

   flxSCBudgetDetails.Rows = 1
   flxSCBudgetDetails.RowHeight(0) = 0
   flxSCBudgetDetails.Clear
   flxSCBudgetDetails.Cols = 11
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|<FundCODE|>FundName|>TotalBudget|>SCArea|>PPSF|FY"
   flxSCBudgetDetails.FormatString = szFlxHeader$

   flxSCBudgetDetails.ColWidth(0) = 0
   flxSCBudgetDetails.ColWidth(1) = 0
   flxSCBudgetDetails.ColAlignment(1) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(2) = 0
   flxSCBudgetDetails.ColWidth(3) = 2000 'lblRentCharges(2).Left - lblRentCharges(0).Left
    flxSCBudgetDetails.ColAlignment(3) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(4) = lblRentCharges(2).Left - lblRentCharges(0).Left
    flxSCBudgetDetails.ColAlignment(4) = vbLeftJustify
   flxSCBudgetDetails.ColWidth(5) = lblRentCharges(3).Left - lblRentCharges(2).Left + 200
   flxSCBudgetDetails.ColWidth(6) = lblRentCharges(4).Left - lblRentCharges(3).Left - 200
   flxSCBudgetDetails.ColWidth(7) = lblRentCharges(4).Left - lblRentCharges(3).Left 'flxSCBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxSCBudgetDetails.ColWidth(8) = 0
   flxSCBudgetDetails.ColWidth(9) = 0
   flxSCBudgetDetails.ColWidth(10) = 0

   txtSCBudgetTotal.Width = flxSCBudgetDetails.ColWidth(7)
   txtSCBudgetTotal.Left = lblRentCharges(2).Left
   
   txtSCTotalArea.Width = flxSCBudgetDetails.ColWidth(6)
   txtSCTotalArea.Left = lblRentCharges(3).Left
End Sub
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 80
   flxClient.ColWidth(1) = 2500
   flxClient.ColWidth(2) = 3500
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   
   txtSearchClientID.Width = 1530
   txtSearchClientName.Visible = True
   'picClient.Width = 5295
   'cmdPicCLose.Left = 5010
   txtSearchClientID.Left = 45
   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "Financial Year"
   lblClientName.Caption = "Financial Year Description"

   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   adoconn.Open getConnectionString
           
        szSQL = "SELECT FYrID, FinancialYear, FY_Description " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.PropertyID = '" & txtPropertyName.Tag & "' order by FinancialYear Desc ;"


   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
      ShowMsgInTaskBar "Financial year has not been created.", "Y", "N"
   Else
            rRow = 1

        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = " " & Trim(rstRec.Fields.Item("FYrID").Value)
           flxClient.TextMatrix(rRow, 1) = Trim(rstRec.Fields.Item("FinancialYear").Value)
           flxClient.TextMatrix(rRow, 2) = Trim(rstRec.Fields.Item("FY_Description").Value)
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadFY(adoconn As ADODB.Connection)
  ' Dim rRow    As Integer
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String


        'Resolved by BOSL
        'issue no 471
        'Modified by anol 11 Sep 2014
        Dim rstSQL  As New ADODB.Recordset
        szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID " & _
                "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
                "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
        rstSQL.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
        If Not rstSQL.EOF Then
            txtBudgetYears.Tag = IIf(IsNull(rstSQL.Fields.Item("FYrID").Value), "", rstSQL.Fields.Item("FYrID").Value)
             txtBudgetYears.text = IIf(IsNull(rstSQL.Fields.Item("CBY").Value), "", rstSQL.Fields.Item("CBY").Value)
        Else
            txtBudgetYears.text = ""
            txtBudgetYears.Tag = ""
           ' ShowMsgInTaskBar "Please set the financial year for this property inGlobal Data.", , "N"
        End If
        'End of modification
   'End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading financial year.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub LoadFlxSCBudgetDetails(adoconn As ADODB.Connection)
'   On Error GoTo Err
   Dim i       As Integer
   Dim Rst     As New ADODB.Recordset
   Call ConfigFlxSCBudgetDetails
'   szSQL = "SELECT g.*, f.FundName " & _
'            "FROM GlobalSC g, Fund f " & _
'            "WHERE CInt(g.Fund)=f.FundId;"
'Modified by anol 20161013
 szSQL = "SELECT g.*, f.FundName,f.FundCode " & _
            "FROM GlobalSC g, Fund f " & _
            "WHERE CInt(g.Fund)=f.FundId AND G.PropertyID='" & txtPropertyName.Tag & "';"
            
   Rst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   With flxSCBudgetDetails
      i = 1
      If Not Rst.EOF Then
       ' txtSCBudgetTotal.text = "0"
         While Not Rst.EOF
            .AddItem ""
            .TextMatrix(i, 0) = Rst!budgetId
            .TextMatrix(i, 1) = Rst!propertyID
            .TextMatrix(i, 2) = Rst!Fund ' fund is fund Id in globalSC
            .TextMatrix(i, 3) = Rst!FundCode
            .TextMatrix(i, 4) = Rst!FundName
            .TextMatrix(i, 5) = Rst!SCArea
           ' txtSCBudgetTotal.text = Format(Rst!TotalBudget, "0.00") + Val(txtSCBudgetTotal.text)
            .TextMatrix(i, 6) = Format(Rst!TotalBudget, "0.00") '
            .TextMatrix(i, 7) = Format(Rst!ppsf, "0.00")
            .TextMatrix(i, 8) = i - 1
            .TextMatrix(i, 9) = IIf(IsNull(Rst!FinancialYear), "", Rst!FinancialYear)
            .RowHeight(i) = 0
            i = i + 1
            Rst.MoveNext
         Wend

      End If
       ' txtSCBudgetTotal.text = Format(txtSCBudgetTotal.text, "0.00")
      initialiseMatrix
      .row = 0
      .col = 0
   End With

   Rst.Close
   Set Rst = Nothing
   Exit Sub
ERR:
   MsgBox ERR.description & "Error form LoadFlxSCBudgetDetails"
End Sub

Private Sub LoadMatrix(adoconn As ADODB.Connection)
   Dim i As Integer
   
    
   For i = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.TextMatrix(i, 0) <> "" Then
         PopulateMatrix flxSCBudgetDetails.TextMatrix(i, 0), Val(flxSCBudgetDetails.TextMatrix(i, 8)), adoconn
      End If
   Next i
End Sub

Private Sub PopulateMatrix(bId As String, row As Integer, adoconn As ADODB.Connection)
   Dim i    As Integer
   Dim Rst  As New ADODB.Recordset

   szSQL = "SELECT g.* " & _
            "FROM  GlobalSCDtls AS g " & _
            "WHERE g.BudgetID = '" & bId & "';"
   Rst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   i = 0
   If Not Rst.EOF Then
      While Not Rst.EOF
         getDetailsFromMatrix(row, i).setBudgetDetailId Rst!BudgetDtlID
         getDetailsFromMatrix(row, i).setBudgetId Rst!budgetId
         getDetailsFromMatrix(row, i).setNCode Rst!NC
         getDetailsFromMatrix(row, i).setNName IIf(IsNull(Rst!NN), "", Rst!NN)
         'Below line fixed by anol null was handled 22 Apr 2015
         getDetailsFromMatrix(row, i).setBudgetAmount Format(IIf(IsNull(Rst!BudgetAmt), "0", Rst!BudgetAmt), "0.00")
         'txtSCBudgetTotal.text = Val(txtSCBudgetTotal.text) + Format(IIf(IsNull(Rst!BudgetAmt), "0", Rst!BudgetAmt), "0.00")
         i = i + 1
         Rst.MoveNext
      Wend
   End If

   Rst.Close
   Set Rst = Nothing
End Sub

Public Sub SCSumTotal()
   Dim iRow As Integer

   txtSCBudgetTotal.text = "0.00"
   txtSCTotalArea.text = "0"
   For iRow = 1 To flxSCBudgetDetails.Rows - 1
      If flxSCBudgetDetails.RowHeight(iRow) > 0 Then
         txtSCBudgetTotal.text = Format(Val(txtSCBudgetTotal.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 6)), "0.00")
         txtSCTotalArea.text = Val(txtSCTotalArea.text) + Val(flxSCBudgetDetails.TextMatrix(iRow, 5))
      End If
   Next iRow
End Sub

Private Sub Command1_Click()
   cboBudgetYears_Change
End Sub

Private Sub flxSCBudgetDetails_DblClick()
    If flxSCBudgetDetails.RowHeight(flxSCBudgetDetails.row) = 0 Then
        ShowMsgInTaskBar "There is no budget entered for this budget Year.", , "N"
        ControlsModeRentBudgetDetails DefaultMode
        Exit Sub
    End If
   Load frmServiceChargeDetails
  'issue 471 Note 6
   'Modified by anol 14 Sep 2014
   bEDIT = True
   'end
   'issue 471 Note 740
   txtBudgetId.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)
   'End of modification
   With frmServiceChargeDetails
      .txtSCFund.Tag = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2)
      .txtSCFundCode.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 3)
      .txtSCFund.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4)
'      .cboSCFund.Locked = True
      .szPropertySelection1 = txtPropertyName.Tag
      .szClientID = txtClientList.Tag
      .txtTotalArea.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5)
      .txtSCDBudgetTotal.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6)
      .txtPpsf.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7)
      .LoadFlxSCBudgetDetailsAnalysis
      LoadForm frmServiceChargeDetails
   End With

   
   bEDIT = True
   Me.Enabled = False
End Sub

Private Sub flxSCBudgetDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub flxSCBudgetDetails_RowColChange()
   If flxSCBudgetDetails.TextMatrix(1, 0) = "" Then Exit Sub
   If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 7) = "X" Then Exit Sub

   On Error Resume Next

'   cboSCFund.Value = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2)
'   txtBudget.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4)
'   txtTotalArea.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5)
'   txtPpsf.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6)
   txtMatrixRow.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 8)
   txtBudgetId.text = flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 0)

   ControlsModeRentBudgetDetails GridRowOnSelection

'   cmdSCBdEdit.SetFocus
End Sub

Private Function hasNoLines() As Boolean
   Dim i As Integer

   hasNoLines = True
   i = 0
   While i < 60
      If detailsMatrix(flxSCBudgetDetails.row - 1, i).getBudgetId <> "" Then
         If detailsMatrix(flxSCBudgetDetails.row - 1, i).getFlgDel <> "X" Then
            hasNoLines = False
            Exit Function
         End If
      Else
         Exit Function
      End If
      i = i + 1
   Wend
End Function

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   On Error Resume Next

   Select Case mode
      Case ComponentMode.DefaultMode
'         cboSCFund.text = ""
'         cboSCFund.Enabled = False
'         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
'         txtBudget.Locked = True
'         txtTotalArea.text = ""
'         txtTotalArea.Locked = True
'         txtPpsf.text = ""
'         txtPpsf.Locked = True

'         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
'         cmdSCBdEdit.Enabled = False
         'cboBudgetYears.Enabled = True
'         cmdSCBdEdit.Caption = "&Edit"
'         cmdSCBdEdit.ToolTipText = "Edit selected budget"
'         If flgChange = 1 Then
''            cmdSCBdSave.Enabled = True
'            cmdSCBdCancel.Enabled = True
'         Else
''            cmdSCBdSave.Enabled = False
'            cmdSCBdCancel.Enabled = False
'         End If
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdNew.SetFocus

      Case ComponentMode.SavedMode
'         cboSCFund.text = ""
'         cboSCFund.Enabled = False
'         txtBudget.text = ""
         txtBudgetId.text = ""
         txtMatrixRow.text = ""
'         txtBudget.Locked = True
'         txtTotalArea.text = ""
'         txtTotalArea.Locked = True
'         txtPpsf.text = ""
'         txtPpsf.Locked = True

'         cmdDetails.Enabled = False
         cmdSCBdNew.Enabled = True
'         cmdSCBdEdit.Enabled = False
        ' cboBudgetYears.Enabled = True
'         cmdSCBdEdit.Caption = "&Edit"
'         cmdSCBdEdit.ToolTipText = "Edit selected budget"
'         If flgChange = 1 Then
''            cmdSCBdSave.Enabled = True
'            cmdSCBdCancel.Enabled = True
'         Else
''            cmdSCBdSave.Enabled = False
'            cmdSCBdCancel.Enabled = False
'         End If

         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = True
         flxSCBudgetDetails.row = 0
         flxSCBudgetDetails.col = 0
         cmdSCBdClose.SetFocus

      Case ComponentMode.EditMode

'         cboSCFund.Enabled = True
'         txtTotalArea.Locked = False
'         txtBudget.Locked = False
'         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
'         cmdSCBdEdit.Caption = "&Update"
'         cmdSCBdEdit.ToolTipText = "Update selected budget"
'         cmdSCBdEdit.Enabled = True
'         cmdSCBdSave.Enabled = False
'         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False

         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.NewEntryMode
'         cboSCFund.text = ""
'         cboSCFund.Enabled = True
'         txtBudget.text = ""
'         txtTotalArea.text = ""
'         txtTotalArea.Locked = False
'         txtPpsf.text = ""

'         cmdDetails.Enabled = True
         cmdSCBdNew.Enabled = False
'         cmdSCBdEdit.Caption = "&Update"
'         cmdSCBdEdit.Enabled = True
         'cboBudgetYears.Enabled = False
'         cmdSCBdSave.Enabled = False
'         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = False
         flxSCBudgetDetails.Enabled = False

      Case ComponentMode.GridRowOnSelection
'         txtBudget.Locked = True
         cmdSCBdNew.Enabled = True
'         cmdSCBdEdit.Enabled = True
            'Resolved by BOSL
            'issue number 471
            'Modified by Anol 09 Sep 2014
            'cboBudgetYears.Enabled = False
            'End of modification

        If flgChange = 1 Then
'            cmdSCBdSave.Enabled = True
         Else
''            cmdSCBdSave.Enabled = False
         End If
'         cmdSCBdCancel.Enabled = True
         cmdSCBdDelete.Enabled = True
   End Select
End Sub

Private Sub Form_Activate()
        bFormLoaded = True
            'If bLoadingProperty Then
        Dim strCheck As String
        Dim rsSQL As New ADODB.Recordset
        Dim adoConn1 As New ADODB.Connection
        If adoConn1.State = 0 Then
            adoConn1.Open getConnectionString
        End If
        
        szSQL = "SELECT P.CBY FROM Property P WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
        rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
        If Not rsSQL.EOF Then
              strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
        End If
        If strCheck = "" Then
                ShowMsgInTaskBar "A service charge budget year has not been set for this property." & vbCrLf & "Please set a service charge budget year in the global data screen.", "Y", "N"
        End If
        
         'End If
          If adoConn1.State = 1 Then
              adoConn1.Close
        End If
'         Dim i       As Integer
'    Dim Rst     As New ADODB.Recordset
'    Call ConfigFlxSCBudgetDetails
'    '   szSQL = "SELECT g.*, f.FundName " & _
'    '            "FROM GlobalSC g, Fund f " & _
'    '            "WHERE CInt(g.Fund)=f.FundId;"
'    'Modified by anol 20161013
'    szSQL = "SELECT g.*, f.FundName,f.FundCode,P.clientID " & _
'            "FROM GlobalSC g, Fund f,Property P " & _
'            "WHERE CInt(g.Fund)=f.FundId AND P.propertyID=G.propertyID and F.CategoryCode=5 and P.ClientID='" & txtClientList.Tag & "';"
'
'    Rst.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
'    If Rst.EOF Then
'            cmdRunFlag.Enabled = True
'    End If
'    Rst.Close
'    If adoConn1.State = 1 Then
'              adoConn1.Close
'        End If
End Sub

Private Sub Form_Load()
   ' On Error GoTo ERR
    Dim errmsg As String
    bFormLoaded = False
    Me.Width = 9720
    Me.Height = 7650
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
    Me.Refresh
    Me.BackColor = MODULEBACKCOLOR
    Frame1(1).BackColor = Me.BackColor
    bEDIT = False
    bNew = False
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    'added by anol 06 Sep 2016
    Dim strID As String
    Dim rsID As New ADODB.Recordset
'    adoConn.Execute "Delete from GlobalSCDtls where NC is null"
    rsID.Open "Select BudgetID from GlobalSCDtls where NC is null", adoconn, adOpenKeyset, adLockReadOnly
    If Not rsID.EOF Then
        strID = rsID.Fields("BudgetID").Value
        adoconn.Execute "Delete from GlobalSCDtls where BudgetID= '" & strID & "'"
        adoconn.Execute "Delete from GlobalSC where BudgetID= '" & strID & "'"
    End If
    
    
    LoadCmbClient adoconn 'LOADING CLIENT , PROPERTY AND BudgetYears
    errmsg = "initialiseGrid"
    'initialiseGrid
    
    flgChange = 0
    '  Loading Clients and Properties
    'PrepareList adoConn, cboClient, cboProperty
    LoadFlxSCBudgetDetails adoconn
    LoadMatrix adoconn
    LoadFund adoconn
    LoadFY adoconn
    adoconn.Close
    
    ControlsModeRentBudgetDetails DefaultMode
    errmsg = "SCSumTotal"
    SCSumTotal
    cboBudgetYears_Change
   
    Call WheelHook(Me.hWnd)
   
    Exit Sub
ERR:
   MsgBox ERR.description & "-From the form load" & errmsg
End Sub

Private Sub initialiseGrid()

'   Call ConfigFlxSCBudgetDetails
'
'
''   flgEdit = 0
'   flgChange = 0
'
''   flgNew = 0
'   If adoConn.State = 0 Then
'        adoConn.Open getConnectionString
'   End If
'
''  Loading Clients and Properties
'   'PrepareList adoConn, cboClient, cboProperty
'
'   LoadFlxSCBudgetDetails
'   LoadMatrix
'   LoadFund
'   LoadFY
'
'   adoConn.Close
'
'   ControlsModeRentBudgetDetails DefaultMode
End Sub
'
'Private Function TotalSCProperty(szPropertyID As String) As Double
'   Dim Rst2 As New ADODB.Recordset
'
'   szSQL = "SELECT GlobalSC.PropertyID, SUM(GlobalSC.TotalBudget) AS TOTALRENT " & _
'           "From   GlobalSC " & _
'           "WHERE  GlobalSC.PropertyID = '" & szPropertyID & "' " & _
'           "GROUP BY GlobalSC.PropertyID;"
''Debug.Print szSQL
'   Rst2.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Not Rst2.EOF Then
'      TotalSCProperty = CDbl(Rst2!TOTALRENT)
'   Else
'      TotalSCProperty = 0
'   End If
'
'   Rst2.Close
'   Set Rst2 = Nothing
'End Function

Private Sub LoadFund(adoconn As ADODB.Connection)
   ' Error Handler
   On Error GoTo Error_Handler
   
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   szSQL = "SELECT FundID, FundName " & _
           "FROM Fund " & _
           "WHERE CategoryCode = 2;"

   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
   Else
      ReDim Data(2, adoRst.RecordCount) As String
    
      rRow = 0
      While Not adoRst.EOF
         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
         rRow = rRow + 1
         adoRst.MoveNext
      Wend
    
'      cboSCFund.Clear
'      cboSCFund.Column() = Data()
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:

   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If frmMMain.IsRibbonVersion Then
   Enabled = True
'   Else
'   frmGlobal1.Enabled = True
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set adoConn = Nothing
    UnLoadForm Me
'   Call WheelUnHook(Me.hwnd)
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub
'
'Private Sub txtBudget_Change()
'   If Trim(txtBudget.text) <> "" And Val(Trim(txtBudget.text)) > 0 Then
'    txtBudget.Locked = False
'    computePpsf
'   End If
'End Sub
'
'Private Sub txtBudget_KeyPress(KeyAscii As Integer)
'   If Not txtBudget.Locked Then DigitTextKeyPress txtBudget, KeyAscii
'End Sub
'
'Private Sub txtBudget_LostFocus()
'   computePpsf
''   If flgEdit = 1 Or flgNew = 1 Then txtBudget.BackColor = &H80000005
'End Sub
'
'Private Sub computePpsf()
'On Error GoTo ErrHandler
'   If Trim(txtBudget.text) <> "" And Trim(txtTotalArea.text) <> "" Then
'      Dim ppsf As Double
'
'      ppsf = CDbl(txtBudget.text) / CDbl(Trim(txtTotalArea.text))
'      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
'   End If
'   Exit Sub
'ErrHandler:
'   ShowMsgInTaskBar "Please ensure that the values for Total Budget and Total Area are valid before continuing", , "N"
'End Sub
'
'Private Sub txtTotalArea_LostFocus()
'   computePpsf
'End Sub

Private Sub updateGrid()
'   If flgEdit = 0 Then
'      flxSCBudgetDetails.AddItem ""
'      If flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) <> "" Then flxSCBudgetDetails.AddItem ""
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 0) = txtBudgetId.text
''      If frmMMain.IsRibbonVersion Then
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 1) = txtPropertyName.tag
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 2) = CInt(cboSCFund.Value)
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 3) = cboSCFund.text
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 5) = txtTotalArea.text
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 6) = txtPpsf.text
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.Rows - 1, 8) = nextRow
'   Else
''      If frmMMain.IsRibbonVersion Then
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = txtPropertyName.tag
''      Else
''      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 1) = frmGlobal1.txtPropertyName.tag
''      End If
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 2) = CInt(cboSCFund.Value)
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 3) = cboSCFund.text
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 5) = txtTotalArea.text
'      flxSCBudgetDetails.TextMatrix(flxSCBudgetDetails.row, 6) = txtPpsf.text
'   End If
'   flgEdit = 0
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




