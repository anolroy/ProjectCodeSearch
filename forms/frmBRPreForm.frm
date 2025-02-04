VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBRPreForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Receipt"
   ClientHeight    =   7920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13425
   Icon            =   "frmBRPreForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13425
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   5085
      ScaleHeight     =   4740
      ScaleWidth      =   6120
      TabIndex        =   17
      Top             =   1890
      Visible         =   0   'False
      Width           =   6150
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
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   25
         Top             =   675
         Width           =   6030
         _ExtentX        =   10636
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
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   22
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
         Left            =   1830
         TabIndex        =   21
         Top             =   90
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
         TabIndex        =   20
         Top             =   375
         Width           =   1575
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2778;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   19
         Top             =   375
         Width           =   4365
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "7699;450"
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
         Width           =   5760
      End
   End
   Begin VB.Frame fraReference 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   930
      TabIndex        =   15
      Top             =   2910
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtCheqNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1050
         TabIndex        =   8
         Top             =   0
         Width           =   1740
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Receipt Method"
      Height          =   615
      Left            =   5250
      TabIndex        =   9
      Top             =   6630
      Visible         =   0   'False
      Width           =   2625
      Begin VB.OptionButton optBR_Bank 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bank"
         Height          =   300
         Left            =   1440
         TabIndex        =   11
         Top             =   220
         Width           =   735
      End
      Begin VB.OptionButton optBR_Cheque 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cheque"
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   220
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   5235
      Left            =   55
      TabIndex        =   26
      Top             =   0
      Width           =   7980
      Begin VB.CommandButton cmdClient 
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
         Left            =   6450
         TabIndex        =   0
         Top             =   225
         Width           =   300
      End
      Begin VB.CommandButton cmdproperty 
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
         Left            =   6450
         TabIndex        =   1
         Top             =   675
         Width           =   300
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   3510
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2745
         TabIndex        =   6
         Top             =   3510
         Width           =   1215
      End
      Begin VB.CheckBox chkMultiple 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Multiple"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Frame fraMultiple 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   270
         TabIndex        =   27
         Top             =   2205
         Width           =   2775
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1050
            TabIndex        =   4
            Top             =   0
            Width           =   1480
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   375
         End
         Begin MSForms.Label lblPostingDate 
            Height          =   300
            Left            =   2535
            TabIndex        =   28
            Top             =   0
            Width           =   225
            ForeColor       =   8421504
            BackColor       =   16761024
            Caption         =   " P"
            Size            =   "397;529"
            FontName        =   "Myriad Web"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin MSForms.TextBox txtClient 
         Height          =   285
         Left            =   1335
         TabIndex        =   38
         Top             =   225
         Width           =   5130
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "9049;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   1320
         TabIndex        =   37
         Top             =   675
         Width           =   5130
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "9049;503"
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
         Index           =   11
         Left            =   270
         TabIndex        =   36
         Top             =   255
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   12
         Left            =   270
         TabIndex        =   35
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   34
         Top             =   1095
         Width           =   360
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There are saved receipts in this batch."
         Height          =   195
         Left            =   420
         TabIndex        =   33
         Top             =   4230
         Width           =   2715
      End
      Begin MSForms.ComboBox cmbBankAc 
         Height          =   360
         Left            =   2445
         TabIndex        =   2
         Top             =   1095
         Width           =   4290
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "7567;635"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   4
         ListRows        =   20
         cColumnInfo     =   4
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058;3527;0;0"
      End
      Begin MSForms.Label Label13 
         Height          =   345
         Index           =   7
         Left            =   1320
         TabIndex        =   32
         Top             =   1110
         Width           =   1125
         BackColor       =   16777215
         Size            =   "1984;617"
         BorderStyle     =   1
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbFund 
         Height          =   360
         Left            =   1320
         TabIndex        =   31
         Top             =   3150
         Visible         =   0   'False
         Width           =   1770
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3122;635"
         BoundColumn     =   2
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;4586"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Fund"
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   30
         Top             =   3195
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Ref:"
      Height          =   195
      Index           =   2
      Left            =   930
      TabIndex        =   14
      Top             =   6750
      Width           =   675
   End
   Begin MSForms.ComboBox cboAutoRef_ 
      Height          =   300
      Left            =   1890
      TabIndex        =   7
      Top             =   6750
      Width           =   2625
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "4630;529"
      TextColumn      =   2
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;7055"
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   930
      TabIndex        =   13
      Top             =   6315
      Width           =   1215
   End
   Begin VB.Label lblRef 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
      Height          =   255
      Left            =   930
      TabIndex        =   12
      Top             =   6390
      Width           =   1215
   End
End
Attribute VB_Name = "frmBRPreForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim szChoice As String, szaChoice() As String, bLoaded As Boolean
Dim sTextBox As String
Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub LoadFund()
    Dim adoConn As New ADODB.Connection
    Dim rsFund As New ADODB.Recordset
    adoConn.Open getConnectionString
    rsFund.Open "SELECT  FundID,FundCode, FundName FROM FUND;", adoConn, adOpenStatic, adLockReadOnly
   Dim Data() As String
   Dim i, TotalRow, TotalCol As Integer, j As Integer

   TotalRow = rsFund.RecordCount
   TotalCol = rsFund.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(rsFund.Fields(j).Value), "", rsFund.Fields(j).Value)
       Next j
       rsFund.MoveNext
       If rsFund.EOF Then Exit For
   Next i
   cmbFund.Column() = Data()
    rsFund.Close
    adoConn.Close
    
End Sub
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

'        Case TypeOf ctl Is PictureBox
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            'Mouse wheel was not responding on picturebox
            'this problem fixed by anol 23 Mar 2016
            Case TypeOf ctl Is PictureBox
'                        If Not ctl Is picClient Then
'                            PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
'                        Else
                            bHandled = False
'                        End If

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
Private Sub cboAutoRef_Click()
'   If cboAutoRef.Value = "F" Then
'      fraReference.Enabled = True
'      txtCheqNo.text = lblRef.Caption
'   Else
'      fraReference.Enabled = False
'      lblRef.Caption = txtCheqNo.text
'      txtCheqNo.text = ""
'   End If
End Sub

Private Sub cboAutoRef_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   If KeyCode = 9 Then Exit Sub
   If KeyCode < 37 Or KeyCode > 40 Then KeyCode = 0
End Sub

'Private Sub cboClient_GotFocus()
'    SelTxtInCtrl cboClient
'End Sub

Private Sub cboClient_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
'   If KeyCode = 9 Then Exit Sub
'   If KeyCode < 37 Or KeyCode > 40 Then KeyCode = 0
End Sub

Private Sub chkMultiple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtDate.Enabled Then
             txtDate.SetFocus
        Else
            OKButton.SetFocus
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnLoadForm Me
End Sub

'Private Sub cboClient_KeyPress(KeyAscii As MSForms.ReturnInteger)
'   If KeyAscii = 13 Then
'        cboProperty.SetFocus
'     End If
'End Sub

Private Sub txtClient_Change()
'    'added by anol 23 Feb 2016
'   If txtClient.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   'LoadProperties adoConn, cboProperty, txtClient.tag
'   LoadBank adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
End Sub

'Private Sub cboProperty_GotFocus()
'    SelTxtInCtrl cboProperty
'End Sub

Private Sub cboProperty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'   If KeyCode = 9 Then Exit Sub
'   If KeyCode < 37 Or KeyCode > 40 Then KeyCode = 0
End Sub

Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmbBankAc.SetFocus
    End If
End Sub

Private Sub chkMultiple_Click()
   If chkMultiple.Value = 1 Then
      lblDate.Caption = txtDate.text
      lblPostingDate.ToolTipText = txtDate.text

      txtCheqNo.text = ""
      txtDate.text = ""
      txtCheqNo.Locked = True
      fraMultiple.Enabled = False
   Else
      txtCheqNo.Locked = False
      fraMultiple.Enabled = True
      txtDate.text = lblDate.Caption
   End If
End Sub

Private Sub cmbBankAc_Change()
       If cmbBankAc.ListIndex <> -1 Then
            Label13(7).Caption = " " & cmbBankAc.Column(0)
       Else
            Label13(7).Caption = ""
       End If
End Sub



Private Sub cmbBankAc_GotFocus()
   SelTxtInCtrl cmbBankAc
End Sub


Public Sub Testing_Command()
   OKButton_Click
End Sub

Private Sub cmbBankAc_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        If txtDate.Enabled = True Then
            txtDate.SetFocus
        Else
            OKButton.SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
    picClient.Visible = True
    LoadflxClient
End Sub

Private Sub cmdClient_Click()
    sTextBox = "1"
    Frame2.Enabled = False
    picClient.Left = 970
    picClient.Top = 225
    txtSearchClientID.text = ""
    txtSearchClientName.text = ""
    LoadflxClient
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxProperty()
    flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 4275
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

  
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)

   lblClientID.Caption = "Property ID"
   lblClientName.Caption = "Property Name"
  
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   Dim rRow As Integer
   Dim szSQL As String
   Dim rstRec As New ADODB.Recordset
    rRow = 1
   
    
    If txtClient.text <> "ALL" Then
        szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "WHERE ClientID = '" & txtClient.Tag & "' " & _
            "ORDER BY PropertyID;"
    Else
        szSQL = "SELECT PropertyID, PropertyName " & _
            "FROM Property " & _
            "ORDER BY PropertyID;"
    End If
    
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   flxClient.TextMatrix(1, 0) = "ALL"
   flxClient.TextMatrix(1, 1) = "All Properties"
   flxClient.AddItem ""
   rRow = 2
   While Not rstRec.EOF
       flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
       flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
       flxClient.RowHeight(rRow) = 280
       rstRec.MoveNext
       If Not rstRec.EOF Then flxClient.AddItem ""
       rRow = rRow + 1
    Wend

 

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdPicCLose_Click()
    Frame2.Enabled = True
    picClient.Visible = False
End Sub

Private Sub flxClient_Click()
    Frame2.Enabled = True
    If sTextBox = "1" Then
        txtClient.text = flxClient.TextMatrix(flxClient.row, 1)
        txtClient.Tag = flxClient.TextMatrix(flxClient.row, 0)
            'added by anol 23 Feb 2016
       If txtClient.text = "" Then Exit Sub
       Dim adoConn As New ADODB.Connection
       adoConn.Open getConnectionString
       LoadBank adoConn
       adoConn.Close
       Set adoConn = Nothing
       cmdProperty.SetFocus
       txtProperty.text = "All Properties"
       txtProperty.Tag = "ALL"
       Dim strTemp As String
       txtClient.ForeColor = vbBlack
       strTemp = isControlAccountSet(txtClient.Tag)
       If Len(strTemp) > 0 Then
            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClient.text & _
            vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
            strTemp = ""
            picClient.Visible = False
            txtClient.ForeColor = vbRed
            Exit Sub
       End If
    ElseIf sTextBox = "2" Then
        txtProperty.text = flxClient.TextMatrix(flxClient.row, 1)
        txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 0)
        cmbBankAc.SetFocus
    End If
    
    picClient.Visible = False
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
            Frame2.Enabled = True
            picClient.Visible = False
            flxClient.Clear
            flxClient.Cols = 2
            flxClient.Rows = 2
            cmdClient.SetFocus
    End If
    If KeyAscii = 13 Then
        Frame2.Enabled = True
        If sTextBox = "1" Then
           txtClient.text = flxClient.TextMatrix(flxClient.row, 1)
           txtClient.Tag = flxClient.TextMatrix(flxClient.row, 0)
               'added by anol 23 Feb 2016
            If txtClient.text = "" Then Exit Sub
            Dim adoConn As New ADODB.Connection
            adoConn.Open getConnectionString
            LoadBank adoConn
            adoConn.Close
            Set adoConn = Nothing
            txtProperty.text = "All Properties"
           txtProperty.Tag = "ALL"
           cmdProperty.SetFocus
        ElseIf sTextBox = "2" Then
           txtProperty.text = flxClient.TextMatrix(flxClient.row, 1)
           txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 0)
           cmbBankAc.SetFocus
        End If
    End If
    picClient.Visible = False
End Sub

Private Sub txtClient_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdClient.SetFocus
    End If
End Sub

Private Sub txtClient_KeyPress(KeyAscii As MSForms.ReturnInteger)
        If KeyAscii = 13 Then
                cmdClient.SetFocus
        End If
End Sub



'Private Sub txtProperty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 13 Then
'        cmdproperty.SetFocus
'    End If
'End Sub

Private Sub txtProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmdProperty.SetFocus
    End If
End Sub

Private Sub txtSearchClientID_Change()
    'Updated by anol 10 Dec 2015
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClient.Rows - 1 To 1 Step -1
      flxClient.RowHeight(i) = 240
      
      If InStr(1, UCase(flxClient.TextMatrix(i, 0)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
            flxClient.RowHeight(i) = 0
      End If
      If flxClient.RowHeight(i) = 240 Then
            flxClient.row = i
      End If
   Next i
End Sub
Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
         txtSearchClientName.SetFocus
    End If
    If KeyAscii = 27 Then
          Frame2.Enabled = True
          flxClient.Clear
          flxClient.Cols = 2
          flxClient.Rows = 2
          picClient.Visible = False
          cmdClient.SetFocus
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
      If InStr(1, UCase(flxClient.TextMatrix(i, 1)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
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
'Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'     If KeyCode = vbKeyDown Then
'            flxClient.SetFocus
'     End If
'End Sub
Private Sub cmdproperty_Click()
    sTextBox = "2"
    Frame2.Enabled = False
    picClient.Left = 970
    picClient.Top = 325
    txtSearchClientID.text = ""
    txtSearchClientName.text = ""
    LoadflxProperty
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub

Private Sub Form_Activate()
   bLoaded = True
   Dim strTemp As String
   txtClient.ForeColor = vbBlack
   If txtClient.Tag = "" Then Exit Sub
   strTemp = isControlAccountSet(txtClient.Tag)
   If Len(strTemp) > 0 Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & strTemp & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        strTemp = ""
        picClient.Visible = False
        txtClient.ForeColor = vbRed
        Exit Sub
    End If
End Sub

Public Sub TestingCommand()
   OKButton_Click
End Sub

Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 1500
   flxClient.ColWidth(1) = 4275
   flxClient.ColWidth(2) = 0
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

  
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   
    lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"

   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
  
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.TextMatrix(rRow, 0) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(1).Value
               flxClient.TextMatrix(rRow, 2) = IIf(IsNull(rstRec.Fields.Item(2).Value), "", rstRec.Fields.Item(2).Value)
               flxClient.RowHeight(rRow) = 280
               rstRec.MoveNext
               If Not rstRec.EOF Then flxClient.AddItem ""
               rRow = rRow + 1
           Wend
   

   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection
   Dim Rst1    As New ADODB.Recordset
   Dim szSQL   As String
'   'cascade form function created by anol 2019 -06-17
'    Dim iLeft As Integer
'    Dim iTop As Integer
'    Call BuildFormlist(Me.Name, iTop, iLeft)
'    frmBPPreForm.Top = iTop
'    frmBPPreForm.Left = iLeft
'    'Cascade Form End
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
   Me.Width = 8175
   Me.Height = 5475
   On Error GoTo Err
'   Me.BackColor = MODULEBACKCOLOR
'   fraMultiple.BackColor = MODULEBACKCOLOR
'   fraReference.BackColor = MODULEBACKCOLOR
'   Frame1.BackColor = MODULEBACKCOLOR
'   chkMultiple.BackColor = MODULEBACKCOLOR
'    LoadFund
'    If cmbFund.ListCount > 0 Then
'        cmbFund.ListIndex = 0
'    End If
   bLoaded = False

   adoConn.Open getConnectionString
   'Resolved by BOSL
  ' Issue 518
  'Modified by anol 15 Dec 2014
''
'''   Add new column PostingDate on 24/05/14 tblBatchReceipt
'''###############################################################################################################
''   On Error GoTo CHANGE_ADD_PostingDate_tblBatchReceipt
''
''   Rst1.Open "SELECT PostingDate FROM tblBatchReceipt;", adoConn, adOpenStatic, adLockReadOnly
''   Rst1.Close
'''Samrat 07/08/2014
''   Rst1.Open "SELECT * FROM tblBatchReceipt WHERE ISNULL(PostingDate) OR PostingDate ='';", adoConn, adOpenStatic, adLockReadOnly
''   If Not Rst1.EOF Then
''      adoConn.Execute "UPDATE tblBatchReceipt SET PostingDate = FORMAT(BRDate, 'DD MMMM YYYY')"
''   End If
''   Rst1.Close
''   Set Rst1 = Nothing
''
''   GoTo WorkContinue
''CHANGE_ADD_PostingDate_tblBatchReceipt:
''   adoConn.Execute "ALTER TABLE tblBatchReceipt ADD COLUMN PostingDate TEXT(20);"
''   adoConn.Execute "UPDATE tblBatchReceipt SET PostingDate = FORMAT(BRDate, 'DD MMMM YYYY')"
''   'Resolved by BOSL
''   'Issue number 0000439
''   'Program Crashes After Batch Receipts is Run After Upgrade
''   'Modified my Anol 05-08-2014
''
''   If Rst1.State = adStateOpen Then
''        Rst1.Close
''        Set Rst1 = Nothing
''   End If
''
''
''WorkContinue:

   'PrepareList adoConn, cboClient, cboProperty
'    szSQL = "SELECT  CLIENTID,ClientName " & _
'           "FROM  Client " & _
'           "ORDER BY CLIENTID;"
'
''Debug.Print szSQL
'   Rst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount = 0 Then
'          MsgBox "You must create a Client before entering receipt.", vbCritical + vbOKOnly, "Batch Payment"
'          Rst1.Close
'          Set Rst1 = Nothing
'          adoConn.Close
'          Set adoConn = Nothing
'
'          Exit Sub
'   Else
'         txtClient.Tag = Rst1.Fields("CLIENTID").Value
'         txtClient.text = Rst1.Fields("ClientName").Value
'         txtProperty.text = "All Propertiess"
'         txtProperty.Tag = "ALL"
'         Rst1.Close
'         LoadBank adoConn
'   End If
   cmbBankAc.AddItem "", 0
   cmbBankAc.ListIndex = 0

'  Set today's date in the date text box
   txtDate.text = Format(Date, "dd/mm/yyyy")

   lblMessage.Visible = IIf(LoadPreviousSelection(adoConn), True, False)
  ' Me.Height = IIf(lblMessage.Visible, 3975, 3705)

   adoConn.Close
   Set adoConn = Nothing

   lblRef.Caption = ""
   LoadAutoRef
   Call WheelHook(Me.hWnd)
   Exit Sub
Err:
    MsgBox Err.description
End Sub
Private Sub loadfirstclient(adoConn As ADODB.Connection)
    Dim szSQL As String
    Dim Rst1 As New ADODB.Recordset
    szSQL = "SELECT  CLIENTID,ClientName " & _
           "FROM  Client " & _
           "ORDER BY CLIENTID;"
   Rst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount = 0 Then
          MsgBox "You must create a Client before entering a receipt.", vbCritical + vbOKOnly, "Batch Payment"
          Rst1.Close
          Set Rst1 = Nothing
'          adoConn.Close
'          Set adoConn = Nothing
          
          Exit Sub
   Else
         txtClient.Tag = Rst1.Fields("CLIENTID").Value
         txtClient.text = Rst1.Fields("ClientName").Value
         txtProperty.text = "All Properties"
         txtProperty.Tag = "ALL"
         Rst1.Close
         LoadBank adoConn
   End If
End Sub
Private Function LoadPreviousSelection(adoConn As ADODB.Connection) As Boolean
'  iFound: will help to find any saved selection is deleted/removed from system.
   Dim szSQL As String
   Dim iFound As Integer
   Dim adoRst As New ADODB.Recordset

   adoRst.Open "SELECT * FROM tblBatchReceipt WHERE Generated = FALSE;", adoConn, adOpenStatic, adLockReadOnly
   LoadPreviousSelection = IIf(adoRst.EOF, False, True)
   adoRst.Close
   Set adoRst = Nothing

'  Remember choice
   szChoice = GetSetting("PropertyManagement", "ChoosedOption", "BR-c" & CStr(SCID))
   szaChoice = Split(szChoice, "#")

   iFound = True

'  iFound variable is not in use completely for the time beign. its need to impliment.
   If UBound(szaChoice) > 0 Then
     
'rem out by anol 23 Feb 2016
'      iFound = DropDownListPoint(cboClient, szaChoice(1))
'      If iFound >= 0 Then cboClient.ListIndex = iFound
'
'      If UBound(szaChoice) >= 2 Then
'         iFound = DropDownListPoint(cboProperty, szaChoice(2))
'         If iFound Then cboProperty.ListIndex = iFound
'
'         If UBound(szaChoice) >= 3 Then
'            iFound = DropDownListPoint(cmbBankAc, szaChoice(3))
'            If iFound Then cmbBankAc.Value = iFound
'         End If
'      End If
      
'      adoRst.Open "SELECT  CLIENTID,ClientName " & _
'           "FROM  Client where ClientName='" & txtClient.text & "' " & _
'           "ORDER BY CLIENTID;", adoConn, adOpenStatic, adLockReadOnly
'       If adoRst.EOF Then
'            adoRst.Close
'            Call loadfirstClient(adoConn)
'            Exit Function
'       End If
     
      'added by anol 16 Feb 2016
       If UBound(szaChoice) >= 6 Then
            adoRst.Open "SELECT  CLIENTID,ClientName " & _
           "FROM  Client where ClientName='" & szaChoice(6) & "' " & _
           "ORDER BY CLIENTID;", adoConn, adOpenStatic, adLockReadOnly
            If adoRst.EOF Then
                 adoRst.Close
                 Call loadfirstclient(adoConn)
                 Exit Function
            End If
            adoRst.Close
            txtClient.text = szaChoice(6)
            txtProperty.text = szaChoice(7)
       End If
        If szaChoice(0) = "Y" Then chkMultiple.Value = 1
        If szaChoice(0) = "N" Then
                chkMultiple.Value = 0
                If UBound(szaChoice) > 3 Then
                   txtDate.text = szaChoice(4)
                   lblPostingDate.ToolTipText = szaChoice(5)
                End If
        End If
         txtClient.Tag = szaChoice(1)
         LoadBank adoConn
        If UBound(szaChoice) >= 2 Then
'         iFound = DropDownListPoint(cboProperty, szaChoice(2))
'         If iFound Then cboProperty.ListIndex = iFound
         txtProperty.Tag = szaChoice(2)
         If UBound(szaChoice) >= 3 Then
            If optBR_Bank.Value Then
               szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.NominalCode, NL.Name " & _
                       "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                       "WHERE  " & _
                           "C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                           "C.ClientID = '" & txtClient.Tag & "' AND DEFAULT_AC =true order by CB.NominalCode ;"
            Else
               szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.NominalCode, NL.Name " & _
                       "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL " & _
                       "WHERE C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                             "C.ClientID = '" & txtClient.Tag & "' AND DEFAULT_AC =true  order by CB.NominalCode;"
            End If
            adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
            If adoRst.EOF Then
                 adoRst.Close
                ' MsgBox "Please set a default Client Bank Account for:" & txtClient.Tag & "", vbInformation, "Warning"
                 'These two line is for bank selection
                iFound = DropDownListPoint(cmbBankAc, szaChoice(3))
                If iFound >= 0 Then cmbBankAc.Value = iFound
            Else
                'These two line is for bank selection
                iFound = DropDownListPoint(cmbBankAc, adoRst("NominalCode").Value)
                adoRst.Close
                If iFound >= 0 Then cmbBankAc.Value = iFound
           End If
'            iFound = DropDownListPoint(cmbBankAc, szaChoice(3))
'            If iFound >= 0 Then cmbBankAc.Value = iFound
         End If
      End If
   End If
End Function

Private Sub PrepareList(adoConn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()

   adoRst.Close
'*************************************** PROPERTY ******************************************
   If cboC.text <> "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboC.Column(0) & "' " & _
              "ORDER BY PropertyID;"
   Else
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   End If
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   'issue 571 by anol 22 July 2015
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadBank(adoConn As ADODB.Connection)
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   Dim iDefaultBankAC As Integer
'Add bank receipt form: add bank code below the bank account
'added by anol 13 Mar 2015
   If txtClient.text = "" Then
      szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.NominalCode, NL.Name,CB.DEFAULT_AC " & _
              "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
              "WHERE C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode order by CB.NominalCode;"
   Else
      If optBR_Bank.Value Then
         szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.NominalCode, NL.Name,CB.DEFAULT_AC " & _
                 "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                 "WHERE  " & _
                     "C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                     "C.ClientID = '" & txtClient.Tag & "' order by CB.NominalCode ;"
      Else
         szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.NominalCode, NL.Name,CB.DEFAULT_AC " & _
                 "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL " & _
                 "WHERE C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                       "C.ClientID = '" & txtClient.Tag & "' order by CB.NominalCode;"
      End If
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   iDefaultBankAC = -1
   If adoRst.EOF Then
   'this message is correct but I have rem it because it creates a time legacy
     ' ShowMsgInTaskBar "No Bank accounts have been created for this Client. Please add a Bank account for this Client"
   Else
      ReDim szaData(4, adoRst.RecordCount - 1) As String

      While Not adoRst.EOF
       
      'old copy
'       szaData(0, iRec) = adoRst.Fields.Item("MY_ID").Value
'         szaData(1, iRec) = adoRst.Fields.Item("Bank_AC_Name").Value
'         szaData(2, iRec) = adoRst.Fields.Item("ClientName").Value
'         szaData(3, iRec) = adoRst.Fields.Item("NominalCode").Value


'Add bank receipt form: add bank code below the bank account
'added by anol 13 Mar 2015
         szaData(0, iRec) = adoRst.Fields.Item("NominalCode").Value
         szaData(1, iRec) = adoRst.Fields.Item("Name").Value
         szaData(2, iRec) = adoRst.Fields.Item("MY_ID").Value
         If CBool(adoRst.Fields.Item("DEFAULT_AC").Value) = True Then
            iDefaultBankAC = iRec
         End If
         
         iRec = iRec + 1
         adoRst.MoveNext
      Wend
   End If

   cmbBankAc.Clear
   cmbBankAc.Column() = szaData()
   If iDefaultBankAC < 0 Then
         MsgBox "Please set a default Client Bank Account for: " & txtClient.Tag & "", vbInformation, "Warning"
         cmbBankAc.ListIndex = 0
   Else
         cmbBankAc.ListIndex = iDefaultBankAC
   End If
  

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
   If IsNull(txtClient.Tag) = True Then
         ShowMsgInTaskBar "Please select a client", "Y"
         txtClient.SetFocus
         Exit Sub
   End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, txtClient.Tag
End Sub

Private Sub OKButton_Click()
On Error GoTo Err:
    If txtClient.text = "" Then
      MsgBox "Please select the client.", vbCritical + vbOKOnly, "Batch Process"
      cmdClient.SetFocus
      Exit Sub
   End If
    If txtClient.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClient.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
'added by anol on 22 July 2015
'issue 571
  If txtProperty.text = "" Then
      MsgBox "Please select a valid property to proceed.", vbCritical + vbOKOnly, "Select a Property"
      cmdProperty.SetFocus
      Exit Sub
   End If
    If txtProperty.text = "" Then
      MsgBox "Please select a valid property to proceed.", vbCritical + vbOKOnly, "Select a Property"
      cmdProperty.SetFocus
      Exit Sub
   End If
   
   If cmbBankAc.ListIndex = "-1" Then
      MsgBox "Please select the bank account.", vbCritical + vbOKOnly, "Batch Process"
      cmbBankAc.SetFocus
      Exit Sub
   End If
   'end of modification
   If txtDate.text = "" And chkMultiple.Value = 0 Then
      MsgBox "Please enter the date.", vbExclamation + vbOKOnly, "Batch Process"
      txtDate.SetFocus
      Exit Sub
   End If

   Dim szSQL   As String
   Dim mbr     As VbMsgBoxResult
   Dim adoRst  As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
    
   adoConn.Open getConnectionString
   If chkMultiple.Value = 0 Then
        If IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 0 Then
                MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
                adoConn.Close
                Set adoConn = Nothing
                Exit Sub
             ElseIf IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 9 Then
                MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
                adoConn.Close
                Set adoConn = Nothing
                Exit Sub
        End If
   End If
        
   szSQL = "SELECT B.*, T.TransactionID, T.RptAmt, T.RptDt, T.Ref " & _
           "FROM tblBatchReceipt AS B, tblBtRptTran AS T " & _
           "WHERE B.BR = T.BR AND B.Generated = FALSE AND " & _
                 "T.RptAmt > 0;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      mbr = MsgBox("There are saved receipts in this batch. Do you wish to process them?", vbQuestion + vbYesNoCancel)

      If mbr = vbYes Then
'         If IsNull(adoRst.Fields.Item("RptDt").Value) Then
'            chkMultiple.Value = 0
'         Else
'            chkMultiple.Value = 1
'         End If
      End If
      If mbr = vbNo Then ClearSavedBR adoConn             'Clear Saved Batch Receipt

      If mbr = vbCancel Then
         adoRst.Close
         Set adoRst = Nothing
         adoConn.Close
         Set adoConn = Nothing

         Exit Sub
      End If
   End If

   adoRst.Close
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   Me.Hide
   Load frmBatchRpt

   frmBatchRpt.lblClient.Caption = txtClient.text
   frmBatchRpt.lblClient.ToolTipText = txtClient.text
   frmBatchRpt.lblProperty.Caption = txtProperty.text
   frmBatchRpt.lblProperty.ToolTipText = txtProperty.Tag
   frmBatchRpt.lblBank.Caption = cmbBankAc.text
   frmBatchRpt.lblBC.Caption = Label13(7).Caption
   frmBatchRpt.lblBank.ToolTipText = cmbBankAc.text
   frmBatchRpt.lblDate.Caption = txtDate.text

   'Issue 0000534. Added by Asif. Date: 11 Feb 2015
   If chkMultiple = 0 Then
      frmBatchRpt.grpUploadReceipts(0).Visible = False
   End If
   '''
   
   frmBatchRpt.FromLoad_DataGrid
   frmBatchRpt.FilterGrid
   frmBatchRpt.Show

   SaveCurrentSelection
   frmBatchRpt.flxSPayment.SetFocus
   frmBatchRpt.flxSPayment.row = 1
   
     frmBatchRpt.flxSPayment.col = 11
   Exit Sub
Err:
   MsgBox Err.description
End Sub

Private Sub SaveCurrentSelection()
   Dim szChoice As String
   Dim szaChoice() As String

   If chkMultiple.Value Then
      ReDim szaChoice(7) As String
      
      szaChoice(0) = "Y"
   Else
      ReDim szaChoice(7) As String
      
      szaChoice(0) = "N"
      szaChoice(4) = Format(txtDate.text, "dd/mm/yyyy")
      szaChoice(5) = Format(lblPostingDate.ToolTipText, "dd/mm/yyyy")
   End If
   
   szaChoice(1) = txtClient.Tag
   szaChoice(2) = txtProperty.Tag
   szaChoice(3) = cmbBankAc.Column(0)
   szaChoice(6) = txtClient.text
   szaChoice(7) = txtProperty.text

   szChoice = Join(szaChoice, "#")

   SaveSetting "PropertyManagement", "ChoosedOption", "BR-c" & CStr(SCID), szChoice
End Sub

'Private Sub cboClient_Click()
''
'   If cboClient.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'
'   adoConn.Open getConnectionString
'
'   LoadProperties adoConn, cboProperty, txtClient.Tag
'   LoadBank adoConn
'
'   adoConn.Close
'   Set adoConn = Nothing
'End Sub

Private Sub LoadProperties(adoConn As ADODB.Connection, cboP As Control, szClientID As String)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, j As Integer
   Dim i As Integer, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer

   On Error GoTo ErrorHandler

'***************************************  PROPERTY  ******************************************
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & szClientID & "' " & _
           "ORDER BY PropertyID;"
'   Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   cboP.Clear
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   'Issue 571 By anol 22  july 2015
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   
   cboP.Column() = Data()
   cboP.ListIndex = 0

NoRes:
   adoRst.Close
   Set adoRst = Nothing

   Exit Sub

ErrorHandler:
   MsgBox Err.description & "::" & Err.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub LoadAutoRef()
'   Dim i As Integer, Data(1, 2) As String
'
'   Data(0, 0) = "L"
'   Data(1, 0) = "Lessee ID"
'   Data(0, 1) = "U"
'   Data(1, 1) = "Unit ID"
'   Data(0, 2) = "F"
'   Data(1, 2) = ""
'
'   cboAutoRef.Column() = Data()
'   cboAutoRef.ListIndex = -1
End Sub

Private Sub optBr_Bank_Click()
   If optBR_Bank.Value Then
      txtClient.text = ""
      txtClient.Tag = ""
      txtCheqNo.text = ""
   End If
End Sub

Private Sub optBr_Cheque_Click()
   If optBR_Cheque.Value Then txtClient.text = ""
End Sub

Private Sub txtCheqNo_GotFocus()
   SelTxtInCtrl txtCheqNo
End Sub

Private Sub txtDate_Change()
   'Resolved by BOSL
   'issue 468
   'modified by Anol 03 Sep 2014
   TextBoxChangeDate txtDate
   lblPostingDate.ToolTipText = txtDate.text
End Sub

Private Sub txtDate_GotFocus()
   If txtDate.text = "dd/mm/yyyy" Then
      txtDate.text = ""
      Exit Sub
   End If
   If Len(txtDate.text) < 10 Then txtDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OKButton.SetFocus
    End If
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
    If txtDate.text <> "" Then
      TextBoxFormatDate txtDate
   Else
      txtDate.text = Format(Now, "dd/mm/yyyy")
   End If
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClient.text = "" Then
          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
          Exit Sub
    End If
    If frmMMain.IsRibbonVersion Then
        Dim adoConn As New ADODB.Connection
        Dim szSQL As String
        adoConn.Open getConnectionString
        If IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 0 Then
           MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 9 Then
           MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
    End If
End Sub
