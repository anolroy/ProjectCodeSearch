VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBPPreForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Payment"
   ClientHeight    =   6780
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   15675
   Icon            =   "frmBPPreForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   15675
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   8055
      ScaleHeight     =   4740
      ScaleWidth      =   6120
      TabIndex        =   16
      Top             =   135
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
         TabIndex        =   17
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   15
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   14
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
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   45
         TabIndex        =   13
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   20
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
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   6540
      Left            =   45
      TabIndex        =   22
      Top             =   -45
      Width           =   8700
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
         Left            =   6255
         TabIndex        =   7
         Top             =   3915
         Width           =   300
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         TabIndex        =   11
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3660
         TabIndex        =   12
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2250
         TabIndex        =   9
         Top             =   4635
         Width           =   1520
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Payment Method"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   2715
         TabIndex        =   24
         Top             =   270
         Width           =   3435
         Begin VB.OptionButton optBP_Cheque 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cheque"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   245
            Left            =   240
            TabIndex        =   0
            Top             =   220
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBP_BACS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "BACS"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   245
            Left            =   240
            TabIndex        =   1
            Top             =   510
            Width           =   735
         End
         Begin VB.OptionButton optBP_MULT 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multiple"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   245
            Left            =   240
            TabIndex        =   2
            Top             =   800
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCheqNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2250
         TabIndex        =   10
         Top             =   5010
         Width           =   1740
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Cheque Print Option"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   2715
         TabIndex        =   23
         Top             =   1665
         Width           =   3435
         Begin VB.OptionButton optSCPO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cheque with Remittance"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optSCPO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Remittance Only"
            BeginProperty Font 
               Name            =   "Myriad Web"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   280
            Width           =   1935
         End
      End
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
         Left            =   6270
         TabIndex        =   5
         Top             =   3150
         Width           =   300
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   2250
         TabIndex        =   37
         Top             =   3915
         Width           =   4005
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "7064;503"
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
         Index           =   11
         Left            =   765
         TabIndex        =   36
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label19 
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
         Index           =   12
         Left            =   765
         TabIndex        =   35
         Top             =   3915
         Width           =   645
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   780
         TabIndex        =   34
         Top             =   4635
         Width           =   375
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
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
         Index           =   0
         Left            =   765
         TabIndex        =   33
         Top             =   4260
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next Cheque No."
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
         Index           =   1
         Left            =   765
         TabIndex        =   32
         Top             =   5010
         Width           =   1215
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There are saved payments in this payment batch!"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   915
         TabIndex        =   31
         Top             =   5925
         Width           =   4440
      End
      Begin MSForms.Label lblPostingDate 
         Height          =   300
         Left            =   3765
         TabIndex        =   30
         Top             =   4635
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
      Begin MSForms.Label Label13 
         Height          =   315
         Index           =   7
         Left            =   2250
         TabIndex        =   29
         Top             =   4290
         Width           =   1080
         BackColor       =   16777215
         Size            =   "1905;556"
         BorderStyle     =   1
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbBankAc 
         Height          =   345
         Left            =   3375
         TabIndex        =   8
         Top             =   4275
         Width           =   4200
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "7408;609"
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
      Begin VB.Label lblOverDraftAmount 
         BackColor       =   &H0080FFFF&
         Height          =   150
         Left            =   4875
         TabIndex        =   28
         Top             =   4860
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label isOverDratAllowed 
         BackColor       =   &H0080FFFF&
         Height          =   150
         Left            =   5640
         TabIndex        =   27
         Top             =   4860
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Category:"
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   26
         Top             =   3555
         Width           =   1455
      End
      Begin MSForms.ComboBox cmdACType 
         Height          =   285
         Left            =   2265
         TabIndex        =   6
         Top             =   3555
         Width           =   2085
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3678;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClient 
         Height          =   285
         Left            =   2265
         TabIndex        =   25
         Top             =   3150
         Width           =   4005
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "7064;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmBPPreForm"
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

Private Sub cmdproperty_Click()
    sTextBox = "2"
    Frame2.Enabled = False
    picClient.Left = 970
    picClient.Top = 870
    
    LoadflxProperty
    picClient.Visible = True
    txtSearchClientID.SetFocus
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
  
   Dim adoconn As New ADODB.Connection
   adoconn.Open getConnectionString
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
    
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   flxClient.TextMatrix(1, 0) = "ALL"
   flxClient.TextMatrix(1, 1) = "All Properties"
   flxClient.RowHeight(1) = 280
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
'Private Sub cboClient_Change()
'comment out by anol 12 Feb 2016
'    'resolved by BOSL
'    'issue 455
'    'Modified by anol 21 Aug 2014
'    If IsNull(cboClient) Then Exit Sub
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    LoadProperties adoConn, cboProperty, txtClient.Tag
'    If cboClient.text = "All Clients" Then Exit Sub
'    LoadBank adoConn
'    adoConn.Close
'    Set adoConn = Nothing
'End Sub

'Private Sub cboClient_GotFocus()
'    SelTxtInCtrl cboClient
'End Sub

'Private Sub txtClient_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cboProperty.SetFocus
'    End If
'End Sub

Private Sub txtClient_Change()
'    'added by anol 11 Nov 2015
'    If txtClient.text = "" Then Exit Sub
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    txtProperty.text = "All Properties"
'    txtProperty.Tag = "ALL"
''    LoadProperties adoConn, cboProperty, txtClient.Tag
''    If cboClient.text = "All Clients" Then Exit Sub
'    LoadBank adoConn
'    adoConn.Close
'    Set adoConn = Nothing
End Sub

'Private Sub cboProperty_GotFocus()
'    SelTxtInCtrl cboProperty
'End Sub

Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        cmbBankAc.SetFocus
    End If
End Sub

Private Sub cmbBankAc_Change()
    If cmbBankAc.ListIndex <> -1 Then
         Label13(7).Caption = " " & cmbBankAc.Column(0)
    Else
         Label13(7).Caption = ""
    End If

    'issue 547 bank overdraft implementation
    If cmbBankAc.ListIndex <> -1 Then
      isOverDratAllowed.Caption = cmbBankAc.Column(7)
      lblOverDraftAmount.Caption = cmbBankAc.Column(8)
    Else
      lblOverDraftAmount.Caption = ""
      isOverDratAllowed.Caption = ""
   End If
End Sub

Private Sub cmbBankAc_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 And txtDate.Enabled Then
        txtDate.SetFocus
    End If
End Sub

'Private Sub cmdACType_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cboProperty.SetFocus
'    End If
'End Sub

Private Sub cmdClient_Click()
    sTextBox = "1"
    Frame2.Enabled = False
    picClient.Left = 970
    picClient.Top = 870
    
    LoadflxClient
    picClient.Visible = True
    txtSearchClientID.SetFocus
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
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

   adoconn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
  
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
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing

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
         'added by anol 11 Nov 2015
        If txtClient.text = "" Then Exit Sub
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        txtProperty.text = "All Properties"
        txtProperty.Tag = "ALL"
        LoadBank adoconn
        adoconn.Close
        Set adoconn = Nothing
        cmdACType.SetFocus
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
        If txtDate.Enabled Then
            txtDate.SetFocus
        Else
            cmbBankAc.SetFocus
        End If
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
                 'added by anol 11 Nov 2015
                If txtClient.text = "" Then Exit Sub
                Dim adoconn As New ADODB.Connection
                adoconn.Open getConnectionString
                txtProperty.text = "All Properties"
                txtProperty.Tag = "ALL"
            '    LoadProperties adoConn, cboProperty, txtClient.Tag
            '    If cboClient.text = "All Clients" Then Exit Sub
                LoadBank adoconn
                adoconn.Close
                Set adoconn = Nothing
                cmdACType.SetFocus
            ElseIf sTextBox = "2" Then
                txtProperty.text = flxClient.TextMatrix(flxClient.row, 1)
                txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 0)
                If txtDate.Enabled Then
                    txtDate.SetFocus
                Else
                    cmbBankAc.SetFocus
                End If
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
                cmdACType.SetFocus
        End If
End Sub



Private Sub txtProperty_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        cmdProperty.SetFocus
    End If
End Sub

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
Private Sub txtSearchClientID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyDown Then
            flxClient.SetFocus
     End If
End Sub
Private Sub Form_Activate()
   bLoaded = True
End Sub

Private Sub Form_Load()
   On Error GoTo Err:
   Dim adoconn As New ADODB.Connection
   Dim Rst1    As New ADODB.Recordset
   Dim szSQL As String
'   'cascade form function created by anol 2019 -06-17
'    Dim iLeft As Integer
'    Dim iTop As Integer
'    Call BuildFormlist(Me.Name, iTop, iLeft)
'    frmBPPreForm.Top = iTop
'    frmBPPreForm.Left = iLeft
'    'Cascade Form End
'   frmMMain.Arrange vbCascade
'   Me.ZOrder 0
'   Me.Top = IsLoadedAndVisibleCount * 200
'   Me.Left = IsLoadedAndVisibleCount * 200
'   Me.Top = 0 ' (frmMMain.Height / 2) - (Me.Height / 2) - 600
'   Me.Left = 0 '(frmMMain.Width / 2) - (Me.Width / 2) - 600
'   Me.BackColor = MODULEBACKCOLOR
'   Frame1.BackColor = MODULEBACKCOLOR
'   optBP_Cheque.BackColor = MODULEBACKCOLOR
'   optBP_BACS.BackColor = MODULEBACKCOLOR
'   optBP_MULT.BackColor = MODULEBACKCOLOR
'   Frame3.BackColor = MODULEBACKCOLOR
'   optSCPO(0).BackColor = MODULEBACKCOLOR
'   optSCPO(1).BackColor = MODULEBACKCOLOR
   Me.Width = 8910
 'added by anol 31 Aug 2015
   'issue 571 note 1156
   cmdACType.AddItem "All Categories", 0
   cmdACType.AddItem "Supplier", 1
   cmdACType.AddItem "Client", 2
   cmdACType.AddItem "Managing Agent", 3
   cmdACType.AddItem "Landlord", 4
   cmdACType.ListIndex = 0
   'end of addition
   bLoaded = False
'added by anol 27 AMay 2015 was not showing date when loading this form
    If Len(txtDate.text) < 10 Then txtDate.text = Format(Date, "dd/mm/yyyy")
'End
   adoconn.Open getConnectionString

'Samrat 07/08/2014
   Rst1.Open "SELECT * FROM tblBatchPayment WHERE ISNULL(PostingDate);", adoconn, adOpenStatic, adLockReadOnly
   If Not Rst1.EOF Then
      adoconn.Execute "UPDATE tblBatchPayment SET PostingDate = FORMAT(BPDate, 'DD MMMM YYYY')  WHERE ISNULL(PostingDate)"
   End If
   Rst1.Close
   Set Rst1 = Nothing
   
  
'   szSQL = "SELECT  CLIENTID,ClientName " & _
'           "FROM  Client " & _
'           "ORDER BY CLIENTID;"
'
''Debug.Print szSQL
'   Rst1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If Rst1.RecordCount = 0 Then
'          MsgBox "You must create a Client before entering payment.", vbCritical + vbOKOnly, "Batch Payment"
'          Rst1.Close
'          Set Rst1 = Nothing
'          adoConn.Close
'          Set adoConn = Nothing
'
'          Exit Sub
'   Else
'         txtClient.Tag = Rst1.Fields("CLIENTID").Value
'         txtClient.text = Rst1.Fields("ClientName").Value
'         Rst1.Close
'         LoadBank adoConn
'
'   End If
   'End of addition
   'comment out by anol 12 Jan 2016
   ' PrepareList adoConn, cboClient, cboProperty
    'cboClient.ListIndex = 0
'  Set today's date in the date text box
   txtDate.text = Format(Date, "dd/mm/yyyy")
  
   lblMessage.Visible = IIf(LoadPreviousSelection(adoconn), True, False)
   'Me.Height = IIf(lblMessage.Visible, 5520, 5280)
   Me.Height = 6870 '7440
   adoconn.Close
   Set adoconn = Nothing

   If Not optBP_MULT.Value And lblPostingDate.ToolTipText = "" And txtDate.text <> "" Then lblPostingDate.ToolTipText = txtDate.text
   
    'issue 468
   'added by anol 27 Jan 2015
   lblPostingDate.ToolTipText = txtDate.text
   Call WheelHook(Me.hWnd)
   Exit Sub
Err:
    
End Sub
Private Sub loadfirstclient(adoconn As ADODB.Connection)
    Dim szSQL As String
    Dim Rst1 As New ADODB.Recordset
     szSQL = "SELECT  CLIENTID,ClientName " & _
           "FROM  Client " & _
           "ORDER BY CLIENTID;"

'Debug.Print szSQL
   Rst1.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If Rst1.RecordCount = 0 Then
          MsgBox "You must create a Client before entering payment.", vbCritical + vbOKOnly, "Batch Payment"
          Rst1.Close
          Set Rst1 = Nothing
          adoconn.Close
          Set adoconn = Nothing
          
          Exit Sub
   Else
         txtClient.Tag = Rst1.Fields("CLIENTID").Value
         txtClient.text = Rst1.Fields("ClientName").Value
         Rst1.Close
         LoadBank adoconn
                 
   End If
End Sub
Private Function LoadPreviousSelection(adoconn As ADODB.Connection) As Boolean
'  iFound: will help to find any saved selection is deleted/removed from system.
   Dim szSQL As String
   Dim iFound As Integer
   Dim adoRst As New ADODB.Recordset

   adoRst.Open "SELECT * FROM tblBatchPayment WHERE Generated = FALSE;", adoconn, adOpenStatic, adLockReadOnly
   LoadPreviousSelection = IIf(adoRst.EOF, False, True)
   adoRst.Close

'  Remember choice
   szChoice = GetSetting("PropertyManagement", "ChoosedOption", "BP-c" & CStr(SCID))
   szaChoice = Split(szChoice, "#")

   iFound = True

'  iFound variable is not in use completely for the time beign. its need to impliment.
'
'        adoRst.Open "SELECT  CLIENTID,ClientName " & _
'           "FROM  Client where CLIENTID='" & txtClient.Tag & "' " & _
'           "ORDER BY CLIENTID;", adoConn, adOpenStatic, adLockReadOnly
'       If adoRst.EOF Then
'            adoRst.Close
'            Exit Function
'       End If
        
       
   If UBound(szaChoice) > 0 Then
      If szaChoice(0) = "C" Then optBP_Cheque.Value = True
      If szaChoice(0) = "B" Then
         optBP_BACS.Value = True
         optSCPO(0).Value = True
      End If
      If szaChoice(0) = "M" Then optBP_MULT.Value = True
       'added by anol 09 Aug 2016
      If UBound(szaChoice) >= 6 Then
            adoRst.Open "SELECT  CLIENTID,ClientName " & _
           "FROM  Client where ClientName='" & szaChoice(6) & "' " & _
           "ORDER BY CLIENTID;", adoconn, adOpenStatic, adLockReadOnly
            If adoRst.EOF Then
                 adoRst.Close
                 Call loadfirstclient(adoconn)
                 Exit Function
            End If
            adoRst.Close
            txtClient.text = szaChoice(6)
            txtProperty.text = szaChoice(7)
       End If
'rem by anol
'      iFound = DropDownListPoint(cboClient, szaChoice(1))
'      If iFound >= 0 Then cboClient.ListIndex = iFound
      txtClient.Tag = szaChoice(1)
      LoadBank adoconn
      If UBound(szaChoice) >= 2 Then
'         iFound = DropDownListPoint(cboProperty, szaChoice(2))
'         If iFound Then cboProperty.ListIndex = iFound
         txtProperty.Tag = szaChoice(2)
         
         If UBound(szaChoice) >= 3 Then
         
            If optBP_BACS.Value Then
                szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef, NL.Name, CB.NominalCode, AllowOverDraft, OverDraftLimit, CLIENT_ID  " & _
                 "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                 "WHERE CB.FileLoc <> '' AND " & _
                     "C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                     "C.ClientID = '" & txtClient.Tag & "' AND DEFAULT_AC =true order by CB.NominalCode ;"
            Else
               szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef, NL.Name, CB.NominalCode, AllowOverDraft, OverDraftLimit, CLIENT_ID  " & _
                       "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                       "WHERE C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                             "C.ClientID = '" & txtClient.Tag & "' AND DEFAULT_AC =true order by CB.NominalCode ;"
            End If
            adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            If adoRst.EOF Then
                 adoRst.Close
                 'MsgBox "Please set a default Client Bank Account for: " & txtClient.Tag & "", vbInformation, "Warning"
                 'These two line is for bank selection
                iFound = DropDownListPoint(cmbBankAc, szaChoice(3))
                If iFound >= 0 Then cmbBankAc.Value = iFound
            Else
                'These two line is for bank selection
                iFound = DropDownListPoint(cmbBankAc, adoRst("NominalCode").Value)
                adoRst.Close
                If iFound >= 0 Then cmbBankAc.Value = iFound
           End If

            If UBound(szaChoice) >= 4 Then
               If szaChoice(4) = "RO" Then optSCPO(0).Value = True
               'Modified by anol 27 May 2015 On cheque with remittance was not selected correctly on multiple
               If szaChoice(4) = "CR" Then optSCPO(1).Value = IIf(optBP_BACS.Value = True, False, True) And optBP_MULT.Value = False
            End If
         End If
         
         
      End If
'      'added by anol 16 Feb 2016
'       If UBound(szaChoice) >= 6 Then
'            txtClient.text = szaChoice(6)
'            txtProperty.text = szaChoice(7)
'       End If
       
      If Not optBP_MULT.Value And UBound(szaChoice) > 4 Then lblPostingDate.ToolTipText = szaChoice(5)
   End If
End Function

Private Sub PrepareList(adoconn As ADODB.Connection, cboC As Control, cboP As Control)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

''Resolved by BOSL
''issue 571
''Modified by Anol 22 July 2015
   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 0 To TotalRow 'end of modification
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboC.Column() = Data()
   adoRst.Close
'*************************************** PROPERTY ******************************************
    'Resolved by BOSL
    'issue 455
    'Modified by Anol 21 Aug 2014
   If cboC.text = "All Clients" Or Trim(cboC.text) = "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
        szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & cboC.Column(0) & "' " & _
              "ORDER BY PropertyID;"
      
   End If
   'end of modification
'   Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
  
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

Private Sub LoadBank(adoconn As ADODB.Connection)
   'On Error GoTo Error_Handler
   Dim iDefaultBankAC As Integer
   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String
   'Add bank receipt form: add bank code below the bank account
   'added by anol 13 Mar 2015
   If txtClient.text = "" Then 'unreachable code
      szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef, NL.Name, CB.NominalCode, AllowOverDraft, OverDraftLimit, CLIENT_ID,DEFAULT_AC  " & _
              "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
              "WHERE CB.FileLoc <> '' AND C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode order by CB.NominalCode ;"
   Else
      If optBP_BACS.Value Then
         szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef, NL.Name, CB.NominalCode, AllowOverDraft, OverDraftLimit, CLIENT_ID,DEFAULT_AC   " & _
                 "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                 "WHERE CB.FileLoc <> '' AND " & _
                     "C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                     "C.ClientID = '" & txtClient.Tag & "' order by CB.NominalCode ;"
      Else
         szSQL = "SELECT CB.MY_ID, CB.Bank_AC_Name, C.ClientName, CB.BANK_AC_NUM, CB.BANK_SC, CB.BacsRef, NL.Name, CB.NominalCode, AllowOverDraft, OverDraftLimit, CLIENT_ID,DEFAULT_AC   " & _
                 "FROM tlbClientBanks AS CB, Client AS C, NominalLedger as NL  " & _
                 "WHERE C.ClientID = CB.CLIENT_ID AND NL.ClientID = CB.CLIENT_ID AND NL.Code =CB.NominalCode AND " & _
                       "C.ClientID = '" & txtClient.Tag & "' order by CB.NominalCode ;"
      End If
   End If
'Debug.Print szSQL
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   iDefaultBankAC = -1
   If adoRst.EOF Then
   'If optBP_BACS.Value Then clause added by anol 30 Aug 2016
      If optBP_BACS.Value Then
      'this message is correct but I have rem it because it  creates a time legacy
           ' ShowMsgInTaskBar "No bank account has been found with Electronic Banking information."
      Else
             If txtClient.text <> "" Then ShowMsgInTaskBar "No bank account has been created for this client."
      End If
'      cmbBankAc.Locked = True
'      cmbBankAc.ListIndex = -1
      cmbBankAc.Clear
   Else
      ReDim szaData(9, adoRst.RecordCount - 1) As String

      With adoRst.Fields
         While Not adoRst.EOF
         'Add bank receipt form: add bank code below the bank account
         'added by anol 13 Mar 2015
         'Old copy
'            szaData(0, iRec) = .Item("MY_ID").Value
'            szaData(1, iRec) = .Item("Bank_AC_Name").Value
'            szaData(2, iRec) = .Item("ClientName").Value
'            szaData(3, iRec) = .Item("BANK_AC_NUM").Value
'            szaData(4, iRec) = .Item("BANK_SC").Value
'            szaData(5, iRec) = IIf(IsNull(.Item("BacsRef").Value), "", .Item("BacsRef").Value)
         szaData(0, iRec) = adoRst.Fields.Item("NominalCode").Value
         szaData(1, iRec) = adoRst.Fields.Item("Name").Value
         szaData(2, iRec) = .Item("MY_ID").Value
         szaData(3, iRec) = .Item("BANK_AC_NUM").Value
         szaData(4, iRec) = .Item("BANK_SC").Value
         szaData(5, iRec) = IIf(IsNull(.Item("BacsRef").Value), "", .Item("BacsRef").Value)
         szaData(6, iRec) = .Item("Bank_AC_Name").Value
         szaData(7, iRec) = IIf(IsNull(.Item("AllowOverDraft").Value), "", .Item("AllowOverDraft").Value)
         szaData(8, iRec) = IIf(IsNull(.Item("OverDraftLimit").Value), "", .Item("OverDraftLimit").Value)
         szaData(9, iRec) = .Item("CLIENT_ID").Value
            If CBool(adoRst.Fields.Item("DEFAULT_AC").Value) = True Then
               iDefaultBankAC = iRec
            End If
            iRec = iRec + 1
            adoRst.MoveNext
         Wend
      End With
      cmbBankAc.Clear
      cmbBankAc.Column() = szaData()
        If iDefaultBankAC < 0 Then
            MsgBox "Please set a default Client Bank Account for: " & txtClient.Tag & "", vbInformation, "Warning"
            cmbBankAc.ListIndex = 0
        Else
             cmbBankAc.ListIndex = iDefaultBankAC
        End If
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Sub

Public Sub TestingCommand()
   OKButton_Click
   frmBatchPayment.Testing_Method
End Sub

Private Sub lblPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    'Resolved by BOSL
    'Issue 468
    'Modified by Anol 03 Sep 2014
    If txtClient.text = "" Then
          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
          Exit Sub
    End If
'    If frmMMain.IsRibbonVersion Then
'        Dim adoConn As New ADODB.Connection
'        Dim szSQL As String
'        adoConn.Open getConnectionString
'        If IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 0 Then
'           ShowMsgInTaskBar "The issue date cannot fall within a closed financial period", "Y", "N"
'           adoConn.Close
'           Set adoConn = Nothing
'           Exit Sub
'        ElseIf IsPeriodStatus(txtDate.text, txtClient.Tag, adoConn) = 9 Then
'           ShowMsgInTaskBar "The issue date does not fall in any existing financial period", "Y", "N"
'           adoConn.Close
'           Set adoConn = Nothing
'           Exit Sub
'        End If
'    End If
   DispayCalendar Me, lblPostingDate.ToolTipText, txtDate.text, txtClient.Tag
End Sub

Private Sub OKButton_Click()
    If txtClient.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClient.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
    If txtClient.text = "" Then
      MsgBox "Please select the client.", vbCritical + vbOKOnly, "Batch Process"
      cmdClient.SetFocus
      Exit Sub
   End If
   If txtProperty.text = "" Then
      MsgBox "Please select the property.", vbCritical + vbOKOnly, "Batch Process"
      cmdProperty.SetFocus
      Exit Sub
   End If
   If cmbBankAc.ListIndex = "-1" Then
      MsgBox "Please select the bank account.", vbCritical + vbOKOnly, "Batch Process"
      cmbBankAc.SetFocus
      Exit Sub
   End If
   If txtDate.text = "" And Not optBP_MULT.Value Then
       MsgBox "Please enter date.", vbCritical + vbOKOnly, "Batch Process"
       txtDate.SetFocus
       Exit Sub
    End If
    Dim adoconn As New ADODB.Connection
    Dim szSQL As String
    If CBool(optBP_MULT.Value) = False And Trim(txtDate.text) <> "" Then
        adoconn.Open getConnectionString
        If IsPeriodStatus(txtDate.text, txtClient.Tag, adoconn) = 0 Then
           MsgBox "The posting date cannot fall within a closed financial period", vbCritical, "Warning"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDate.text, txtClient.Tag, adoconn) = 9 Then
           MsgBox "The posting date does not fall in any existing financial period", vbCritical, "Warning"
           adoconn.Close
           Set adoconn = Nothing
           FocusControl txtDate
           Exit Sub
        End If
    End If
    If optBP_Cheque.Value = True Then
       If optSCPO(0).Value = False And optSCPO(1).Value = False Then
          MsgBox "Please select which cheque print option you wish to use?", vbExclamation + vbOKOnly, "Batch Payment Process"
          Frame3.ForeColor = vbRed
          Exit Sub
       End If
       If txtCheqNo.text = "" Then
          If MsgBox("You have not entered a next cheque no." & Chr(13) & _
                    "Do you wish to enter a cheque no for this batch?", _
                    vbExclamation + vbYesNo, "Batch Process") = vbYes Then
             txtCheqNo.SetFocus
             Exit Sub
          End If
       End If
    End If

    
    
        
   Me.Hide
    'Resolved by BOSL
    'Issue number 0000447
    'Batch Payments - Incorrect Filltering
    'Modified my Anol 05-08-2014

   frmBatchPayment.vPropertyName = txtProperty.text
   frmBatchPayment.vClientName = txtClient.text
   frmBatchPayment.vPropertyID = txtProperty.Tag
   frmBatchPayment.vClientID = txtClient.Tag
   Load frmBatchPayment

   frmBatchPayment.optBP_Cheque.Value = optBP_Cheque.Value
   frmBatchPayment.optBP_BACS.Value = optBP_BACS.Value
   frmBatchPayment.optBP_MULT.Value = optBP_MULT.Value
   
   frmBatchPayment.lblClient.Caption = txtClient.text
   frmBatchPayment.lblClient.ToolTipText = txtClient.text
   frmBatchPayment.lblProperty.Caption = txtProperty.text
   frmBatchPayment.lblProperty.ToolTipText = txtProperty.text
   frmBatchPayment.lblBank.Caption = cmbBankAc.text
   frmBatchPayment.lblBank.ToolTipText = cmbBankAc.text
   frmBatchPayment.lblBC.Caption = Label13(7).Caption
   frmBatchPayment.lblDate.Caption = txtDate.text
   If optBP_Cheque.Value Then frmBatchPayment.txtChqNo.text = txtCheqNo.text
   If optBP_BACS.Value Then frmBatchPayment.txtChqNo.text = "BACS"
   If optBP_MULT.Value Then frmBatchPayment.txtChqNo.text = "MULTIPLE"

   frmBatchPayment.Show

   SaveCurrentSelection
End Sub

Private Sub SaveCurrentSelection()
    Dim szChoice As String, szaChoice(7) As String
'
'   ReDim szaChoice(4) As String

   szaChoice(0) = IIf(optBP_Cheque.Value, "C", IIf(optBP_BACS.Value, "B", "M"))
   szaChoice(1) = txtClient.Tag
   szaChoice(2) = txtProperty.Tag
    'resolved by bosl
    'issue 455
    'modified by anol 21 Aug 2014
    If cmbBankAc.ListIndex >= 0 Then
         szaChoice(3) = cmbBankAc.Column(0)
    Else
          szaChoice(3) = ""
    End If
   'end of modification
   szaChoice(4) = IIf(optSCPO(0).Value, "RO", "CR")
   szaChoice(5) = Format(lblPostingDate.ToolTipText, "dd/mm/yyyy")
   'added by anol 17 Feb 2016
   szaChoice(6) = txtClient.text
   szaChoice(7) = txtProperty.text
   szChoice = Join(szaChoice, "#")

   SaveSetting "PropertyManagement", "ChoosedOption", "BP-c" & CStr(SCID), szChoice
End Sub

'Private Sub cboClient_Click()
'   'resolved by BOSL
'    'issue 455
'    'Modified by anol 21 Aug 2014
'    Dim adoConn As New ADODB.Connection
'    adoConn.Open getConnectionString
'    LoadProperties adoConn, cboProperty, txtClient.Tag
'    If cboClient.text = "All Clients" Then Exit Sub
'    LoadBank adoConn
'    adoConn.Close
'    Set adoConn = Nothing
'End Sub

Private Sub LoadProperties(adoconn As ADODB.Connection, cboP As Control, szClientID As String)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, j As Integer
   Dim i As Integer, Data() As String
   Dim TotalRow As Integer, TotalCol As Integer

   On Error GoTo ErrorHandler

'***************************************  PROPERTY  ******************************************
   If txtClient.text = "All Clients" Or Trim(txtClient.text) = "" Then
      szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
        szSQL = "SELECT PropertyID, PropertyName, " & _
                  "ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & txtClient.Tag & "' " & _
              "ORDER BY PropertyID;"
      
   End If
   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

   cboP.Clear
   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Properties"
   'issue 571 by anol
   'Date 22 July 2015
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

Private Sub optBP_BACS_Click()
     'Fixed by anol 30 Aug 2016
'       cmbBankAc.Clear
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        LoadBank adoconn
        adoconn.Close
        'End of modification
   If optBP_BACS.Value Then
      'cboClient.ListIndex = -1
      Rem by anol 30 Aug 2016
'       txtClient.text = ""
'       txtClient.Tag = ""
      txtDate.Enabled = True
      lblPostingDate.Enabled = True
      txtCheqNo.Enabled = False
      txtCheqNo.text = ""
      Frame3.Enabled = False
      optSCPO(0).Value = True
      optSCPO(1).Value = False
   End If

   optBP_MULT_Click
End Sub

Private Sub optBP_Cheque_Click()
    'added by anol 30 Aug 2016
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        LoadBank adoconn
        adoconn.Close
   'End of modification
   If optBP_Cheque.Value Then
       Rem by anol 30 Aug 2016
'      txtClient.text = ""
'      txtClient.Tag = ""
      cmbBankAc.Locked = False

      txtCheqNo.Enabled = True
      Frame3.Enabled = True
   End If

   optBP_MULT_Click
End Sub

Private Sub optBP_MULT_Click()
'added by anol 30 Aug 2016
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        LoadBank adoconn
        adoconn.Close
   'End of modification
        
   If optBP_MULT.Value Then
      optSCPO(0).Value = False
      optSCPO(1).Value = False
      Frame3.Enabled = False
      txtDate.text = ""
      txtDate.Enabled = False
      txtCheqNo.text = ""
      txtCheqNo.Enabled = False
       
      cmbBankAc.Locked = False
      lblPostingDate.Enabled = False
      lblPostingDate.ToolTipText = ""
   Else
      If Not optBP_BACS.Value Then
         Frame3.Enabled = True
         txtDate.Enabled = True
         txtCheqNo.Enabled = True
         lblPostingDate.Enabled = True
      End If
   End If
End Sub

Private Sub optSCPO_Click(Index As Integer)
   Frame3.ForeColor = vbDefault
End Sub

Private Sub txtCheqNo_GotFocus()
   SelTxtInCtrl txtCheqNo
End Sub

Private Sub txtCheqNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OKButton.SetFocus
    End If
   DigitTextKeyPress txtCheqNo, KeyAscii, 0
End Sub

Private Sub txtDate_Change()
   'Resolved by BOSL
   'Issue 468
   'Modified by Anol 03 Sep 2014
   TextBoxChangeDate txtDate
   lblPostingDate.ToolTipText = txtDate.text
End Sub

Private Sub txtDate_GotFocus()
   If txtDate.text = "dd/mm/yyyy" Then
      txtDate.text = ""
      Exit Sub
   End If
   
   If Len(txtDate.text) < 10 Then txtDate.text = Format(Date, "dd/mm/yyyy")
   If lblPostingDate.ToolTipText = "" Then lblPostingDate.ToolTipText = txtDate.text
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCheqNo.Enabled = False Then
            FocusControl OKButton
        Else
            FocusControl txtCheqNo
        End If
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
'    If cboClient.ListIndex = "-1" Then
'          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
'          cboClient.SetFocus
'          Exit Sub
'    End If
     If txtClient.text = "" Then
          ShowMsgInTaskBar "Please select a Client.", "Y", "N"
          cmdClient.SetFocus
          Exit Sub
    End If
    'Added by anol 26 Mar 2015
    If IsDate(txtDate.text) = False Then Exit Sub
    If frmMMain.IsRibbonVersion Then
        Dim adoconn As New ADODB.Connection
        Dim szSQL As String
        adoconn.Open getConnectionString
        If IsPeriodStatus(txtDate.text, txtClient.Tag, adoconn) = 0 Then
           MsgBox "The posting date cannot fall within a closed financial period", vbInformation, "Warning"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(txtDate.text, txtClient.Tag, adoconn) = 9 Then
           MsgBox "The posting date does not fall in any existing financial period", vbInformation, "Warning"
           adoconn.Close
           Set adoconn = Nothing
           Exit Sub
        End If
    End If
End Sub

Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
         flxClient.SetFocus
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
