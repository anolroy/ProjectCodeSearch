VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSCDetailsSplit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Charge Details"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   135
      TabIndex        =   9
      Top             =   4095
      Visible         =   0   'False
      Width           =   6630
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
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClientList 
         Height          =   4335
         Left            =   90
         TabIndex        =   7
         Top             =   675
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   7646
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
      Begin VB.Shape Shape2 
         Height          =   5010
         Left            =   0
         Top             =   45
         Width           =   6585
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   180
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "Client ID"
         Size            =   "1296;353"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientID 
         Height          =   255
         Left            =   270
         TabIndex        =   4
         Top             =   390
         Width           =   1350
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2381;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1665
         TabIndex        =   5
         Top             =   390
         Width           =   3555
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6271;450"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Index           =   1
         Left            =   1635
         TabIndex        =   11
         Top             =   195
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
         Index           =   2
         Left            =   5175
         TabIndex        =   10
         Top             =   180
         Width           =   1185
         VariousPropertyBits=   8388627
         Caption         =   "Balance"
         Size            =   "2090;344"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   255
         Left            =   5265
         TabIndex        =   6
         Top             =   390
         Width           =   1170
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2064;450"
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
         Left            =   90
         Top             =   135
         Width           =   6345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5460
      Left            =   45
      TabIndex        =   13
      Top             =   -45
      Width           =   6855
      Begin VB.CommandButton cmdNCode 
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
         Left            =   6495
         TabIndex        =   0
         Top             =   375
         Width           =   300
      End
      Begin VB.TextBox txtNCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
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
         Height          =   315
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   375
         Width           =   4845
      End
      Begin VB.TextBox txtNName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
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
         Height          =   315
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   4890
      End
      Begin VB.TextBox txtBudget 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
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
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   1320
         Width           =   1200
      End
      Begin VB.CommandButton cmdSCDBdClose 
         Caption         =   "Save"
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
         Left            =   3600
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSCDBdCancel 
         Caption         =   "&Cancel"
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
         Left            =   1575
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Name"
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
         Left            =   135
         TabIndex        =   18
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal Code"
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
         Left            =   135
         TabIndex        =   17
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   135
         TabIndex        =   16
         Top             =   1320
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmSCDetailsSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strCommandSource As String

Private Sub cmdGridUnitLookup_Click()
    '      Frame1(1).Enabled = True
'    Frame5.Visible = False
    FocusControl cmdNCode
    Frame5.Visible = False
    Frame1.Enabled = True
End Sub

Private Sub cmdSCDBdCancel_Click()
   Unload Me
End Sub

Private Sub cmdSCDBdClose_Click()
   If txtNName.text = "" Or txtNCode.text = "" Then
      ShowMsgInTaskBar "Please select a nominal.", , "N"
      cmdNCode.SetFocus
      Exit Sub
   End If
   If txtBudget.text = "" Then
      ShowMsgInTaskBar "Please enter an amount.", , "N"
      txtBudget.SetFocus
      Exit Sub
   End If
    Dim iRow As Integer
    Dim isfound As Boolean
  
   
   With frmServiceChargeDetails
        For iRow = 1 To .flxSCBudgetDetailsAnalysis.Rows - 1
            If txtNCode.text = .flxSCBudgetDetailsAnalysis.TextMatrix(iRow, 2) Then
                isfound = True
                Exit For
            End If
        Next iRow
        If frmServiceChargeDetails.bNewLine And isfound Then
            ShowMsgInTaskBar "An entry already exists for this nominal code.", , "N"
            Exit Sub
        End If
        'issue 471 Note 740
        'Modified by anol 05 Nov 2014
        If frmServiceChargeDetails.bNewLine = True Then
            .flxSCBudgetDetailsAnalysis.AddItem ""
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 0) = UniqueID()
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 1) = frmServiceCharge.txtBudgetId.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 2) = txtNCode.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 3) = txtNName.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 4) = txtBudget.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.Rows - 1, 5) = "N"             'N->New Line
            .SCSumTotal
        Else
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 2) = txtNCode.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 3) = txtNName.text
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 4) = Format((txtBudget.text), "0.00")
             
            If .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 5) = "" Then _
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 5) = "A"             'A->Amended
            If .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 5) = "N" Then _
            .flxSCBudgetDetailsAnalysis.TextMatrix(.flxSCBudgetDetailsAnalysis.row, 5) = "NA"             'NA->New Line Amended
           
            .SCSumTotal
        End If
        'End of modification
   End With
   Unload Me
End Sub

Private Sub cmdNCode_Click()
'    Frame1(1).Enabled = False
    strCommandSource = "NC"
    Call LoadNC
    Frame5.Top = 90
    Frame5.Left = 130
    Frame5.Visible = True
    Frame1.Enabled = False
    FocusControl txtSearchClientID
End Sub

Private Sub flxClientList_Click()
        If strCommandSource = "NC" Then
                txtNCode.Tag = flxClientList.TextMatrix(flxClientList.row, 1) '1 ID,2 code,3 fund Name
                txtNCode.text = flxClientList.TextMatrix(flxClientList.row, 1)
                txtNName.text = flxClientList.TextMatrix(flxClientList.row, 2)
'                txtTotalArea.text = "1"
'                FocusControl txtTotalArea
        End If
       
        Frame5.Visible = False
        Frame1.Enabled = True
        FocusControl txtBudget
End Sub

Private Sub flxClientList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClientList_Click
    End If
End Sub

Private Sub Form_Load()
   Me.Height = 5970
   Me.Width = 7020
   
   LoadNC
    Call WheelHook(Me.hWnd)
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
Private Sub LoadNC()
  'My Ideal loading flexgrid component by anol 2020-12-17
  'Learning: inside a picturebox you cannot resize a Textbox, I am I am adding frame and shape to replace this picturebox
   Dim rRow As Integer
   Dim szSQL As String
   Dim iSel As Integer
   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   Dim rsFundMatrix As New ADODB.Recordset
   'you just change label position then searchbox and grid coulumn will try to fit accordingly
   lblClientID(0).Left = 250
   lblClientID(1).Left = 1365
   lblClientID(2).Left = 3510

   flxClientList.RowHeight(0) = 0
   flxClientList.Cols = 3
   flxClientList.ColWidth(0) = 200
   flxClientList.ColWidth(1) = lblClientID(1).Left - lblClientID(0).Left
   
   
   txtSearchClientID.Width = lblClientID(1).Left - lblClientID(0).Left - 20
   txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   TextBox1.Width = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left - 20
   
     
   If flxClientList.Cols > 3 Then
        flxClientList.ColWidth(2) = lblClientID(2).Left - lblClientID(1).Left
        txtSearchClientName.Width = lblClientID(2).Left - lblClientID(1).Left - 20
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(2) = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
        txtSearchClientName.Width = cmdGridUnitLookup.Left - lblClientID(1).Left - 200
   End If
   If flxClientList.Cols = 4 Then
        flxClientList.ColWidth(3) = cmdGridUnitLookup.Left + cmdGridUnitLookup.Width - lblClientID(2).Left
        TextBox1.Visible = True
   ElseIf flxClientList.Cols = 3 Then
        flxClientList.ColWidth(3) = 0
        TextBox1.Visible = False
   End If
   
   
   txtSearchClientName.Visible = True

   
   flxClientList.Clear
   flxClientList.Rows = 2
   flxClientList.ColAlignment(0) = vbLeftJustify
   flxClientList.ColAlignment(1) = vbLeftJustify
   flxClientList.ColAlignment(2) = vbLeftJustify
   flxClientList.ColAlignment(3) = vbLeftJustify
   
   lblClientID(0).Caption = "Nominal Code"
   lblClientID(1).Caption = "Nominal Name"
   lblClientID(2).Caption = ""
   
   txtSearchClientID.Left = lblClientID(0).Left
   txtSearchClientName.Left = lblClientID(1).Left
   
   
   TextBox1.Left = lblClientID(2).Left
   TextBox1.Width = cmdGridUnitLookup.Left - lblClientID(2).Left + 40
   
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   TextBox1.text = ""
    
   adoconn.Open getConnectionString
   'szSQL = "SELECT ID, TYPE FROM DemandTypes where PropertyID='" & szPropertySelection1 & "';"
   
   'rsFundMatrix.Open "Select isfundAssign from shoppingcentre", adoconn, adOpenStatic, adLockReadOnly
'   If rsFundMatrix("isfundAssign").Value = False Then
        iSel = 0
'        szSQL = "SELECT FundID, FundName, FundCode,CategoryCode FROM Fund;"
          szSQL = "SELECT NominalLedger.* " & _
           "FROM NominalLedger " & _
           "WHERE Type = 2 AND ClientID = '" & frmServiceCharge.txtClientList.Tag & "' " & _
           "ORDER BY Code;"

'   Else
'        iSel = 1
'        szSQL = "Select F.* from Fund F,fundMatrix M where F.FundID=M.FundID AND PropertyID='" & _
'                szPropertySelection1 & "' and ClientID='" & szClientID & "' and isDeleted=false"
'   End If
   'rsFundMatrix.Close
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
   If rstRec.EOF Then
        If iSel = 0 Then
            ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
         Else
            ShowMsgInTaskBar "There are no funds assigned for this property. Please assign a fund.", , "N"
         End If
      flxClientList.Clear
      flxClientList.Rows = 2
   Else
                rRow = 1
                While Not rstRec.EOF
                    flxClientList.row = 1
                    flxClientList.RowSel = 1
                    flxClientList.ColSel = 1
                    flxClientList.TextMatrix(rRow, 0) = ""
                    flxClientList.TextMatrix(rRow, 1) = rstRec.Fields.Item("Code").Value
                    flxClientList.TextMatrix(rRow, 2) = rstRec.Fields.Item("Name").Value
                    flxClientList.TextMatrix(rRow, 3) = ""
                    flxClientList.RowHeight(rRow) = 280
                    rstRec.MoveNext
                    If Not rstRec.EOF Then flxClientList.AddItem ""
                    rRow = rRow + 1
                 Wend
         
   End If
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub

Private Sub TextBox1_Change()
     Dim i As Integer

   If Len(txtSearchClientName.text) > 0 Then
        txtSearchClientID.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
        flxClientList.RowHeight(i) = 240
       
        If InStr(1, UCase(flxClientList.TextMatrix(i, 2)), UCase(TextBox1.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
        End If
       
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientID_Change()
     Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240

      If InStr(1, UCase(flxClientList.TextMatrix(i, 1)), UCase(txtSearchClientID.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
End Sub

Private Sub txtSearchClientName_Change()
           'Updated by anol 10 Dec 2015
                  
   Dim i As Integer

   If Len(txtSearchClientID.text) > 0 Then
        txtSearchClientName.text = ""
   End If

   For i = flxClientList.Rows - 1 To 1 Step -1
      flxClientList.RowHeight(i) = 240

      If InStr(1, UCase(flxClientList.TextMatrix(i, 2)), UCase(txtSearchClientName.text), vbTextCompare) = 0 Then
            flxClientList.RowHeight(i) = 0
      End If
      If flxClientList.RowHeight(i) = 240 Then
            flxClientList.row = i
      End If
   Next i
  
End Sub
'Private Sub txtBudget_Change()
'   If IsNumeric(txtBudget.text) = False Then
'      txtBudget.text = ""
'   End If
'End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSCDBdClose
    End If
   DigitTextKeyPress1 txtBudget, KeyAscii
   
   
End Sub
Public Sub DigitTextKeyPress1(conTextBox As Control, ByRef KeyAscii As Integer, Optional digitAfterDot As Integer = 2)
   If KeyAscii = 27 Then
      conTextBox.text = ""
   End If
   If KeyAscii = 45 Then
      Exit Sub
   End If
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   
   If KeyAscii = 46 Then
      If InStr(1, conTextBox.text, ".") > 0 Or digitAfterDot = 0 Then
         KeyAscii = 0
      End If
   End If
   
   If InStr(1, conTextBox.text, ".") > 0 Then
      If (Len(conTextBox.text) - InStr(1, conTextBox.text, ".") = digitAfterDot) And _
               (conTextBox.SelStart >= InStr(1, conTextBox.text, ".")) And _
               KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End If
End Sub
'Private Sub cboNCode_Change()
'   If Trim(txtNCode.text) = "" Then Exit Sub
'   If cboNCode.ListIndex >= 0 Then txtNName.text = cboNCode.Column(1)
'End Sub

'Private Sub LoadNC()
'   On Error GoTo Error_Handler
'
'   Dim adoconn    As ADODB.Connection
'   Dim rRow       As Integer
'   Dim iRec       As Integer
'   Dim Data()     As String
'   Dim adoRst     As New ADODB.Recordset
'   Dim szSQL      As String
'
'   Set adoconn = New ADODB.Connection
'   adoconn.Open getConnectionString
'
'   szSQL = "SELECT NominalLedger.* " & _
'           "FROM NominalLedger " & _
'           "WHERE Type = 2 AND ClientID = '" & frmServiceCharge.txtClientList.Tag & "' " & _
'           "ORDER BY Code;"
'
'   adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
'
'   If adoRst.EOF Then
'      ShowMsgInTaskBar "Fund has not been setup for this company.", , "N"
'   Else
'      Dim count As Integer
'      count = adoRst.RecordCount
'      ReDim Data(2, count) As String
'      rRow = 0
'      While Not adoRst.EOF
'         Data(1, rRow) = Trim(adoRst.Fields.Item("Name").Value)
'         Data(0, rRow) = Trim(adoRst.Fields.Item("Code").Value)
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'
'      cboNCode.Clear
'      cboNCode.Column() = Data()
'   End If
'
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoconn = Nothing
'
'   Exit Sub
'
'   ' Error Handling Code
'Error_Handler:
'
'   ShowMsgInTaskBar "Error in Loading fund.", , "N"
'   ' Destroy Objects
'   Set adoRst = Nothing
'   Set adoconn = Nothing
'End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmServiceChargeDetails.Enabled = True
   UnLoadForm Me
End Sub

Private Sub txtBudget_LostFocus()
   If txtBudget.text = "" Then Exit Sub

   txtBudget.text = Format(txtBudget.text, "0.00")
End Sub




Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl flxClientList
    End If
End Sub
