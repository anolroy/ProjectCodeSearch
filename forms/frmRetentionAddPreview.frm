VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRetentionAddPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Retentions"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRetentionAddPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   13350
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   1215
      ScaleHeight     =   4740
      ScaleWidth      =   6255
      TabIndex        =   20
      Top             =   4365
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
         TabIndex        =   24
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   4020
         Left            =   45
         TabIndex        =   23
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   4
         Left            =   1515
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   21
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
         TabIndex        =   22
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   90
      TabIndex        =   19
      Top             =   3555
      Width           =   13020
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   355
         Left            =   9180
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   355
         Left            =   10890
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Left            =   90
      TabIndex        =   12
      Top             =   45
      Width           =   13020
      Begin VB.TextBox txtSLNumber 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   1
         Top             =   810
         Width           =   5100
      End
      Begin VB.TextBox txtStatementID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   5100
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
         Left            =   6165
         TabIndex        =   3
         Top             =   1710
         Width           =   300
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
         Left            =   6165
         TabIndex        =   2
         Top             =   1305
         Width           =   300
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11100
         TabIndex        =   7
         Top             =   855
         Width           =   1470
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   254
         TabIndex        =   4
         Top             =   2250
         Width           =   5100
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2745
         Width           =   5100
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   11100
         TabIndex        =   6
         Top             =   360
         Width           =   1465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SL No:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   30
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Statement ID:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   29
         Top             =   360
         Width           =   960
      End
      Begin MSForms.TextBox txtProperty 
         Height          =   285
         Left            =   1350
         TabIndex        =   11
         Top             =   1710
         Width           =   4860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8572;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   1350
         TabIndex        =   10
         Top             =   1305
         Width           =   4860
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "8572;503"
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
         Index           =   5
         Left            =   270
         TabIndex        =   18
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   3
         Left            =   8400
         TabIndex        =   17
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         Height          =   195
         Index           =   10
         Left            =   8400
         TabIndex        =   16
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   15
         Top             =   2745
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   11
         Left            =   270
         TabIndex        =   14
         Top             =   2250
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   13
         Top             =   1710
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmRetentionAddPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vatOptionEnabled As Boolean
Public frmRetentionAdd_CALLING_FROM As String             'Name of the form
'Public frmRetentionAdd_CALLING_MODE As String             'Calling for Add or Edit Bank Transacitons
Private bUnitLoaded  As Boolean
Public szTransID     As String
Dim sTextBox As String
Dim szSQL As String







Private Sub FillCboVatCode(adoConnection As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim SQLStr1 As String

   

   SQLStr1 = "SELECT VAT_ID, VAT_CODE, VAT_RATE FROM TLBVATCODE where IN_USE Order By VAT_ID"
   adoRst.Open SQLStr1, adoConnection, adOpenDynamic, adLockPessimistic

   txtSearchClientID.text = ""
   txtSearchClientID.Left = 250
   
   txtSearchClientID.Width = 3200
   txtSearchClientName.Visible = False
   
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   
   flxClient.ColWidth(0) = 100
   flxClient.ColWidth(1) = 1200
   flxClient.ColWidth(2) = 1800
   picClient.Width = 3500
   flxClient.Width = 3300
   cmdPicCLose.Left = 3200
   txtSearchClientID.Left = 145
   
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify

   '~~~ Added by Anol Configuring width and position of labels and search boxes.
   lblClientID.Caption = "VAT Code"
   lblClientName.Caption = ""
   lblClientID.Width = 1400
   lblClientID.Left = 250
   lblClientName.Width = 3600
   Dim rRow As Integer
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
    
    rRow = 1
    While Not adoRst.EOF
       flxClient.row = 1
       flxClient.TextMatrix(rRow, 0) = "  " & adoRst.Fields.Item("VAT_ID").Value
       flxClient.TextMatrix(rRow, 1) = adoRst.Fields.Item("VAT_CODE").Value
       flxClient.TextMatrix(rRow, 2) = adoRst.Fields.Item("VAT_RATE").Value
        flxClient.RowHeight(rRow) = 280
       adoRst.MoveNext
       If Not adoRst.EOF Then flxClient.AddItem ""
       rRow = rRow + 1
    Wend
  
   adoRst.Close
   Set adoRst = Nothing
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
                        
                            bHandled = False
                       

        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then FocusControl (ctl)

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

Private Sub cmdNC_Click()
    picClient.Left = 6500.029
    picClient.Top = 455.299
    sTextBox = "5"
    Call LoadflxNC  'calling the main account here
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdUnit_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "4"
    LoadflxUnit
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdVatCode_Click()
    picClient.Left = 8910.029
    picClient.Top = 200.299
    sTextBox = "7"
    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString
    Call FillCboVatCode(adoconn)
    
    adoconn.Close
    
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub Form_Activate()
        Dim strTemp As String
        txtClientList.ForeColor = vbBlack
        If Trim(txtClientList.Tag) = "" Then Exit Sub
        strTemp = Trim(isControlAccountSet(txtClientList.Tag))
        If Len(strTemp) > 0 Then
            MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientList.Tag & _
            vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
            strTemp = ""
            txtClientList.ForeColor = vbRed
            Exit Sub
        End If
End Sub

Private Function MaxStatementID() As Integer
    Dim adoconn As New ADODB.Connection
    Dim rsMaxstatementID As New ADODB.Recordset
    adoconn.Open getConnectionString
    rsMaxstatementID.Open "Select max(statementID) as ID from rentsummaryStatement", adoconn, adOpenStatic, adLockReadOnly
    MaxStatementID = IIf(IsNull(rsMaxstatementID("ID").Value), 0, rsMaxstatementID("ID").Value) + 1
    rsMaxstatementID.Close
   adoconn.Close
End Function




Private Sub cboFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl txtDate
    End If
End Sub





Private Sub cmdBC_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "2"
    LoadflxBank
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    FocusControl cmdClientList
End Sub
'Private Sub cboProperty_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cboUnit1.SetFocus
'    End If
'End Sub

Private Sub cboUnit1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDetails
    End If
End Sub

'Private Sub cboVat_Click()
'   'Modified by BOSL
'    'issue 463 manual vat
'    'Anol 20 Aug 2014
'    If txtNet.text = "" Then Exit Sub
'    txtVat_.text = Format(Val(txtNet.text) * (cboVat.text / 100), "0.00")
'    txtTotal.text = Format(Val(txtNet.text) * (1 + (cboVat.text / 100)), "0.00")
'End Sub

Private Sub cmdClientList_Click()
    picClient.Left = 269.029
    picClient.Top = 255.299
    sTextBox = "1"
    LoadflxClient
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub
Private Sub LoadflxBank()
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
   lblClientID.Caption = "Bank Nominal Code"
   lblClientName.Caption = "Bank Nominal Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   'lblClientName.Left = lblClientID.Left + flxClient.ColWidth(0)
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240

   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   
   adoconn.Open getConnectionString
  szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "NominalLedger.Name AS BNN, AllowOverDraft, OverDraftLimit " & _
           "FROM tlbClientBanks, NominalLedger " & _
           "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
               "tlbClientBanks.CLIENT_ID = NominalLedger.ClientID AND " & _
               "tlbClientBanks.CLIENT_ID <> '' AND " & _
               "NominalLedger.ClientID = '" & txtClientList.Tag & "';"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
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
Private Sub LoadflxUnit()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
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
   lblClientID.Caption = "Unit No"
   lblClientName.Caption = "Unit Name"
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
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   
   adoconn.Open getConnectionString
           
'        szSQL = "SELECT PropertyID, PropertyName " & _
'                    "FROM Property " & _
'                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'                    "ORDER BY PropertyID;"
        If txtClientList.text <> "" Then
             szSQL = " Client.ClientName='" & txtClientList.text & "'"
        End If
        If txtProperty.text <> "" Then
             szSQL = " Property.PropertyName='" & txtProperty.text & "'"
        End If
        If szSQL <> "" Then
             szSQL = " where" & szSQL
        End If
        szSQL = "SELECT UnitNumber, UnitName FROM (Units INNER JOIN Property ON Units.PropertyID=Property.PropertyID) INNER JOIN client ON client.ClientID=Property.ClientID " & szSQL & ";"
       ' Debug.Print szSQL
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
            
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
           'rRow = rRow + 1
           flxClient.AddItem ""
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = ""
           flxClient.TextMatrix(rRow, 2) = ""
           flxClient.RowHeight(rRow) = 280
           flxClient.AddItem ""
           'rRow = 2
   
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadflxProperty()
    Dim rRow As Integer
   Dim szSQL As String

   Dim adoconn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset
   txtSearchClientID.text = ""
   txtSearchClientName.text = ""
   flxClient.RowHeight(0) = 0
   flxClient.Cols = 3
   flxClient.ColWidth(0) = 100
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

   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   adoconn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
            rRow = 1
            
        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
               flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = ""
           flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item(0).Value
           flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item(1).Value
           flxClient.RowHeight(rRow) = 280
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
        
        flxClient.AddItem ""
        flxClient.TextMatrix(rRow, 0) = ""
        flxClient.TextMatrix(rRow, 1) = ""
        flxClient.TextMatrix(rRow, 2) = ""
        flxClient.RowHeight(rRow) = 280
        flxClient.AddItem ""
'           rRow = 2
           
   rstRec.Close
   adoconn.Close
   Set rstRec = Nothing
   Set adoconn = Nothing
End Sub
Private Sub LoadflxNC()
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
   lblClientID.Caption = "Nominal Code"
   lblClientName.Caption = "Nominal Name"
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
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new
   
   adoconn.Open getConnectionString
  szSQL = "SELECT N.* " & _
      "FROM NominalLedger AS N " & _
      "WHERE N.ClientID = '" & txtClientList.Tag & "' AND " & _
      "Posting AND CAFixed=0 AND CODE NOT IN " & _
      "(SELECT NominalCode FROM tlbClientBanks where ClientID = '" & txtClientList.Tag & "')" & _
      " ORDER BY N.Code;"

   rstRec.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
           
           rRow = 1
           While Not rstRec.EOF
               flxClient.row = 1
               flxClient.RowSel = 1
               flxClient.ColSel = 1
               flxClient.TextMatrix(rRow, 0) = ""
               flxClient.TextMatrix(rRow, 1) = rstRec.Fields.Item("CODE").Value
               flxClient.TextMatrix(rRow, 2) = rstRec.Fields.Item("NAME").Value
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
   flxClient.Clear
   flxClient.Rows = 2
   flxClient.ColAlignment(0) = vbLeftJustify
   flxClient.ColAlignment(1) = vbLeftJustify
   flxClient.ColAlignment(2) = vbLeftJustify
   lblClientID.Caption = "Client ID"
   lblClientName.Caption = "Client Name"
   lblClientID.Width = 1400
   lblClientID.Left = 50
   lblClientName.Width = 2600
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   txtSearchClientID.Left = 45
   
   'New
   
   picClient.Width = 5295
   cmdPicCLose.Left = 5010
   lblClientName.Left = lblClientID.Left + flxClient.ColWidth(1) + flxClient.ColWidth(0)
   txtSearchClientName.Width = 3240
 
   picClient.Height = 4095
   flxClient.Height = 3345
   flxClient.Width = 5175
   
   'End of new

   
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

Private Sub cmdproperty_Click()
    picClient.Left = 269.029
    picClient.Top = 455.299
    sTextBox = "3"
    LoadflxProperty
    Frame2.Enabled = False
    Frame1.Enabled = False
    picClient.Visible = True
    FocusControl txtSearchClientID
End Sub

Private Function LoadVatOption(Conn As ADODB.Connection) As Integer
    Dim rsGlobalData As New ADODB.Recordset
    rsGlobalData.Open "Select vatOptionEnabled from Globaldata where PropertyID='" & txtProperty.Tag & "'", Conn, adOpenStatic, adLockReadOnly
    If Not rsGlobalData.EOF Then
            LoadVatOption = IIf(IsNull(rsGlobalData("vatOptionEnabled").Value), 0, rsGlobalData("vatOptionEnabled").Value)
    End If
    rsGlobalData.Close
    Set rsGlobalData = Nothing
End Function
Private Sub flxClient_Click()
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
        Dim szSQL As String
        Dim rstRec As New ADODB.Recordset
        Dim rstVat As New ADODB.Recordset
        Dim nTaxCode As Double
        Frame1.Enabled = True
        Frame2.Enabled = True
        If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                LoadFirstProperty adoconn
                FocusControl cmdProperty
        End If
        If sTextBox = "3" Then
                txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                FocusControl txtDetails
              
        End If
        picClient.Visible = False
        adoconn.Close
        Set adoconn = Nothing
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        picClient.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
         If sTextBox = "1" Then
            FocusControl cmdClientList
         ElseIf sTextBox = "3" Then
            FocusControl cmdProperty
         End If
        
    End If
    
    If KeyAscii = 13 Then
        Dim adoconn As New ADODB.Connection
        adoconn.Open getConnectionString
       
        Frame1.Enabled = True
        Frame2.Enabled = True
            If sTextBox = "1" Then
                txtClientList.Tag = flxClient.TextMatrix(flxClient.row, 1)
                txtClientList.text = flxClient.TextMatrix(flxClient.row, 2)
                'LoadBankAccountInCombo adoconn
'                LoadNCinCombo adoconn
                LoadFirstProperty adoconn
                FocusControl cmdProperty
            ElseIf sTextBox = "2" Then
                    'txtBankCode.text = flxClient.TextMatrix(flxClient.row, 1)
                    'txtBankName.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl cmdProperty
            ElseIf sTextBox = "3" Then
                    txtProperty.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    txtProperty.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl txtDetails
            ElseIf sTextBox = "4" Then
'                    txtUnit.Tag = flxClient.TextMatrix(flxClient.row, 1)
                    'txtUnit.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl txtDetails
            ElseIf sTextBox = "5" Then
'                    txtNC.Tag = flxClient.TextMatrix(flxClient.row, 1)
'                    txtNC.text = flxClient.TextMatrix(flxClient.row, 2)
'                    FocusControl cmdFund
'             ElseIf sTextBox = "6" Then
'                    txtFund.Tag = flxClient.TextMatrix(flxClient.row, 0)
'                    txtFund.text = flxClient.TextMatrix(flxClient.row, 2)
                    FocusControl txtDate
'            ElseIf sTextBox = "7" Then
'                    Label1(24).Caption = flxClient.TextMatrix(flxClient.row, 1)
'                    Label1(24).Tag = flxClient.TextMatrix(flxClient.row, 2)
'                    txtVat_.text = Format(Val(txtNet.text) * (Val(flxClient.TextMatrix(flxClient.row, 2)) / 100), "0.00")
'                    txtTotal.text = Format(Val(txtNet.text) + Val(txtVat_.text), "0.00")
'                    FocusControl cmdSave
            End If
        
        picClient.Visible = False
        adoconn.Close
        Set adoconn = Nothing
    End If
End Sub

Private Sub Label13_Click(Index As Integer)
   ' MsgBox Label13(7).Caption
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub



Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtDate
    End If
End Sub

'Private Sub txtNC_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        cmdNC.SetFocus
'    End If
'End Sub

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
        If sTextBox <> 7 Then
            FocusControl txtSearchClientName
           End If
    End If
End Sub

Private Sub txtSearchClientID_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 27 Then
'         picClient.Visible = False
'          Frame1.Enabled = True
'          Frame2.Enabled = True
'          If sTextBox = "1" Then
'                 cmdClientList.SetFocus
''           ElseIf sTextBox = "2" Then
''                cmdproperty.SetFocus
''           ElseIf sTextBox = "3" Then
''                cmdFundLookUp.SetFocus
'           End If
'    End If
If KeyAscii = 27 Then
        picClient.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
         If sTextBox = "1" Then
         ElseIf sTextBox = "1" Then
            FocusControl cmdClientList
'         ElseIf sTextBox = "2" Then
'            FocusControl cmdBC
         ElseIf sTextBox = "3" Then
            FocusControl cmdProperty
'         ElseIf sTextBox = "4" Then
'            FocusControl cmdUnit
'
'         ElseIf sTextBox = "5" Then
'            FocusControl cmdNC
'         ElseIf sTextBox = "6" Then
'            FocusControl cmdFund
'         ElseIf sTextBox = "7" Then
'            FocusControl cmdVATCode
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
         FocusControl flxClient
    End If
    If KeyCode = vbKeyDown Then
        If flxClient.Visible Then
            FocusControl flxClient
        End If
    End If
End Sub

Private Sub txtClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdClientList
    End If
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtReference
    End If
End Sub

'Private Sub txtReference_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then
'        FocusControl cmdNC
'    End If
'End Sub

'Private Sub txtSearchClientName_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 27 Then
'        picClient.Visible = False
'        Frame1.Enabled = True
'        Frame2.Enabled = True
'         If sTextBox = "1" Then
'         ElseIf sTextBox = "1" Then
'            FocusControl cmdClientList
'         ElseIf sTextBox = "2" Then
'            FocusControl cmdBC
'         ElseIf sTextBox = "3" Then
'            FocusControl cmdProperty
'         ElseIf sTextBox = "4" Then
'            FocusControl cmdUnit
'
'         ElseIf sTextBox = "5" Then
'            FocusControl cmdNC
'         ElseIf sTextBox = "6" Then
'            FocusControl cmdFund
'         ElseIf sTextBox = "7" Then
'            FocusControl cmdVATCode
'         End If
'
'    End If
'End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub



'Private Sub txtUnit_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    If KeyAscii = 13 Then
'        FocusControl cmdUnit
'    End If
'End Sub

'Private Sub txtVat__Change()
'    'Resolved by BOSL
'    'issue 463 manual vat
'    'newly added by Anol 20 Aug 2014
'    If IsNumeric(txtVat_.text) = False Then
'        txtVat_.text = "0.00"
'    End If
'End Sub

Private Sub txtVat__KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
End Sub
'
'Private Sub txtVat__LostFocus()
'    'Resolved by BOSL
'    'issue 463 manual vat
'    'newly added by Anol 20 Aug 2014
'     txtVat_.text = Format(txtVat_.text, "0.00")
'     txtTotal.text = Format((Val(txtNet.text) + Val(txtVat_.text)), "0.00")
'End Sub
Private Sub cboVat_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
    If txtClientList.ForeColor = vbRed Then
        MsgBox "No Nominal Account Codes have been setup in the Control Accounts for the Client: " & txtClientList.text & _
        vbNewLine & "Please setup the Control Accounts in Tools > Configuration > Control Accounts"
        Exit Sub
    End If
    If txtClientList.text = "" Then
      MsgBox "Please select a Client.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl cmdClientList
      
      Exit Sub
   End If

   If Trim(txtProperty.text) = "" Then
          If MsgBox("You have not selected a property. Do you wish to add a property?", vbYesNo, "Select a Property") = vbYes Then
              cmdProperty.SetFocus
              Exit Sub
          End If
   End If
   If txtDetails.text = "" Then
      MsgBox "Please enter Description.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtDetails
      
      Exit Sub
   End If
  If txtStatementID.text = "" Then
      MsgBox "Please enter statement ID.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtStatementID
      
      Exit Sub
   End If
   
   If txtDate.text = "" Then
      MsgBox "Please enter the Retention date.", vbExclamation + vbOKOnly, "Saving..."
      FocusControl txtDate
      Exit Sub
   End If

    Dim adoconn As New ADODB.Connection
    adoconn.Open getConnectionString

    cmdSave.Enabled = False
    If Left(Me.Caption, 3) = "Add" Then
        If Not SaveNewTrans Then Exit Sub
        txtClientList.text = ""
        txtClientList.Tag = ""
        txtProperty.text = ""
        txtProperty.Tag = ""
        
        
        txtDetails.text = ""
        txtReference.text = ""

      If MsgBox("Would you like to add another Retention Preview?", vbQuestion + vbYesNo) = vbNo Then
         Unload Me
      Else
         txtDetails.text = ""
         txtReference.text = ""
         
         
         
         txtNet.text = ""
         cmdSave.Enabled = True
'         Label1(24).Caption = ""
'         Label1(24).Tag = ""
         
         'modidfied by anol 20161013
         'cmdClientList.SetFocus
         Call clientFocus
      End If
   End If
   If Left(Me.Caption, 4) = "Edit" Then
'        SaveEditTrans
        If IsLoadedAndVisible("frmCashbook") = True Then
            frmCashbook.cboBC_Click
        End If
   End If
End Sub
Private Sub clientFocus()
    On Error GoTo Err
    cmdClientList.SetFocus
    Exit Sub
    
Err:
End Sub

Private Function SaveNewTrans() As Boolean
   Dim szStr As String
   Dim adoconn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

   'On Error GoTo ErrHandler
   adoconn.Open getConnectionString
  

   szStr = "SELECT * FROM RetentionDetailsPrev"

   adoRst.Open szStr, adoconn, adOpenDynamic, adLockOptimistic
   With adoRst
      .AddNew
      .Fields.Item("statementID").Value = txtStatementID.text
      .Fields.Item("SLNumber").Value = txtSLNumber.text
      .Fields.Item("ClientID").Value = txtClientList.Tag
      .Fields.Item("PropertyID").Value = txtProperty.Tag
      .Fields.Item("DESCRIPTION").Value = txtDetails.text
      .Fields.Item("Reference").Value = txtReference.text
      .Fields.Item("RDATE").Value = txtDate.text
      .Fields.Item("AMOUNT").Value = Val(txtNet.text)
      .Update
      .Close
    End With
    SaveNewTrans = True
    MsgBox "Retention has been saved successfully"
End Function



Private Sub Form_Load()
   Dim adoconn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Me.Height = 4930
   Me.Width = 13230
    adoconn.Open getConnectionString
    szSQL = "SELECT CLIENTID, CLIENTNAME FROM CLIENT order by CLIENTID"
    adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly
    If Not adoRst.EOF Then
            txtClientList.Tag = adoRst.Fields("CLIENTID").Value
            txtClientList.text = adoRst.Fields("CLIENTNAME").Value
            adoRst.Close
    End If
   adoconn.Close
   Set adoconn = Nothing
   'txtStatementID.text = MaxStatementID
   Call WheelHook(Me.hWnd)

End Sub



Private Sub LoadFirstProperty(adoconn As ADODB.Connection) 'load first property if there is only one property issue 713
   Dim adoRst As New ADODB.Recordset

    '*************************************** PROPERTY ******************************************
    If txtClientList.Tag <> "" Then
        szSQL = "SELECT Property.PropertyID, PropertyName " & _
               "FROM Property, GlobalData, tlbVatCode " & _
               "WHERE Property.PropertyID = GlobalData.PropertyID AND " & _
                   "GlobalData.VATRate = tlbVatCode.VAT_ID " & _
                   "AND Property.ClientID = '" & txtClientList.Tag & "' " & _
               "ORDER BY Property.PropertyID;"

           adoRst.Open szSQL, adoconn, adOpenStatic, adLockReadOnly

           If adoRst.RecordCount = 1 Then

                txtProperty.text = adoRst.Fields("PropertyName").Value
                txtProperty.Tag = adoRst.Fields("PropertyID").Value
           Else
                 txtProperty.text = ""
                 txtProperty.Tag = ""
           End If
           adoRst.Close
           Set adoRst = Nothing
    End If

  
End Sub




Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
   
End Sub

Private Sub txtDate_GotFocus()
   If Len(txtDate.text) < 10 Then txtDate.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNet.SetFocus
    End If
    TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
  If txtDate.text <> "" Then
      If TextBoxFormatDate(txtDate) Then
         
      End If
   End If

        If txtClientList.text = "" Then
            ShowMsgInTaskBar "Please select a Client.", "Y", "N"
            Exit Sub
        End If

End Sub


Private Sub txtNet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdSave
    End If
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub

Private Sub txtNet_LostFocus()
   txtNet.text = Format(Val(txtNet.text), "0.00")
End Sub



Private Sub txtSLNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdClientList
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub



Private Sub txtStatementID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtSLNumber
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
      Exit Sub
   End If
End Sub
