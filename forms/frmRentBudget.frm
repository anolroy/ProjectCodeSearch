VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentBudget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Budget Details"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmRentBudget.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   4425
      ScaleWidth      =   5580
      TabIndex        =   29
      Top             =   1620
      Visible         =   0   'False
      Width           =   5610
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxClient 
         Height          =   3750
         Left            =   45
         TabIndex        =   15
         Top             =   675
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   6615
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
      Begin MSForms.TextBox txtSearchClientName 
         Height          =   255
         Left            =   1620
         TabIndex        =   14
         Top             =   375
         Width           =   3915
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6906;450"
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
      Begin MSForms.Label lblClientName 
         Height          =   195
         Left            =   1620
         TabIndex        =   33
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
      Begin MSForms.Label lblClientID 
         Height          =   195
         Left            =   120
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblFlxPayee 
         Caption         =   "EMPTY"
         Height          =   255
         Index           =   4
         Left            =   2115
         TabIndex        =   30
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
         Width           =   5220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5880
      Index           =   1
      Left            =   40
      TabIndex        =   17
      Top             =   840
      Width           =   10950
      Begin VB.CommandButton cmdRCFund 
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
         Left            =   3870
         TabIndex        =   37
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton cmdRCBdClose 
         Caption         =   "Cl&ose"
         Height          =   435
         Left            =   9720
         TabIndex        =   12
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRCBdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   8400
         TabIndex        =   11
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdDelete 
         Caption         =   "&Delete"
         Height          =   435
         Left            =   3960
         TabIndex        =   10
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   2760
         TabIndex        =   9
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRCBdEdit 
         Caption         =   "&Edit"
         Height          =   435
         Left            =   1440
         TabIndex        =   8
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdRCBdNew 
         Caption         =   "&New"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtRCBudgetTotal 
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
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   4815
         Width           =   2715
      End
      Begin VB.TextBox txtRCTotalArea 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4815
         Width           =   1935
      End
      Begin VB.TextBox txtTotalArea 
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
         Left            =   6960
         TabIndex        =   4
         Top             =   360
         Width           =   1905
      End
      Begin VB.TextBox txtBudget 
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
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   2745
      End
      Begin VB.TextBox txtRentChargesIDEdit 
         Height          =   285
         Left            =   12720
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtPpsf 
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1605
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxRentBudgetDetails 
         Height          =   3945
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   6959
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
      Begin MSForms.TextBox txtRCFund 
         Height          =   285
         Left            =   135
         TabIndex        =   38
         Top             =   360
         Width           =   3780
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "6667;503"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   24
         Top             =   4815
         Width           =   390
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Budget"
         Height          =   195
         Index           =   2
         Left            =   4200
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
         Left            =   8880
         TabIndex        =   20
         Top             =   120
         Width           =   945
      End
      Begin VB.Label lblRentCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Area"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   19
         Top             =   120
         Width           =   735
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
      Left            =   5040
      TabIndex        =   0
      Top             =   135
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
      Left            =   5040
      TabIndex        =   2
      Top             =   495
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
      Left            =   10170
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin MSForms.TextBox txtClientList 
      Height          =   285
      Left            =   1440
      TabIndex        =   36
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
   Begin MSForms.TextBox txtPropertyName 
      Height          =   315
      Left            =   6345
      TabIndex        =   35
      Top             =   90
      Width           =   3825
      VariousPropertyBits=   746604575
      Size            =   "6747;556"
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtBudgetYears 
      Height          =   285
      Left            =   1440
      TabIndex        =   34
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
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Property:"
      Height          =   195
      Index           =   3
      Left            =   5655
      TabIndex        =   28
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   2
      Left            =   255
      TabIndex        =   27
      Top             =   165
      Width           =   435
   End
   Begin VB.Label lblRentCharges 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budget Year:"
      Height          =   195
      Index           =   1
      Left            =   255
      TabIndex        =   26
      Top             =   540
      Width           =   930
   End
End
Attribute VB_Name = "frmRentBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Public sModule As String

'Dim Conn As New ADODB.Connection
'Dim Rst As New ADODB.Recordset
Dim SQLStr As String
Dim flgEdit As Integer
Dim flgNew As Integer
Dim flgChange As Integer
Dim isloaded As Boolean
Dim sTextBox  As String
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

Private Sub cmdRCFund_Click()
        sTextBox = "4"
    picClient.Left = 135
    picClient.Top = 1170
    picClient.Visible = True
    LoadGridFFund
    cmdProperty.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub flxClient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxClient_Click
    End If
End Sub


Private Sub LoadGridFFund()
       
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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
   lblClientID.Caption = "Fund Code"
   lblClientName.Caption = "Fund Name"

   
   txtSearchClientName.Left = 1620
   txtSearchClientName.text = ""
   txtSearchClientID.text = ""
   'txtSearchClientName.Width = 3240
   txtSearchClientID.Left = 45
   adoConn.Open getConnectionString
           
  If sModule = "RB" Then
      szSQL = "SELECT FundID,FundCode, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 1;"
   Else
      szSQL = "SELECT FundID,FundCode, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 3;"
   End If


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rstRec.EOF Then
      ShowMsgInTaskBar "Financial year has not been created.", "Y", "N"
   Else
            rRow = 1

        While Not rstRec.EOF
           flxClient.row = 1
           flxClient.RowSel = 1
           flxClient.ColSel = 1
           flxClient.TextMatrix(rRow, 0) = " " & Trim(rstRec.Fields.Item("FundID").Value)
           flxClient.TextMatrix(rRow, 1) = Trim(rstRec.Fields.Item("FundCode").Value)
           flxClient.TextMatrix(rRow, 2) = Trim(rstRec.Fields.Item("FundName").Value)
           flxClient.RowHeight(rRow) = 240
           rstRec.MoveNext
           If Not rstRec.EOF Then flxClient.AddItem ""
           rRow = rRow + 1
        Wend
   End If
   rstRec.Close
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub



Private Sub txtPpsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRCBdEdit.SetFocus
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
Private Sub LoadPropertyList()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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
   
   
   adoConn.Open getConnectionString
           
        szSQL = "SELECT PropertyID, PropertyName " & _
                    "FROM Property " & _
                    "WHERE ClientID = '" & txtClientList.Tag & "' " & _
                    "ORDER BY PropertyID;"
          
'Debug.Print szSQL
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub
Private Sub LoadflxClient()
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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
   
   adoConn.Open getConnectionString
   szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"

   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
           
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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing

End Sub
Private Sub cboBudgetYears_Click()
   Dim iRow As Integer
   For iRow = 1 To flxRentBudgetDetails.Rows - 1
      If txtBudgetYears.Tag = flxRentBudgetDetails.TextMatrix(iRow, 8) Then
         flxRentBudgetDetails.RowHeight(iRow) = 240
      Else
         flxRentBudgetDetails.RowHeight(iRow) = 0
      End If
   Next iRow
   RCSumTotal
'   If isloaded Then
'        ControlsModeRentBudgetDetails DefaultMode
'   End If
End Sub

'Private Sub cboClient_Click()
'    If txtClientList.text = "" Then Exit Sub
'
'   Dim adoConn As New ADODB.Connection
'   Dim rsClient As New ADODB.Recordset
'   Dim szSQL As String, Data() As String
'   Dim TotalRow As Integer, TotalCol As Integer
'   Dim i As Integer, j As Integer
'
'   'On Error GoTo ErrorHandler
'
'   adoConn.Open getConnectionString
'
'   szSQL = "SELECT PropertyID, PropertyName, " & _
'               "ProAddressLine1, ProPostCode " & _
'           "FROM Property " & _
'           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
'           "ORDER BY PropertyID;"
'
'   rsClient.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
'
'   If rsClient.EOF Then GoTo NoRes
'
'   TotalRow = rsClient.RecordCount - 1
'   TotalCol = rsClient.Fields.count - 1
'
'   ReDim Data(TotalCol, TotalRow) As String
'
'   For i = 0 To TotalRow
'       For j = 0 To TotalCol
'           Data(j, i) = IIf(IsNull(rsClient.Fields(j).Value), "", rsClient.Fields(j).Value)
'       Next j
'       rsClient.MoveNext
'       If rsClient.EOF Then Exit For
'   Next i
'   cboProperty.Column() = Data()
'   cboProperty.ListIndex = 0
'    'Issue 473
'   'Modified by anol 15 Sep 2014
''   If Conn.State = 0 Then
''        Conn.Open getConnectionString
''        End If
''        Dim rstSQL As New ADODB.Recordset
''        'Dim szSQL As String
''        If cboBudgetYears.ListCount = 0 Then Exit Sub
''        szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID " & _
''                "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
''                "WHERE P.PropertyID = '" & txtPropertyName.tag & "';"
''        rstSQL.Open szSQL, Conn, adOpenStatic, adLockReadOnly
''        If Not rstSQL.EOF Then
''            cboBudgetYears.text = IIf(IsNull(rstSQL.Fields.Item("CBY").Value), "", rstSQL.Fields.Item("CBY").Value)
''        Else
''            cboBudgetYears.ListIndex = -1
''        End If
''        If Conn.State = 1 Then
''            Conn.Close
''        End If
''        If rstSQL.State = 1 Then
''            rstSQL.Close
''        End If
''
'    'LoadFY
'NoRes:
'   rsClient.Close
'   adoConn.Close
'   Set rsClient = Nothing
'   Set adoConn = Nothing
'   Exit Sub

'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number
'
'   adoConn.Close
'   Set rsClient = Nothing
'   Set adoConn = Nothing
'
'End Sub

'Public Sub cboProperty_Click()

        'Resolved by BOSL
        'issue no 473
        'Modified by anol 15 Sep 2014
'        If Conn.State = 0 Then
'        Conn.Open getConnectionString
'        End If
'        Dim rstSQL As New ADODB.Recordset
'        Dim szSQL As String
'        If cboBudgetYears.ListCount = 0 Then Exit Sub
'        szSQL = "SELECT F.FinancialYear AS CBY, F.FYrID " & _
'                "FROM Property AS P LEFT JOIN FinancialYear AS F ON P.CBY = F.FYrID " & _
'                "WHERE P.PropertyID = '" & txtPropertyName.tag & "';"
'        rstSQL.Open szSQL, Conn, adOpenStatic, adLockReadOnly
'        If Not rstSQL.EOF Then
'            cboBudgetYears.text = IIf(IsNull(rstSQL.Fields.Item("CBY").Value), "", rstSQL.Fields.Item("CBY").Value)
'        Else
'            cboBudgetYears.ListIndex = -1
'        End If
'        If Conn.State = 1 Then
'            Conn.Close
'        End If
'        If rstSQL.State = 1 Then
'            rstSQL.Close
'        End If
'        If Conn.State = 0 Then
'             Conn.Open getConnectionString
'        End If


'        Call LoadFY
'        Call LoadFlxRCMain
'        Call cboBudgetYears_Click



'        If isloaded Then
'            ControlsModeRentBudgetDetails DefaultMode
'        End If
        'MsgBox flxRentBudgetDetails.Rows
       'End of modification
'End Sub

Private Sub cmdBudgetYears_Click()
    sTextBox = "3"
    picClient.Left = 915
    picClient.Top = 450
    picClient.Visible = True
    LoadGridFY
    cmdProperty.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdClientList_Click()
     sTextBox = "1"
    picClient.Left = 915
    picClient.Top = 70
    picClient.Visible = True
    LoadflxClient
    cmdProperty.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdPicCLose_Click()
    picClient.Visible = False
    Frame1(1).Enabled = True
    If cmdRCBdNew.Enabled Then
        cmdClientList.Enabled = True
        cmdProperty.Enabled = True
        cmdBudgetYears.Enabled = True
    Else
        cmdClientList.Enabled = False
        cmdProperty.Enabled = False
        cmdBudgetYears.Enabled = False

    End If
End Sub

Private Sub cmdproperty_Click()
    sTextBox = "2"
    picClient.Left = 5215
    picClient.Top = 70
    picClient.Visible = True
    LoadPropertyList
    cmdClientList.Enabled = False
    cmdBudgetYears.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdRCBdCancel_Click()
   If cmdRCBdEdit.Caption = "&Update" Then
      If MsgBox("Do you want to cancel the operation?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
         ControlsModeRentBudgetDetails DefaultMode
         Exit Sub
      End If
   End If
   
'   If MsgBox("Are you sure to cancel the changes made since the last Save?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
'      flxRentBudgetDetails.Clear
'      initialiseGrid
'      ControlsModeRentBudgetDetails DefaultMode
'      RCSumTotal
'   End If
End Sub

Private Sub cmdRCBdClose_Click()
   If flgChange = 1 Then
      If MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNo, "New Budget") = vbYes Then
         cmdRCBdSave_Click
      End If
   End If

   Me.Hide
   Unload Me
End Sub

Private Sub cmdRCBdDelete_Click()
   If MsgBox("Do you want to delete the selected " & IIf(sModule = "RB", "Rent", "Insurance") & " Charge Budget Detail?", vbQuestion + vbYesNo, "Saving") = vbNo Then
      Exit Sub
   End If
   If Trim(flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 0)) <> "" Then
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 7) = "X"
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 2) = "X"
      flxRentBudgetDetails.RowHeight(flxRentBudgetDetails.row) = 0
      deleteRC flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 0)
   Else
      flxRentBudgetDetails.RemoveItem (flxRentBudgetDetails.row)
   End If
   flgChange = 1
   'Call cmdRCBdSave_Click
   ControlsModeRentBudgetDetails DefaultMode
End Sub

Private Sub cmdRCBdEdit_Click()
   If cmdRCBdEdit.Caption = "&Edit" Then
      flgEdit = 1
      flgNew = 0
      ControlsModeRentBudgetDetails EditMode
      cmdRCFund.SetFocus
   Else
        'Modified by anol 15 Sep 2014
        'issue 473
        Dim iRow As Integer
        
        If flgNew Then
                For iRow = 1 To flxRentBudgetDetails.Rows - 1
                    If txtRCFund.Tag = flxRentBudgetDetails.TextMatrix(iRow, 2) And txtBudgetYears.Tag = flxRentBudgetDetails.TextMatrix(iRow, 8) Then
                            ShowMsgInTaskBar "This property already has a budget for this fund.", "Y", "N"
                            ControlsModeRentBudgetDetails DefaultMode
                            Exit Sub
                    End If
                Next iRow
        End If

      If txtRCFund.Tag = "" Then
         ShowMsgInTaskBar "Please select the fund."
         cmdRCFund.SetFocus
         Exit Sub
      End If
      If txtBudget.text = "" Then
         ShowMsgInTaskBar "Please enter the total budget."
         txtBudget.SetFocus
         Exit Sub
      End If
      If txtTotalArea.text = "" Then
         ShowMsgInTaskBar "Please enter the total area."
         txtTotalArea.SetFocus
         Exit Sub
      End If
      If IsNull(txtBudgetYears.Tag) Then
         ShowMsgInTaskBar "Please enter the total budget."
         txtBudget.SetFocus
         Exit Sub
      End If
      updateGrid
      flgChange = 1
      RCSumTotal
      ControlsModeRentBudgetDetails DefaultMode
   End If
End Sub

Private Sub cmdRCBdNew_Click()
   flgNew = 1
   ControlsModeRentBudgetDetails NewEntryMode
   'cmdRCFund.SetFocus
   'cboRCFund.DropDown
    sTextBox = "4"
    picClient.Left = 135
    picClient.Top = 1170
    picClient.Visible = True
    LoadGridFFund
   
    cmdClientList.Enabled = False
        cmdProperty.Enabled = False
        cmdBudgetYears.Enabled = False
    Frame1(1).Enabled = False
    txtSearchClientID.SetFocus
End Sub

Private Sub cmdRCBdSave_Click()
   Dim iRow As Integer, iRowChild As Integer
   Dim Conn As New ADODB.Connection
   For iRow = 1 To flxRentBudgetDetails.Rows - 1
      If flxRentBudgetDetails.TextMatrix(iRow, 7) = "X" Then
          'MsgBox flxRentBudgetDetails.TextMatrix(iRow, 0) & "Delete" & iRow
         deleteRC flxRentBudgetDetails.TextMatrix(iRow, 0)
      Else
         'MsgBox flxRentBudgetDetails.TextMatrix(iRow, 0) & "Add" & iRow
         saveUpdateRC iRow
      End If
   Next iRow

'  Updating the lease
   If sModule = "IB" Then
   'insurance budget
'      SQLStr = "UPDATE LInsuranceCharges AS I, GlobalInsurance AS G, Frequencies AS F, " & _
'                  "LeaseDetails AS L, Units AS U " & _
'               "SET I.TotalYearlyInsurance = G.Amount * I.ChargingFigure / 100, " & _
'                  "I.InsuranceEachPeriod = G.Amount * I.ChargingFigure / 100 / F.PartOfYear " & _
'               "WHERE I.ChargingType = 2 AND CByte(I.InsuranceDept) = G.FundType AND " & _
'                  "I.InsuranceFrequency = F.ID AND I.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber;"
'Modified by anol 15 Jun 2016
     SQLStr = "UPDATE LInsuranceCharges AS I, GlobalInsurance AS G, Frequencies AS F, " & _
                  "LeaseDetails AS L, Units AS U " & _
               "SET I.TotalYearlyInsurance = G.Amount * I.ChargingFigure / 100, " & _
                  "I.InsuranceEachPeriod = G.Amount * I.ChargingFigure / 100 / F.PartOfYear " & _
               "WHERE I.ChargingType = 2 AND CByte(I.InsuranceDept) = G.FundType AND " & _
                  "I.InsuranceFrequency = F.ID AND I.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND U.PropertyID=G.PropertyID;"
      If Conn.State = 0 Then Conn.Open getConnectionString
      Conn.Execute SQLStr

      'Charging method: Global
'      SQLStr = "UPDATE LInsuranceCharges AS I, GlobalInsurance AS G, Frequencies AS F, " & _
'                  "LeaseDetails AS L, Units AS U " & _
'               "SET I.ChargingFigure = (G.PPSF * U.TotalArea), " & _
'                  "I.TotalYearlyInsurance = G.PPSF * U.TotalArea, " & _
'                  "I.InsuranceEachPeriod = (G.PPSF * U.TotalArea) / F.PartOfYear " & _
'               "WHERE I.ChargingType = 4 AND CByte(I.InsuranceDept) = G.FundType AND " & _
'                  "I.InsuranceFrequency = F.ID AND I.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber;"
'Modified by anol 15 Jun 2016

   SQLStr = "UPDATE LInsuranceCharges AS I, GlobalInsurance AS G, Frequencies AS F, " & _
                  "LeaseDetails AS L, Units AS U " & _
               "SET I.ChargingFigure = (G.PPSF * U.TotalArea), " & _
                  "I.TotalYearlyInsurance = G.PPSF * U.TotalArea, " & _
                  "I.InsuranceEachPeriod = (G.PPSF * U.TotalArea) / F.PartOfYear " & _
               "WHERE I.ChargingType = 4 AND CByte(I.InsuranceDept) = G.FundType AND " & _
                  "I.InsuranceFrequency = F.ID AND I.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND U.PropertyID=G.PropertyID;"
      If Conn.State = 0 Then Conn.Open getConnectionString
      Conn.Execute SQLStr
   Else
    'Rent budget
      'Charging method: Global
      SQLStr = "UPDATE LRentCharges AS R, GlobalRC AS G, Frequencies AS F, LeaseDetails AS L, Units AS U, Property AS P " & _
               "SET R.spare2 = CStr(G.PPSF * U.TotalArea), " & _
                  "R.BRTotal = G.PPSF * U.TotalArea, " & _
                  "R.BRAmount = (G.PPSF * U.TotalArea) / F.PartOfYear " & _
               "WHERE U.PropertyID = P.PropertyID AND P.CBY = G.FinancialYear AND R.spare1 = '4' AND R.RentChargeDept = G.Fund AND " & _
                  "R.BRFrequency = F.ID AND R.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND U.PropertyID=P.PropertyID AND G.PropertyID = P.PropertyID;"
      If Conn.State = 0 Then Conn.Open getConnectionString
      Conn.Execute SQLStr
      
      'Charging method: Percentage
      SQLStr = "UPDATE LRentCharges AS R, GlobalRC AS G, Frequencies AS F, LeaseDetails AS L, Units AS U, Property AS P " & _
               "SET " & _
                  "R.BRTotal = (R.spare2/100) * G.TotalBudget, " & _
                  "R.BRAmount = ((R.spare2/100) * G.TotalBudget) / F.PartOfYear " & _
               "WHERE U.PropertyID = P.PropertyID AND P.CBY = G.FinancialYear AND R.spare1 = '2' AND R.RentChargeDept = G.Fund AND " & _
                  "R.BRFrequency = F.ID AND R.LeaseID = L.LeaseID AND L.UnitNumber = U.UnitNumber AND U.PropertyID=P.PropertyID AND G.PropertyID = P.PropertyID;"
      If Conn.State = 0 Then Conn.Open getConnectionString
      Conn.Execute SQLStr
   End If
'Debug.Print SQLStr
   
   

   'Conn.Execute SQLStr

'  Updating the global form display
'   If sModule = "IB" Then
'      If frmMMain.IsRibbonVersion Then
'         frmGlobalx.cmdYearlyInsurance.Caption = Format(TotalRCProperty(frmGlobalx.txtPropertyName.tag), "£0.00")
'      Else
'         frmGlobal1.cmdYearlyInsurance.Caption = Format(TotalRCProperty(frmGlobal1.txtPropertyName.tag), "£0.00")
'      End If
'   Else
'      If frmMMain.IsRibbonVersion Then
'         frmGlobalx.cmdYearlyRent.Caption = Format(TotalRCProperty(frmGlobalx.txtPropertyName.tag), "£0.00")
'      Else
'         frmGlobal1.cmdYearlyRent.Caption = Format(TotalRCProperty(frmGlobal1.txtPropertyName.tag), "£0.00")
'      End If
'   End If

   'Set Rst = Nothing
   If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
   End If

   flgEdit = 0
   flgNew = 0
   
   flgChange = 0
   ControlsModeRentBudgetDetails SavedMode
End Sub

Private Function saveUpdateRC(ByVal Index As Integer)
    'MsgBox "Index"
   Dim Rst As New ADODB.Recordset
   Dim Conn As New ADODB.Connection
   If sModule = "RB" Then
      SQLStr = "DELETE * " & _
               "FROM GlobalRC " & _
               "WHERE BudgetID = '" & flxRentBudgetDetails.TextMatrix(Index, 0) & "' "
   Else
      SQLStr = "DELETE * " & _
               "FROM GlobalInsurance " & _
               "WHERE ID = " & IIf(flxRentBudgetDetails.TextMatrix(Index, 0) = "", 0, flxRentBudgetDetails.TextMatrix(Index, 0)) & " or Amount=0 "
   End If
   If Conn.State = 0 Then
        Conn.Open getConnectionString
   End If
   Conn.Execute SQLStr

   If sModule = "RB" Then
      SQLStr = "SELECT * " & _
               "FROM GlobalRC;"
   Else
        SQLStr = "SELECT I.ID, I.PropertyID, I.FundType AS Fund, " & _
                   "I.Amount AS TotalBudget, I.PPSF, I.SCArea, I.FinancialYear " & _
                "FROM GlobalInsurance AS I;"
   End If
   
   
   Dim Rst1 As New ADODB.Recordset
   Rst1.Open SQLStr, Conn, adOpenDynamic, adLockOptimistic
'   If isDuplicate(conn, Val(flxRentBudgetDetails.TextMatrix(Index, 2)), txtPropertyName.tag, flxRentBudgetDetails.TextMatrix(Index, 8)) = False Then
      Rst1.AddNew
'   End If
   If sModule = "RB" Then
      Rst1!budgetId = UniqueID()
      flxRentBudgetDetails.TextMatrix(Index, 0) = Rst1!budgetId
   Else
      Rst1!ID = SlNumber("IB", "GlobalInsurance", Conn)
      flxRentBudgetDetails.TextMatrix(Index, 0) = Rst1!ID
   End If
'   If frmMMain.IsRibbonVersion Then
    Rst1!propertyID = txtPropertyName.Tag
'   Else
'      Rst1!PropertyID = frmGlobal1.txtPropertyName.tag
'   End If
   
   Rst1!Fund = Val(flxRentBudgetDetails.TextMatrix(Index, 2))
   Rst1!TotalBudget = Val(flxRentBudgetDetails.TextMatrix(Index, 4))
   Rst1!SCArea = Val(flxRentBudgetDetails.TextMatrix(Index, 5))
   Rst1!ppsf = Val(flxRentBudgetDetails.TextMatrix(Index, 6))
   Rst1!FinancialYear = flxRentBudgetDetails.TextMatrix(Index, 8)
    'MsgBox Val(flxRentBudgetDetails.TextMatrix(Index, 4))
   Rst1.Update
   Rst1.Close
   If Conn.State = 1 Then
     Conn.Close
   End If
End Function

Private Function isDuplicate(Conn As Connection, strfund As String, strPropertyID As String, StrFinancialYear) As Boolean
      Dim rsCheck As New ADODB.Recordset
       rsCheck.Open "SELECT I.ID, I.PropertyID, I.FundType AS Fund, " & _
                   "I.Amount AS TotalBudget, I.PPSF, I.SCArea, I.FinancialYear " & _
                "FROM GlobalInsurance AS I Where FundType='" & strfund & "' and FinancialYear='" & StrFinancialYear & "' AND PropertyID='" & strPropertyID & "';", Conn, adOpenStatic, adLockReadOnly
'                If strPropertyID = "BR01" Then
'                  MsgBox "Hi"
'                End If
                
      If rsCheck.EOF = True Then
         isDuplicate = False
      Else
         isDuplicate = True
      End If
End Function
Private Function deleteRC(ByVal bId As String)
 Dim Conn As New ADODB.Connection
   If sModule = "RB" Then
      SQLStr = "DELETE * " & _
               "FROM GlobalRC " & _
               "WHERE BudgetID = '" & bId & "' "
   Else
      SQLStr = "DELETE * " & _
               "FROM GlobalInsurance " & _
               "WHERE ID = " & bId & ";"
   End If
   If Conn.State = 0 Then
        Conn.Open getConnectionString
   End If
   Conn.Execute SQLStr
   If Conn.State = 1 Then
      Conn.Close
      Set Conn = Nothing
   End If
End Function
'Private Sub cmdRCClose_Click()
'
' initialiseGrid
' Me.Hide
'
'
'End Sub

Private Sub ConfigureFlxBRMain()
   Dim szFlxHeader As String
   flxRentBudgetDetails.Rows = 1
   flxRentBudgetDetails.RowHeight(0) = 0
   flxRentBudgetDetails.Clear
   flxRentBudgetDetails.Cols = 9
   szFlxHeader$ = "BudgetID|PropertyID|<Fund|>FundName|>TotalBudget|>SCArea|>PPSF|FY"
   flxRentBudgetDetails.FormatString = szFlxHeader$

   flxRentBudgetDetails.ColWidth(0) = 0
   flxRentBudgetDetails.ColWidth(1) = 0
   flxRentBudgetDetails.ColWidth(2) = 0
   flxRentBudgetDetails.ColWidth(3) = lblRentCharges(2).Left - lblRentCharges(0).Left
   flxRentBudgetDetails.ColWidth(4) = lblRentCharges(3).Left - lblRentCharges(2).Left
   flxRentBudgetDetails.ColWidth(5) = lblRentCharges(4).Left - lblRentCharges(3).Left
   flxRentBudgetDetails.ColWidth(6) = flxRentBudgetDetails.Width - lblRentCharges(4).Left - 300
   flxRentBudgetDetails.ColWidth(7) = 0
   flxRentBudgetDetails.ColWidth(8) = 0
End Sub

Private Sub LoadFlxRCMain(Conn As ADODB.Connection)
       flxRentBudgetDetails.Clear
       Call ConfigureFlxBRMain
       Dim i As Integer
      
       Dim Rst As New ADODB.Recordset
      
    
       If sModule = "RB" Then
            'rent budget
          SQLStr = "SELECT g.*, f.FundName " & _
                   "FROM GlobalRC AS g, Fund AS f " & _
                   "WHERE CInt(g.Fund) = f.FundId AND " & _
                         "PropertyID = '" & txtPropertyName.Tag & "';"
       Else
            'insurence budget
          SQLStr = "SELECT I.ID AS BudgetID, I.FundType AS Fund, I.Amount AS TotalBudget, " & _
                      "f.FundName, I.PropertyID, I.SCArea, I.PPSF, I.FinancialYear " & _
                   "FROM GlobalInsurance I, Fund AS f " & _
                   "WHERE I.PropertyID = '" & txtPropertyName.Tag & "' AND " & _
                         "I.FundType = f.FundID;"
       End If
       Debug.Print SQLStr
       Rst.Open SQLStr, Conn, adOpenStatic, adLockReadOnly
       i = 1
       If Not Rst.EOF Then
          While Not Rst.EOF
             flxRentBudgetDetails.AddItem ""
             flxRentBudgetDetails.TextMatrix(i, 0) = Rst!budgetId
             flxRentBudgetDetails.TextMatrix(i, 1) = Rst!propertyID
             flxRentBudgetDetails.TextMatrix(i, 2) = Rst!Fund
             flxRentBudgetDetails.TextMatrix(i, 3) = Rst!FundName
             flxRentBudgetDetails.TextMatrix(i, 4) = Format(Rst!TotalBudget, "0.00")
             flxRentBudgetDetails.TextMatrix(i, 5) = IIf(IsNull(Rst!SCArea), "", Rst!SCArea)
             flxRentBudgetDetails.TextMatrix(i, 6) = IIf(IsNull(Rst!ppsf), "", Rst!ppsf)
             flxRentBudgetDetails.TextMatrix(i, 8) = IIf(IsNull(Rst!FinancialYear), "", Rst!FinancialYear)
             flxRentBudgetDetails.RowHeight(i) = 0
             i = i + 1
             Rst.MoveNext
          Wend
       End If
       flxRentBudgetDetails.row = 0
       flxRentBudgetDetails.col = 0
    
       Rst.Close
       Set Rst = Nothing
       
End Sub

Private Sub RCSumTotal()
   Dim iRow As Integer

   txtRCBudgetTotal.text = "0.00"
   txtRCTotalArea.text = "0"
   For iRow = 1 To flxRentBudgetDetails.Rows - 1
      If flxRentBudgetDetails.RowHeight(iRow) > 0 Then
         txtRCBudgetTotal.text = Format(Val(txtRCBudgetTotal.text) + Val(flxRentBudgetDetails.TextMatrix(iRow, 4)), "0.00")
         txtRCTotalArea.text = Val(txtRCTotalArea.text) + Val(flxRentBudgetDetails.TextMatrix(iRow, 5))
      End If
   Next iRow
End Sub

Private Sub flxRentBudgetDetails_RowColChange()
   If flxRentBudgetDetails.TextMatrix(1, 2) = "" Then Exit Sub

   'On Error Resume Next

   txtRCFund.Tag = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 2)
   txtRCFund.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 3)
   txtBudget.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 4)
   txtTotalArea.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 5)
   txtPpsf.text = flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 6)
   

   ControlsModeRentBudgetDetails GridRowOnSelection
   cmdRCBdEdit.SetFocus
End Sub

Private Sub ControlsModeRentBudgetDetails(ByVal mode As ComponentMode)
   Select Case mode
      Case ComponentMode.DefaultMode
         'Resolved by BOSL
         'issue 473 note 667
         'Modified by anol 19 Oct 2014
         If flgChange = 1 Then
            cmdClientList.Enabled = False
            cmdProperty.Enabled = False
            cmdBudgetYears.Enabled = False
         Else
            cmdClientList.Enabled = True
            cmdProperty.Enabled = True
            cmdBudgetYears.Enabled = True
         End If
         'End of modification
         
         txtRCFund.text = ""
         txtRCFund.Tag = ""
         cmdRCFund.Enabled = False
         txtBudget.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True

         cmdRCBdNew.Enabled = True
         cmdRCBdEdit.Enabled = False
         'cboBudgetYears.Enabled = True
         cmdRCBdEdit.Caption = "&Edit"
         If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
            cmdRCBdCancel.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
            cmdRCBdCancel.Enabled = False
         End If
         
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = True
         flxRentBudgetDetails.row = 0
         flxRentBudgetDetails.col = 0
         Me.Show
         cmdRCBdNew.SetFocus
         
      Case ComponentMode.SavedMode
         'Resolved by BOSL
         'issue 473 note 667
         'Modified by anol 19 Oct 2014
         cmdClientList.Enabled = True
         cmdProperty.Enabled = True
         cmdBudgetYears.Enabled = True
         'End of modification
         txtRCFund.text = ""
         txtRCFund.Tag = ""
         cmdRCFund.Enabled = False
         txtBudget.text = ""
         txtBudget.Locked = True
         txtTotalArea.text = ""
         txtTotalArea.Locked = True
         txtPpsf.text = ""
         txtPpsf.Locked = True
         

         cmdRCBdNew.Enabled = True
         cmdRCBdEdit.Enabled = False
         'cboBudgetYears.Enabled = True
         cmdRCBdEdit.Caption = "&Edit"
         If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
            cmdRCBdCancel.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
            cmdRCBdCancel.Enabled = False
         End If
         
         cmdRCBdDelete.Enabled = False
         
         flxRentBudgetDetails.Enabled = True
         flxRentBudgetDetails.row = 0
         flxRentBudgetDetails.col = 0
         frmRentBudget.Show
         
         cmdRCBdClose.SetFocus

      Case ComponentMode.EditMode
         'Resolved by BOSL
         'issue 473 note 667
         'Modified by anol 19 Oct 2014
         cmdClientList.Enabled = False
         cmdProperty.Enabled = False
         cmdBudgetYears.Enabled = False
         'End of modification
         
         cmdRCFund.Enabled = True
         txtBudget.Locked = False
         txtTotalArea.Locked = False
         
         cmdRCBdNew.Enabled = False
         cmdRCBdEdit.Caption = "&Update"
         cmdRCBdEdit.Enabled = True
         cmdRCBdSave.Enabled = False
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = False

      Case ComponentMode.NewEntryMode
         'Reolved by BOSL
         'issue 473 note 667
         'Modified by anol 19 Oct 2014
         cmdClientList.Enabled = False
         cmdProperty.Enabled = False
         cmdBudgetYears.Enabled = False
         'End of modification
         txtRCFund.text = ""
         txtRCFund.Tag = ""
         cmdRCFund.Enabled = True
         txtBudget.text = ""
         txtBudget.Locked = False
         txtTotalArea.text = ""
         txtTotalArea.Locked = False
         txtPpsf.text = ""
         
         cmdRCBdNew.Enabled = False
         cmdRCBdEdit.Caption = "&Update"
         cmdRCBdEdit.Enabled = True
         'Resolved by BOSL
         'Modified by Anol 15 Sep 2014
         'issue 473
         ''cboBudgetYears.Enabled = False
         'End of modification
         cmdRCBdSave.Enabled = False
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = False

         flxRentBudgetDetails.Enabled = False

      Case ComponentMode.GridRowOnSelection
         cmdRCBdNew.Enabled = True
         'cboBudgetYears.Enabled = False
         cmdRCBdEdit.Enabled = True
        If flgChange = 1 Then
            cmdRCBdSave.Enabled = True
         Else
            cmdRCBdSave.Enabled = False
         End If
         cmdRCBdCancel.Enabled = True
         cmdRCBdDelete.Enabled = True
   End Select
End Sub
Private Sub LoadFY(adoConn As ADODB.Connection)
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
        rstSQL.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
Private Sub cboProperty_Change()
    'Created by Anol 29 Oct 2014
    'issue 471 note 724
    If isloaded = False Then Exit Sub
    Dim strCheck As String
    Dim rsSQL As New ADODB.Recordset
    Dim adoConn1 As New ADODB.Connection
    If adoConn1.State = 0 Then
        adoConn1.Open getConnectionString
    End If
    Dim szSQL  As String
    szSQL = "SELECT P.CBY " & _
        "FROM Property P " & _
        "WHERE P.PropertyID = '" & txtPropertyName.Tag & "';"
    rsSQL.Open szSQL, adoConn1, adOpenStatic, adLockReadOnly
    If Not rsSQL.EOF Then
        strCheck = IIf(IsNull(rsSQL.Fields.Item("CBY").Value), "", rsSQL.Fields.Item("CBY").Value)
        txtBudgetYears.Tag = strCheck
    End If
    If strCheck = "" Then
        MsgBox "" & IIf(sModule = "RB", "A Rent", "An Insurance") & " charge budget year has not been set for this property." & (Chr(13) + Chr(10)) & " Please set " & IIf(sModule = "RB", "a Rent", "an Insurance") & " charge budget year in the global data screen", vbInformation, "Warning!!"
    End If
    If adoConn1.State = 1 Then
        adoConn1.Close
    End If
End Sub
Private Sub flxClient_Click()
        Frame1(1).Enabled = True
        If cmdRCBdNew.Enabled Then
            cmdClientList.Enabled = True
            cmdProperty.Enabled = True
            cmdBudgetYears.Enabled = True
        Else
            cmdClientList.Enabled = False
            cmdProperty.Enabled = False
            cmdBudgetYears.Enabled = False
        End If
        Dim adoConn As New ADODB.Connection
        adoConn.Open getConnectionString
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
                adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
                If Not adoRst.EOF Then
                        txtPropertyName.text = adoRst.Fields(1).Value
                        txtPropertyName.Tag = adoRst.Fields(0).Value
                        LoadFlxRCMain adoConn
                        Call cboProperty_Change
                        LoadFY adoConn
                        cboBudgetYears_Click
                Else
                        txtPropertyName.text = ""
                        txtPropertyName.Tag = ""
                End If
                
                cmdProperty.SetFocus
                
        End If
        If sTextBox = "2" Then
               
                txtPropertyName.text = flxClient.TextMatrix(flxClient.row, 2)
                txtPropertyName.Tag = flxClient.TextMatrix(flxClient.row, 1)
                LoadFlxRCMain adoConn
                Call cboProperty_Change
                LoadFY adoConn
                cboBudgetYears_Click
                cmdBudgetYears.SetFocus
        End If
        If sTextBox = "3" Then
                txtBudgetYears.text = flxClient.TextMatrix(flxClient.row, 1)
                txtBudgetYears.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                cboBudgetYears_Click
                If cmdRCBdNew.Enabled Then cmdRCBdNew.SetFocus
        End If
        If sTextBox = "4" Then
                txtRCFund.text = flxClient.TextMatrix(flxClient.row, 2)
                txtRCFund.Tag = Trim(flxClient.TextMatrix(flxClient.row, 0))
                If txtBudget.Enabled Then txtBudget.SetFocus
        End If
        adoConn.Close
       
        picClient.Visible = False
End Sub
Private Sub Form_Load()
    On Error GoTo ERR
    Dim errmsg As String
    Me.Width = 11145
    Me.Height = 7215
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
    Me.BackColor = MODULEBACKCOLOR
    Frame1(1).BackColor = Me.BackColor
    
    Call ConfigureFlxBRMain
    flgEdit = 0
    flgChange = 0
    flgNew = 0
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    LoadCmbClient adoConn 'LOADING CLIENT , PROPERTY AND BudgetYears
    LoadFlxRCMain adoConn
    LoadFund adoConn
    LoadFY adoConn
    adoConn.Close
    Set adoConn = Nothing
    ControlsModeRentBudgetDetails DefaultMode
    RCSumTotal
    cboBudgetYears_Click
    Call WheelHook(Me.hWnd)
    isloaded = True
    '###############
    
   
    Exit Sub
ERR:
   MsgBox ERR.description & "-From the form load" & errmsg
End Sub
Private Sub LoadCmbClient(adoConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   szSQL = "SELECT CLIENTID, CLIENTNAME " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

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
                adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
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
'Private Sub initialiseGrid()
'   Call ConfigureFlxBRMain
'   flgEdit = 0
'   flgChange = 0
'   flgNew = 0
'
'  ' PrepareList Conn, cboClient, cboProperty
'
'   Call LoadFlxRCMain
'   Call LoadFund
'   Call LoadFY
'
''   Conn.Close
''   Set Conn = Nothing
'   ControlsModeRentBudgetDetails DefaultMode
'End Sub
Private Sub LoadGridFY()
   
   Dim rRow As Integer
   Dim szSQL As String

   Dim adoConn As New ADODB.Connection
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
   adoConn.Open getConnectionString
           
        szSQL = "SELECT FYrID, FinancialYear, FY_Description " & _
           "FROM   FinancialYear AS F, Property AS P " & _
           "WHERE  F.ClientID = P.ClientID AND " & _
                  "P.PropertyID = '" & txtPropertyName.Tag & "' order by FinancialYear Desc ;"


   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

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
   adoConn.Close
   Set rstRec = Nothing
   Set adoConn = Nothing
End Sub



Private Sub LoadFund(Conn As ADODB.Connection)
   ' Error Handler
   'On Error GoTo Error_Handler
   
   Dim rRow As Integer, iRec As Integer, Data() As String
   Dim szSQL As String, adoRst As New ADODB.Recordset

   If sModule = "RB" Then
      szSQL = "SELECT FundID, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 1;"
   Else
      szSQL = "SELECT FundID, FundName " & _
              "FROM Fund " & _
              "WHERE CategoryCode = 3;"
   End If

   adoRst.Open szSQL, Conn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      ShowMsgInTaskBar "An insurance fund has not been setup for this property", , "N"
   Else
'      ReDim Data(2, adoRst.RecordCount) As String
'
'      rRow = 0
'      While Not adoRst.EOF
'         Data(0, rRow) = Trim(adoRst.Fields.Item("FundID").Value)
'         Data(1, rRow) = Trim(adoRst.Fields.Item("FundName").Value)
'         rRow = rRow + 1
'         adoRst.MoveNext
'      Wend
'
'      cboRCFund.Clear
'      cboRCFund.Column() = Data()
        txtRCFund.text = Trim(adoRst.Fields.Item("FundID").Value)
        txtRCFund.Tag = Trim(adoRst.Fields.Item("FundName").Value)
   End If

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
'Error_Handler:
'
'   ShowMsgInTaskBar "Error in Loading fund.", , "N"
   ' Destroy Objects
  ' Set adoRst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub txtBudget_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTotalArea.SetFocus
    End If
   DigitTextKeyPress txtBudget, KeyAscii
End Sub

Private Sub txtBudget_LostFocus()
   computePpsf
End Sub

Private Sub computePpsf()
 If Trim(txtBudget.text) <> "" And Trim(txtTotalArea.text) <> "" Then
      Dim ppsf As Double
      ppsf = CDbl(txtBudget.text) / CDbl(Trim(txtTotalArea.text))
      txtPpsf.text = FormatNumber(CDbl(ppsf), 2, , , vbDefault)
   End If
End Sub

Private Sub txtTotalArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPpsf.SetFocus
    End If
   DigitTextKeyPress txtTotalArea, KeyAscii
End Sub

Private Sub txtTotalArea_LostFocus()
   computePpsf
End Sub

Private Sub updateGrid()
   If flgEdit = 0 Then
      flxRentBudgetDetails.AddItem ""
      If flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 0) <> "" Then flxRentBudgetDetails.AddItem ""
'      If frmMMain.IsRibbonVersion Then
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 1) = txtPropertyName.Tag
'      Else
'      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 1) = txtPropertyName.tag
'      End If
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 2) = CInt(txtRCFund.Tag)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 3) = txtRCFund.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 5) = txtTotalArea.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 6) = txtPpsf.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.Rows - 1, 8) = txtBudgetYears.Tag
   Else
'      If frmMMain.IsRibbonVersion Then
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 1) = txtPropertyName.Tag
'      Else
'      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 1) = txtPropertyName.tag
'      End If
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 2) = CInt(txtRCFund.Tag)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 3) = txtRCFund.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 4) = FormatNumber(CDbl(Trim(txtBudget.text)), 2, , , vbDefault)
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 5) = txtTotalArea.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 6) = txtPpsf.text
      flxRentBudgetDetails.TextMatrix(flxRentBudgetDetails.row, 8) = txtBudgetYears.Tag
   End If

   flgEdit = 0
   flgNew = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If frmMMain.IsRibbonVersion Then
   frmGlobalx.Enabled = True
   Else
'   frmGlobal1.Enabled = True
   End If
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control, cboProperty As Control)
   Dim rsClient1 As New ADODB.Recordset
   Dim szSQL As String
    Dim Conn As New ADODB.Connection
   'On Error GoTo ErrorHandler
   If Conn.State = 0 Then
        Conn.Open getConnectionString
   End If
'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTID;"

   rsClient1.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rsClient1.EOF Then
        rsClient1.Close
        Set rsClient1 = Nothing
        Exit Sub
    End If
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = rsClient1.RecordCount - 1
   TotalCol = rsClient1.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(rsClient1.Fields(j).Value), "", rsClient1.Fields(j).Value)
       Next j
       rsClient1.MoveNext
       If rsClient1.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   If rsClient1.State = 1 Then
        rsClient1.Close
        Set rsClient1 = Nothing
   End If
   
'*************************************** PROPERTY ******************************************
   If Conn.State = 0 Then
        Conn.Open getConnectionString
   End If
   Dim rsProperty As New ADODB.Recordset
   szSQL = "SELECT PropertyID, PropertyName, " & _
               "ProAddressLine1, ProPostCode " & _
           "FROM Property " & _
           "WHERE ClientID = '" & txtClientList.Tag & "' " & _
           "ORDER BY PropertyID;"

   rsProperty.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If rsProperty.EOF Then
        rsProperty.Close
        Set rsProperty = Nothing
        Exit Sub
    End If

   TotalRow = rsProperty.RecordCount - 1
   TotalCol = rsProperty.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(rsProperty.Fields(j).Value), "", rsProperty.Fields(j).Value)
      Next j
      rsProperty.MoveNext
      If rsProperty.EOF Then Exit For
   Next i
   cboProperty.Column() = Data()
   cboProperty.ListIndex = 0
   If rsProperty.State = 1 Then
        rsProperty.Close
            Set rsProperty = Nothing
   End If
   If Conn.State = 1 Then
        Conn.Close
    End If
'NoRes:
'    Resolved by BOSL
'   Modified by Anol 15 Sep 2014
'   Issue number 473
'       If adoRst.State = 1 Then
'            adoRst.Close
'            Set adoRst = Nothing
'       End If

  ' Exit Sub

'ErrorHandler:
'   MsgBox ERR.description & "::" & ERR.Number

   'Resolved by BOSL
   'Modified by Anol 15 Sep 2014
   'Issue number 473
'       If adoRst.State = 1 Then
'            adoRst.Close
'            Set adoRst = Nothing
'       End If
End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================







