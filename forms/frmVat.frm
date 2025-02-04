VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VAT Rates"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8400
   Begin VB.CheckBox chkInUse 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdSaveVAT 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtVatDesp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtVatRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxVat 
      Height          =   3195
      Left            =   120
      TabIndex        =   10
      Top             =   675
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5636
      _Version        =   393216
      ForeColor       =   0
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   12632256
      BackColorSel    =   -2147483638
      ForeColorSel    =   0
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
      _Band(0).Cols   =   5
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InUse"
      ForeColor       =   &H80000004&
      Height          =   195
      Index           =   4
      Left            =   6030
      TabIndex        =   12
      Top             =   120
      Width           =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6600
      Y1              =   4545
      Y2              =   4545
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   615
      Left            =   120
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vat Rate"
      ForeColor       =   &H80000004&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   135
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H80000004&
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vat Code"
      ForeColor       =   &H80000004&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   135
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  ID"
      ForeColor       =   &H80000004&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   135
      Width           =   585
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmVat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  By this module user can create new bank. Basically bank contains
'  user name, branch name, bank sort code, etc.

Dim VAT_NEW_ENTRY_ As Boolean
Dim VAT_MODIFIED_ As Boolean
''

Private Sub cmdCancel_Click()    '#
   txtVatDesp.text = ""
   txtVatRate.text = ""
   chkInUse.Value = 0
   VAT_MODIFIED_ = False

   Dim adoConn As New ADODB.Connection

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ";"

   LoadGrid adoConn

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSaveVAT_Click()         '#
   Dim sSQLQuery_ As String
   Dim adoConn As New ADODB.Connection
   Dim adoRST As New ADODB.Recordset

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

   sSQLQuery_ = "SELECT VAT_ID, VAT_RATE, DESCRIPTIONS, IN_USE " & _
                "FROM tlbVatCode;"

   adoRST.Open sSQLQuery_, adoConn, adOpenDynamic, adLockOptimistic

   While Not adoRST.EOF
      If flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 2) <> adoRST.Fields.Item("VAT_RATE").Value Then
         adoRST.Fields.Item("VAT_RATE").Value = flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 2)
         adoRST.Update
      End If
      If flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 3) <> adoRST.Fields.Item("DESCRIPTIONS").Value Then
         adoRST.Fields.Item("DESCRIPTIONS").Value = flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 3)
         adoRST.Update
      End If
      If flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 4) = "YES" And Not adoRST.Fields.Item("IN_USE").Value Then
         adoRST.Fields.Item("IN_USE").Value = True
      End If
      If flxVat.TextMatrix(adoRST.Fields.Item("VAT_ID").Value + 1, 4) = "NO" And adoRST.Fields.Item("IN_USE").Value Then
         adoRST.Fields.Item("IN_USE").Value = False
      End If

      adoRST.MoveNext
   Wend

   adoRST.Close
   Set adoRST = Nothing

   LoadGrid adoConn
   adoConn.Close
   Set adoConn = Nothing
   VAT_MODIFIED_ = False

   ShowMsgInTaskBar "VAT has been updated successfully."
End Sub

Private Sub cmdUpdate_Click()    '#
   If flxVat.TextMatrix(flxVat.row, 2) <> txtVatRate.text Then VAT_MODIFIED_ = True
   If flxVat.TextMatrix(flxVat.row, 3) = txtVatDesp.text Then VAT_MODIFIED_ = True
   If flxVat.TextMatrix(flxVat.row, 4) = "YES" And chkInUse.Value = 0 Then VAT_MODIFIED_ = True
   If flxVat.TextMatrix(flxVat.row, 4) = "NO" And chkInUse.Value = 1 Then VAT_MODIFIED_ = True

   flxVat.TextMatrix(flxVat.row, 2) = txtVatRate.text
   flxVat.TextMatrix(flxVat.row, 3) = txtVatDesp.text
   flxVat.TextMatrix(flxVat.row, 4) = IIf(chkInUse.Value = 1, "YES", "NO")
End Sub

Private Sub Form_Activate()
   If Forms.Count > 2 Then
      ShowMsgInTaskBar "Please close all modules to access VAT setup.", , "N"
      Unload Me
   End If
End Sub

Private Sub Form_Load()          '#
   Dim adoConn As New ADODB.Connection
   Dim rstRec As New ADODB.Recordset

   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   VAT_MODIFIED_ = False

   adoConn.Open "DSN=" & Adsn & ";UID=;PWD=" & accessDBPws & ""

   LoadGrid adoConn

   adoConn.Close
   Set adoConn = Nothing

   Call WheelHook(Me.hWnd)
End Sub

'  System loads all Bank Records into the grid of the form
Public Function LoadGrid(adoConn As ADODB.Connection)          '#
   Dim rRow As Integer, szSQL As String
   Dim rstRec As New ADODB.Recordset

   flxVat.ColWidth(0) = Label1(1).Left - Label1(0).Left - 20
   flxVat.ColAlignment(0) = vbLeftJustify
   flxVat.ColWidth(1) = Label1(2).Left - Label1(1).Left - 20
   flxVat.ColAlignment(1) = vbLeftJustify
   flxVat.ColWidth(2) = Label1(3).Left - Label1(2).Left - 20
   flxVat.ColAlignment(2) = vbRightJustify
   flxVat.ColWidth(3) = Label1(4).Left - Label1(3).Left - 20
   flxVat.ColWidth(4) = flxVat.Width - Label1(4).Left - 120

   szSQL = "SELECT * " & _
           "FROM tlbVatCode;"
   rstRec.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If Not rstRec.EOF Then
      flxVat.Clear
      flxVat.Rows = 2

      rRow = 1
      While Not rstRec.EOF
         flxVat.TextMatrix(rRow, 0) = rstRec!VAT_ID
         flxVat.TextMatrix(rRow, 1) = rstRec!VAT_CODE
         flxVat.TextMatrix(rRow, 2) = rstRec!VAT_RATE
         flxVat.TextMatrix(rRow, 3) = rstRec!DESCRIPTIONS
         flxVat.TextMatrix(rRow, 4) = IIf(rstRec!IN_USE, "YES", "NO")
         rstRec.MoveNext
         If Not rstRec.EOF Then flxVat.AddItem ""
         rRow = rRow + 1
      Wend
   End If

   rstRec.Close

   Set rstRec = Nothing
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMMain.MousePointer = vbDefault
End Sub

Private Sub flxVat_RowColChange()            '#
   If flxVat.TextMatrix(1, 0) = "" Then Exit Sub

   txtVatRate.text = flxVat.TextMatrix(flxVat.row, 2)
   txtVatDesp.text = flxVat.TextMatrix(flxVat.row, 3)
   chkInUse.Value = IIf(flxVat.TextMatrix(flxVat.row, 4) = "YES", 1, 0)

'   cmdUpdate.Enabled = True
'   cmdSaveVAT.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)            '#
   Dim X

   If VAT_MODIFIED_ Then
      X = MsgBox("Do you want to save changes?", vbQuestion + vbYesNoCancel, "Data Saving")
      If X = vbCancel Then Cancel = 1
      If X = vbYes Then cmdSaveVAT_Click
   End If

   'If Cancel = 0 Then frmMMain.fraCmdButton.Enabled = True

   Call WheelUnHook(Me.hWnd)
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
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos

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
