VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelectionProperty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Property"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6090
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
      Left            =   4590
      TabIndex        =   3
      Top             =   5400
      Width           =   1440
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
      Left            =   2790
      TabIndex        =   2
      Top             =   5400
      Width           =   1440
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProperties 
      Height          =   4845
      Left            =   135
      TabIndex        =   0
      Top             =   435
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   8546
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
      Left            =   210
      TabIndex        =   1
      Top             =   135
      Width           =   645
   End
End
Attribute VB_Name = "frmSelectionProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public szSelectedClientID As String
Private Sub cmdGDPCancel_Click()
    Unload Me
End Sub

Private Sub cmdGDPOk_Click()
     frmManagementFeeSelection.txtAssignedProperty.text = flxProperties.TextMatrix(flxProperties.row, 1)
        Unload Me
End Sub

Private Sub flxProperties_Click()
      SelectOnly1RowFlxGrid flxProperties, flxProperties.row
End Sub

Private Sub flxProperties_DblClick()
    frmManagementFeeSelection.txtAssignedProperty.text = flxProperties.TextMatrix(flxProperties.row, 1)
    Unload Me
End Sub

Private Sub flxProperties_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmManagementFeeSelection.txtAssignedProperty.text = flxProperties.TextMatrix(flxProperties.row, 1)
        Unload Me
    End If
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
Private Sub Form_Load()
    Dim szSQL As String, r As Integer
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset

'   connect to database
   adoConn.Open getConnectionString

   szSQL = "SELECT   PROPERTYID, PROPERTYNAME, ClientID " & _
           "FROM     PROPERTY where ClientID='" & szSelectedClientID & "'" & _
           "ORDER BY PROPERTYID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   r = 1
   ReDim szaProp(adoRst.RecordCount) As String
    flxProperties.Clear
    flxProperties.Cols = 4
    flxProperties.Rows = 2
    flxProperties.FormatString = "|<ID|<Property Name|"
    flxProperties.ColWidth(0) = 220
    flxProperties.ColWidth(1) = 1220
    flxProperties.ColWidth(2) = 3520
    flxProperties.ColWidth(3) = 0
   
    
    
    
   While Not adoRst.EOF
      szaProp(r) = adoRst.Fields.Item("PROPERTYID").Value & " ## " & _
                   adoRst.Fields.Item("PROPERTYNAME").Value & " ## " & _
                   adoRst.Fields.Item("ClientID").Value
                   flxProperties.AddItem ""
                   flxProperties.TextMatrix(r, 1) = adoRst.Fields.Item("PROPERTYID").Value
                   flxProperties.TextMatrix(r, 2) = adoRst.Fields.Item("PROPERTYNAME").Value
      r = r + 1
      adoRst.MoveNext
   Wend
'   iProperty = r

   adoRst.Close


   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing
     Call WheelHook(Me.hWnd)
End Sub
