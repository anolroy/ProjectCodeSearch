VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPCF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projected Income for Next term"
   ClientHeight    =   8820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPCF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGenerateReportButton 
      Caption         =   "Projected Income"
      Default         =   -1  'True
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPropertyLookup 
      Height          =   2805
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4948
      _Version        =   393216
      ForeColor       =   0
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   15329508
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   8421504
      WordWrap        =   -1  'True
      HighLight       =   2
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.CheckBox chkSelectAllRow 
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   240
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "415;450"
      Value           =   "0"
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Width           =   735
      VariousPropertyBits=   276824083
      Caption         =   "Post Code"
      Size            =   "1296;344"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   1245
      VariousPropertyBits=   276824083
      Caption         =   "Properties Name"
      Size            =   "2196;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   180
      VariousPropertyBits=   276824083
      Caption         =   "ID"
      Size            =   "317;344"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1275
      VariousPropertyBits=   276824083
      Caption         =   "Select Properties:"
      Size            =   "2249;450"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboClientList 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   75
      Width           =   4125
      VariousPropertyBits=   1753237531
      DisplayStyle    =   3
      Size            =   "7276;556"
      TextColumn      =   1
      ColumnCount     =   3
      ListRows        =   20
      cColumnInfo     =   3
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "1762;4937;1763"
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   80
      Width           =   465
   End
End
Attribute VB_Name = "frmPCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private szPropertyList As String

Private Sub cboClientList_Click()
   Dim i As Integer

   chkSelectAllRow.Value = False

   For i = 1 To flxPropertyLookup.Rows - 1
      flxPropertyLookup.RowHeight(i) = 240
   Next i

   If cboClientList.Value = "ALL" Then Exit Sub

   For i = 1 To flxPropertyLookup.Rows - 1
      If flxPropertyLookup.TextMatrix(i, 4) <> cboClientList.Value Then
         flxPropertyLookup.RowHeight(i) = 0
         flxPropertyLookup.TextMatrix(i, 0) = ""
      End If
   Next i
End Sub

Private Sub cboClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = 0
End Sub

Private Sub chkSelectAllRow_Click()
   Dim i As Integer

   For i = 1 To flxPropertyLookup.Rows - 1
      If chkSelectAllRow.Value And flxPropertyLookup.RowHeight(i) = 240 Then
         If flxPropertyLookup.TextMatrix(i, 0) = "" Then
            SelectFlxGridRow 0, flxPropertyLookup, i
         End If
      Else
         If flxPropertyLookup.TextMatrix(i, 0) = "X" Then
            SelectFlxGridRow 0, flxPropertyLookup, i
         End If
      End If
   Next i
End Sub

Private Sub flxPropertyLookup_Click()
   SelectFlxGridRow 0, flxPropertyLookup, flxPropertyLookup.row

   If flxPropertyLookup.TextMatrix(flxPropertyLookup.row, 0) = "X" Then
      chkSelectAllRow.Value = False
   End If
End Sub

Private Sub Form_Load()
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   PrepareList adoConn, cboClientList
   ConfigFlxPropertyLookup
   LoadProperties adoConn

   adoConn.Close
   Set adoConn = Nothing

   Me.Top = 0
   Me.Left = 0
   Me.Width = 5745
   Me.Height = 5040
   Me.BackColor = MODULEBACKCOLOR
   Call WheelHook(Me.hWnd)
End Sub

Public Function LoadProperties(adoConn As ADODB.Connection)
   Dim rstProperty_ As New ADODB.Recordset
   Dim iRow As Integer, szSQL As String

   szSQL = "SELECT PROPERTYID, PROPERTYNAME, ProPostCode, ClientID " & _
           "FROM   PROPERTY " & _
           "ORDER BY PropertyName;"

   rstProperty_.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   iRow = 1
   On Error Resume Next
   While Not rstProperty_.EOF
      flxPropertyLookup.TextMatrix(iRow, 1) = IIf(rstProperty_.Fields.Item(0) = Null, "", rstProperty_.Fields.Item(0))
      flxPropertyLookup.TextMatrix(iRow, 2) = IIf(rstProperty_.Fields.Item(1) = Null, "", rstProperty_.Fields.Item(1))
      flxPropertyLookup.TextMatrix(iRow, 3) = IIf(rstProperty_.Fields.Item(2) = Null, "", rstProperty_.Fields.Item(2))
      flxPropertyLookup.TextMatrix(iRow, 4) = IIf(rstProperty_.Fields.Item(3) = Null, "", rstProperty_.Fields.Item(3))

      rstProperty_.MoveNext
      If Not rstProperty_.EOF Then flxPropertyLookup.AddItem ""
      iRow = iRow + 1
   Wend

   rstProperty_.Close
   Set rstProperty_ = Nothing
End Function

Private Sub ConfigFlxPropertyLookup()
   flxPropertyLookup.Clear
   flxPropertyLookup.Rows = 2
   flxPropertyLookup.Cols = 5

   flxPropertyLookup.ColWidth(0) = Label2(0).Left - flxPropertyLookup.Left
   flxPropertyLookup.ColWidth(1) = Label2(1).Left - Label2(0).Left
   flxPropertyLookup.ColWidth(2) = Label2(2).Left - Label2(1).Left
   flxPropertyLookup.ColWidth(3) = flxPropertyLookup.Width - Label2(2).Left - 120
   flxPropertyLookup.ColWidth(4) = 0
   flxPropertyLookup.RowHeight(0) = 0
End Sub

Private Sub PrepareList(adoConn As ADODB.Connection, cboClient As Control)
   Dim adoRST As New ADODB.Recordset
   Dim szSQL As String

   On Error GoTo ErrorHandler

'*************************************** CLIENT COMBO ******************************************
   szSQL = "SELECT CLIENTID, CLIENTNAME, CLIENTPOSTCODE,  " & _
               "LandLordSageCustAC, LandLordSageSuppAC " & _
           "FROM CLIENT " & _
           "ORDER BY CLIENTNAME;"

   adoRST.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRST.EOF Then
      adoRST.Close
      Set adoRST = Nothing
      Exit Sub
   End If

   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   TotalRow = adoRST.RecordCount
   TotalCol = adoRST.Fields.Count - 1

   ReDim Data(TotalCol, TotalRow) As String

   Data(0, 0) = "ALL"
   Data(1, 0) = "All Clients"
   For i = 1 To TotalRow
       For j = 0 To TotalCol - 1
           Data(j, i) = IIf(IsNull(adoRST.Fields(j).Value), "", adoRST.Fields(j).Value)
       Next j
       adoRST.MoveNext
       If adoRST.EOF Then Exit For
   Next i

   cboClient.Column() = Data()
   cboClient.ListIndex = 0
   adoRST.Close
   Set adoRST = Nothing
   Exit Sub

ErrorHandler:
   ShowMsgInTaskBar Err.description & "::" & Err.Number, , "N"

   adoRST.Close
   Set adoRST = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call WheelUnHook(Me.hWnd)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Function CreateListOfProp() As Integer
   Dim i As Integer

   szPropertyList = ""
   
   For i = 0 To flxPropertyLookup.Rows - 1
      If flxPropertyLookup.TextMatrix(i, 0) = "X" Then
         szPropertyList = "'" & flxPropertyLookup.TextMatrix(i, 1) & "'" & ", " & szPropertyList
      End If
   Next i
   If Len(szPropertyList) > 2 Then
      szPropertyList = Left(szPropertyList, Len(szPropertyList) - 2)
      CreateListOfProp = Len(szPropertyList)
      Exit Function
   End If
   CreateListOfProp = 0
End Function

Private Sub MarkPropOfSelection(adoConn As ADODB.Connection)
   Dim szSQL As String

   szSQL = "UPDATE PROPERTY " & _
           "SET    RCC = '';"
   adoConn.Execute szSQL

   szSQL = "UPDATE PROPERTY " & _
           "SET    RCC = 'X' " & _
           "WHERE  PropertyID IN (" & szPropertyList & ");"
'Debug.Print szSQL
   adoConn.Execute szSQL
End Sub

Private Sub cmdGenerateReportButton_Click()
   If CreateListOfProp = 0 Then
      ShowMsgInTaskBar "Please select Property from the grid."
      flxPropertyLookup.SetFocus
      Exit Sub
   End If

   Dim adoConn As New ADODB.Connection
   Dim datetype As Integer
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   adoConn.Open getConnectionString

   MarkPropOfSelection adoConn

   adoConn.Close
   Set adoConn = Nothing

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\ProjectedCashFlow.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Load frmReport
   frmReport.LoadReportViewer Report
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
