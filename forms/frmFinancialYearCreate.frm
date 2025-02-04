VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFinancialYearCreate 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Year"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinancialYearCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdClosePeriod 
      Caption         =   "&Close Period"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenPeriod 
      Caption         =   "&Open Period"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Period"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   6900
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F9F9F9&
      Caption         =   "Financial Year:"
      Height          =   580
      Left            =   120
      TabIndex        =   26
      Top             =   420
      Width           =   4575
      Begin VB.TextBox txtEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   1
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox txtStDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   200
         Width           =   750
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   27
         Top             =   200
         Width           =   675
      End
   End
   Begin VB.TextBox txtGridInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEADA&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2640
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPeriods 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9F9F9&
      Caption         =   "Basis:"
      Height          =   580
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   4575
      Begin VB.OptionButton optCustome 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Custom"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   200
         Width           =   855
      End
      Begin VB.OptionButton optPreDefined 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Pre-defined"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSForms.ComboBox cboPredefined 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   195
         Width           =   1815
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "3201;556"
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
         Object.Width           =   "0"
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenrateFinancialYear 
      Caption         =   "Generate Financial Year"
      Height          =   535
      Left            =   4775
      TabIndex        =   6
      Top             =   1125
      Width           =   2685
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Financial Year"
      Height          =   375
      Left            =   -9480
      TabIndex        =   19
      Top             =   -2520
      Width           =   2055
   End
   Begin VB.TextBox txtYearDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5670
      MaxLength       =   50
      TabIndex        =   2
      Top             =   585
      Width           =   1785
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPeriods 
      Height          =   4740
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8361
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
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   495
      Left            =   3165
      Top             =   6840
      Width           =   2685
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No of Periods:"
      Height          =   195
      Index           =   4
      Left            =   4775
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblClientName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   210
      Left            =   1200
      TabIndex        =   22
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   195
      Index           =   5
      Left            =   4775
      TabIndex        =   18
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period Status"
      Height          =   195
      Index           =   4
      Left            =   5880
      TabIndex        =   17
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   16
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period Description"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
      Height          =   195
      Index           =   0
      Left            =   140
      TabIndex        =   14
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   1800
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   120
      Top             =   1800
      Width           =   7320
   End
End
Attribute VB_Name = "frmFinancialYearCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FinancialYearID  As String
Private GridIndex       As Integer
Private iDataEntryRow   As Integer
Private iDataEntryCol   As Integer
Dim szUndoText As String
Public CallingFrom As String
Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   Dim iRow    As Integer
   Dim szPd    As String
   Dim dtFySt  As Date
   Dim dtFyEd  As Date
   Dim szFy1ID As String
   Dim szFy2ID As String

   iRow = flxPeriods.Rows - 1

   If flxPeriods.TextMatrix(iRow, 0) = "" Then Exit Sub
   If flxPeriods.TextMatrix(iRow, 5) = "Close" Then
      ShowMsgInTaskBar "Closed period would not deleted", "Y", "N"
      Exit Sub
   End If
   If flxPeriods.TextMatrix(iRow, 7) = "True" Then
      ShowMsgInTaskBar "There are transations of this period or later periods", "Y", "N"
      Exit Sub
   End If

   szPd = flxPeriods.TextMatrix(iRow, 0)
   dtFySt = CDate(flxPeriods.TextMatrix(iRow, 3))

   If iRow > 1 Then
      txtEndDate.text = flxPeriods.TextMatrix(iRow - 1, 4)
      txtPeriods.text = Val(txtPeriods.text) - 1
      flxPeriods.RemoveItem iRow
      txtEndDate_LostFocus
   Else
      ConfigureFlxPeriods
      cboPredefined.Enabled = True
'      txtEndDate.text = ""
'      txtYearDesc.text = ""
   End If

   If Right(Me.Caption, 4) = "Edit" Then
      Dim conConn          As New ADODB.Connection
      Dim adoRst           As New ADODB.Recordset
      Dim adoTemp          As New ADODB.Recordset
      Dim adoPeriod        As New ADODB.Recordset
      Dim szSQL            As String

      conConn.Open getConnectionString
      szFy1ID = FinancialYearID
      If Trim(txtEndDate.text) = "" Then
           MsgBox "Please enter end date.", vbInformation, "Financial year"
           Exit Sub
      End If
   '  ***Update the Financial Year****
      szSQL = "SELECT * FROM FinancialYear " & _
              "WHERE FYrID = '" & szFy1ID & "';"
      adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

      adoRst.Fields.Item("FinancialYear").Value = txtStDate.text & "-" & txtEndDate.text
      adoRst.Fields.Item("FY_Description").Value = txtYearDesc.text
      adoRst.Fields.Item("PeriodsCount").Value = txtPeriods.text
      adoRst.Fields.Item("Freq").Value = cboPredefined.Value
      adoRst.Fields.Item("Basis").Value = IIf(optPreDefined.Value, "P", "C")
      adoRst.Fields.Item("FY_EndDate").Value = txtEndDate.text
      adoRst.Fields.Item("Status").Value = True
      adoRst.Update
      adoRst.Close

'     **** Update/Shifting Periods *****
'     ***Now shift back the next financial years (if any)

      szSQL = "SELECT * FROM FinancialYear " & _
              "WHERE FYrID > '" & szFy1ID & "' AND " & _
                    "ClientID = '" & lblClientName.Tag & "' " & _
              "ORDER BY FY_StDate;"
'Debug.Print szSQL
      adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

      If Not adoRst.EOF Then
         With adoRst.Fields
            While Not adoRst.EOF
               szFy2ID = .Item("FYrID").Value

               '1. Change the FY ID of the deleted period with the next FY id
               conConn.Execute "UPDATE Periods SET FYrID = '" & szFy2ID & "' WHERE PeriodID = '" & szPd & "';"

               '2. Get the End of the financial year. Update the FY start and end date.
               '   But FY description will not be changed.
               szSQL = "SELECT MAX(P_StDate) " & _
                       "FROM Periods " & _
                       "WHERE FYrID = '" & szFy2ID & "';"
               adoTemp.Open szSQL, conConn, adOpenStatic, adLockReadOnly
               dtFyEd = DateAdd("d", -1, adoTemp.Fields.Item(0).Value)

               .Item("FinancialYear").Value = dtFySt & "-" & dtFyEd
               .Item("FY_StDate").Value = dtFySt
               .Item("FY_EndDate").Value = dtFyEd
               adoRst.Update

               dtFySt = adoTemp.Fields.Item(0).Value 'Get the next FY start date
               adoTemp.Close

               szSQL = "SELECT PeriodID " & _
                       "FROM Periods " & _
                       "WHERE P_StDate = #" & Format(dtFySt, "DD MMMM YYYY") & "# AND FYrID = '" & szFy2ID & "';"
'Debug.Print szSQL
               adoTemp.Open szSQL, conConn, adOpenStatic, adLockReadOnly
               szPd = adoTemp.Fields.Item(0).Value 'Get the next FY start date
               adoTemp.Close

               adoRst.MoveNext
            Wend
         End With
         adoRst.Close
      End If

      conConn.Execute "DELETE * FROM Periods WHERE PeriodID = '" & szPd & "';"

      Set adoRst = Nothing
      conConn.Close
      Set conConn = Nothing
       If iRow = 0 Then
          txtEndDate.text = ""
          txtYearDesc.text = ""
          optPreDefined.Enabled = True
       End If
      ShowMsgInTaskBar "The period has been deleted", "Y", "P"
   End If
End Sub

Private Sub cmdGenrateFinancialYear_Click()
   If lblClientName.Caption = "" Then Exit Sub
   If txtStDate.text = "" Then
      ShowMsgInTaskBar "Please enter the Start Date.", "Y", "N"
      txtStDate.SetFocus
      Exit Sub
   End If
   If txtEndDate.text = "" Then
      ShowMsgInTaskBar "Please enter the End Date.", "Y", "N"
      txtEndDate.SetFocus
      Exit Sub
   End If
   If cboPredefined.text = "" Then
      ShowMsgInTaskBar "Please select the frequency of periods.", "Y", "N"
      txtPeriods.SetFocus
      Exit Sub
   End If

   Dim iRow       As Integer
   Dim bFreq      As Boolean
   Dim dtStDt     As Date
   Dim dtEndDt    As Date
   Dim szF        As String
   Dim iFF        As Integer

   If cboPredefined.Value = "WEEKLY" Then
      szF = "ww"
      iFF = 1
   End If
   If cboPredefined.Value = "MONTHLY" Then
      szF = "m"
      iFF = 1
   End If
   If cboPredefined.Value = "QTR" Then
      szF = "q"
      iFF = 1
   End If
   If cboPredefined.Value = "HY" Then
      szF = "m"
      iFF = 6
   End If

   If optPreDefined.Value Then                                 'Generate Periods
      With flxPeriods
         dtStDt = CDate(txtStDate.text)

         .Rows = 2

         iRow = 1
         bFreq = True
         While bFreq
            .TextMatrix(iRow, 0) = UniqueID()
            .TextMatrix(iRow, 1) = Format(txtStDate.text, "yy") & Format(iRow, "00")

            If szF = "ww" Then
               .TextMatrix(iRow, 2) = Format(dtStDt, "yyyy") & "Week " & CStr(iRow)
            Else
               .TextMatrix(iRow, 2) = Format(dtStDt, "mmmm") & " " & Format(dtStDt, "yyyy")
            End If

            .TextMatrix(iRow, 3) = Format(dtStDt, "dd/mm/yyyy") 'modified by anol 20190222

            dtEndDt = DateAdd(szF, iFF, dtStDt)
            dtEndDt = DateAdd("d", -1, dtEndDt)
            If dtEndDt >= CDate(txtEndDate.text) Then
               dtEndDt = CDate(txtEndDate.text)
               bFreq = False
            End If
            .TextMatrix(iRow, 4) = Format(dtEndDt, "dd/mm/yyyy") 'modified by anol 20190222
            .TextMatrix(iRow, 5) = "Open"
            .TextMatrix(iRow, 6) = FinancialYearID
            dtStDt = DateAdd("d", 1, dtEndDt)
            If bFreq Then iRow = iRow + 1
            If bFreq Then .AddItem ""
         Wend
      End With
   End If

   If optCustome.Value Then                                 'Add a Period
      With flxPeriods
         For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, 0) = "" Then Exit For
         Next iRow

         If iRow = 1 Then                 'Empty grid, :- adding period at the first line of the grid
            .TextMatrix(iRow, 0) = UniqueID()
            .TextMatrix(iRow, 1) = Format(txtStDate.text, "yy") & Format(iRow, "00")

            If szF = "ww" Then
               .TextMatrix(iRow, 2) = Format(txtStDate.text, "yyyy") & "Week " & CStr(iRow)
            Else
               .TextMatrix(iRow, 2) = Format(txtStDate.text, "mmmm") & " " & Format(txtStDate.text, "yyyy")
            End If
            .TextMatrix(iRow, 3) = txtStDate.text

            dtEndDt = DateAdd(szF, iFF, txtStDate.text)
            dtEndDt = DateAdd("d", -1, dtEndDt)
            If dtEndDt >= CDate(txtEndDate.text) Then
               dtEndDt = CDate(txtEndDate.text)
               bFreq = False
            End If
            .TextMatrix(iRow, 4) = dtEndDt
            .TextMatrix(iRow, 5) = "Open"
            .TextMatrix(iRow, 6) = FinancialYearID
            .TextMatrix(iRow, 8) = "A"
         Else                             'Adding periods at the bottom
            .AddItem ""
            dtStDt = DateAdd("d", 1, .TextMatrix(iRow - 1, 4))

            .TextMatrix(iRow, 0) = UniqueID()
            .TextMatrix(iRow, 1) = Format(txtStDate.text, "yy") & Format(iRow, "00")

            If szF = "ww" Then
               .TextMatrix(iRow, 2) = Format(dtStDt, "yyyy") & "Week " & CStr(iRow)
            Else
               .TextMatrix(iRow, 2) = Format(dtStDt, "mmmm") & " " & Format(dtStDt, "yyyy")
            End If
            .TextMatrix(iRow, 3) = dtStDt

            dtEndDt = DateAdd(szF, iFF, dtStDt)
            dtEndDt = DateAdd("d", -1, dtEndDt)
            If dtEndDt >= CDate(txtEndDate.text) Then txtEndDate.text = dtEndDt

            .TextMatrix(iRow, 4) = dtEndDt
            .TextMatrix(iRow, 5) = "Open"
            .TextMatrix(iRow, 6) = FinancialYearID
            .TextMatrix(iRow, 8) = "A"
         End If
      End With
   End If

   txtPeriods.text = iRow
End Sub

Private Sub cmdClosePeriod_Click()
   If flxPeriods.TextMatrix(flxPeriods.row, 5) = "Open" Then
      If MsgBox("Closing this period will close all prior open periods." & Chr(13) & _
                "Do you wish to continue?", vbYesNo + vbQuestion, "Close opened periods") = vbNo Then Exit Sub

      Dim conConn As New ADODB.Connection
      conConn.Open getConnectionString

      'Resolved By BOSL. Issue: 0000484
      'Modified by Asif. 26-10-2014
      CloseOpenedPeriod flxPeriods.TextMatrix(flxPeriods.row, 6), flxPeriods.TextMatrix(flxPeriods.row, 1), conConn
      'CloseOpenedPeriod lblClientName.tag, flxPeriods.TextMatrix(flxPeriods.row, 0), conConn
      
      conConn.Close
      Set conConn = Nothing
      ShowMsgInTaskBar "Periods have been closed", "Y", "P"
      ConfigureFlxPeriods
      LoadFlxPeriods
   End If
End Sub

Private Sub cmdOpenPeriod_Click()
   If flxPeriods.TextMatrix(flxPeriods.row, 5) = "Closed" Then
     
     'Resolved By BOSL. Issue: 0000484
     'Modified by Asif. 26-10-2014
      If MsgBox("Opening this period will open all following closed periods. " & Chr(13) & _
                "Do you wish to continue?", vbYesNo + vbQuestion, "Open closed periods") = vbNo Then Exit Sub

      Dim conConn As New ADODB.Connection
      conConn.Open getConnectionString

      OpenClosedPeriod flxPeriods.TextMatrix(flxPeriods.row, 6), flxPeriods.TextMatrix(flxPeriods.row, 1), conConn
      'End
      
      conConn.Close
      Set conConn = Nothing
      ShowMsgInTaskBar "Periods have been opened", "Y", "P"
      ConfigureFlxPeriods
      LoadFlxPeriods
   End If
End Sub

Private Sub cmdSave_Click()
   If txtPeriods.text = "" Then Exit Sub
   If txtEndDate.text = "" Then
        MsgBox "Please enter end date", vbInformation, "Warning"
        FocusControl txtEndDate
        Exit Sub
   End If
   If txtEndDate.text <> flxPeriods.TextMatrix(flxPeriods.Rows - 1, 4) And flxPeriods.TextMatrix(flxPeriods.Rows - 1, 4) <> "" Then
      MsgBox "The last Period End date entered must be same as the Financial Year end date.", vbCritical, "Warning"
      txtGridInput.text = txtEndDate.text
      Exit Sub
   End If

   Dim conConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim iRow  As Integer

   'On Error GoTo ErrorHandler

   conConn.Open getConnectionString

   If Right(Me.Caption, 4) = "Edit" Then GoTo Edit_Mode

'***Save Financial Year****
   szSQL = "SELECT * FROM FinancialYear " & _
           "WHERE FinancialYear = '" & txtStDate.text & "-" & txtEndDate.text & "' AND " & _
                 "ClientID = '" & frmFinancialYear.txtClientList.Tag & "';"
' Error has been fixed by anol 11-08-2023 frmFinancialYear.txtClientList.Tag was frmFinancialYear.txtClientList.Text (client Name). which was obviously wrong
                 
   adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

   If Not adoRst.EOF Then
      adoRst.Close
      Set adoRst = Nothing
      conConn.Close
      Set conConn = Nothing

      MsgBox "Financial year of these periods have already been generated.", vbInformation + vbOKOnly, "Financial Year Saving"
      Exit Sub
   Else
      adoRst.AddNew
      adoRst.Fields.Item("FYrID").Value = FinancialYearID
      adoRst.Fields.Item("ClientID").Value = lblClientName.Tag
      adoRst.Fields.Item("FinancialYear").Value = txtStDate.text & "-" & txtEndDate.text
      adoRst.Fields.Item("FY_Description").Value = txtYearDesc.text
      adoRst.Fields.Item("PeriodsCount").Value = txtPeriods.text
      adoRst.Fields.Item("Freq").Value = cboPredefined.Value
      adoRst.Fields.Item("Basis").Value = IIf(optPreDefined.Value, "P", "C")
      adoRst.Fields.Item("FY_StDate").Value = txtStDate.text
      adoRst.Fields.Item("FY_EndDate").Value = txtEndDate.text
      adoRst.Fields.Item("Status").Value = True

      adoRst.Update
      adoRst.Close
   End If

'  The following method will update (if any) existing budgets financial year
   UpdateBudgetByFinancialYear conConn

'***Save Periods*****
   szSQL = "SELECT * FROM Periods;"
   adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxPeriods.Rows - 1
      adoRst.AddNew
      adoRst.Fields.Item("PeriodID").Value = flxPeriods.TextMatrix(iRow, 0)
      adoRst.Fields.Item("FYrID").Value = FinancialYearID
      adoRst.Fields.Item("Period").Value = flxPeriods.TextMatrix(iRow, 1)
      adoRst.Fields.Item("Period_Descp").Value = flxPeriods.TextMatrix(iRow, 2)
      adoRst.Fields.Item("P_StDate").Value = flxPeriods.TextMatrix(iRow, 3)
      adoRst.Fields.Item("P_EndDate").Value = flxPeriods.TextMatrix(iRow, 4)
      adoRst.Fields.Item("Status").Value = IIf(flxPeriods.TextMatrix(iRow, 5) = "Open", True, False)

      adoRst.Update
   Next iRow
   adoRst.Close
   Set adoRst = Nothing

  
  
   
  
   'Unload Me
   MsgBox "Periods have been created successfully", vbInformation, "Saved"
   Unload Me
   If CallingFrom = "NominalLedger" Then
        frmNominalLedger.isFinancialYearCreated = True
        frmNominalLedger.LoadFirstFinancialYear
        frmNominalLedger.LoadPeriods conConn
        frmNominalLedger.NominalLedgerSetupForNewClient
        frmNominalLedger.cmdFilter.Enabled = True
   End If
   conConn.Close
   Set conConn = Nothing
   Exit Sub

Edit_Mode:
'***Save Financial Year****
   szSQL = "SELECT * FROM FinancialYear " & _
           "WHERE FYrID = '" & FinancialYearID & "';"
   adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

   adoRst.Fields.Item("FinancialYear").Value = txtStDate.text & "-" & txtEndDate.text
   adoRst.Fields.Item("FY_Description").Value = txtYearDesc.text
   adoRst.Fields.Item("PeriodsCount").Value = txtPeriods.text
   adoRst.Fields.Item("Freq").Value = cboPredefined.Value
   adoRst.Fields.Item("Basis").Value = IIf(optPreDefined.Value, "P", "C")
   adoRst.Fields.Item("FY_StDate").Value = txtStDate.text
   adoRst.Fields.Item("FY_EndDate").Value = txtEndDate.text
   adoRst.Fields.Item("Status").Value = True
   adoRst.Update
   adoRst.Close

'***Save Periods*****
   conConn.Execute "DELETE * FROM Periods WHERE FYrID = '" & FinancialYearID & "';"
   szSQL = "SELECT * FROM Periods;"
   adoRst.Open szSQL, conConn, adOpenDynamic, adLockOptimistic

   For iRow = 1 To flxPeriods.Rows - 1
      If flxPeriods.TextMatrix(iRow, 0) <> "" Then
          adoRst.AddNew
          adoRst.Fields.Item("PeriodID").Value = flxPeriods.TextMatrix(iRow, 0)
          adoRst.Fields.Item("FYrID").Value = FinancialYearID
          adoRst.Fields.Item("Period").Value = flxPeriods.TextMatrix(iRow, 1)
          adoRst.Fields.Item("Period_Descp").Value = flxPeriods.TextMatrix(iRow, 2)
          adoRst.Fields.Item("P_StDate").Value = flxPeriods.TextMatrix(iRow, 3)
          adoRst.Fields.Item("P_EndDate").Value = flxPeriods.TextMatrix(iRow, 4)
          adoRst.Fields.Item("Status").Value = IIf(flxPeriods.TextMatrix(iRow, 5) = "Open", True, False)
          adoRst.Update
      End If
   Next iRow
   adoRst.Close
   Set adoRst = Nothing

   conConn.Close
   Set conConn = Nothing

   cmdSave.Enabled = False
   ShowMsgInTaskBar "Periods have been updated successfully", "Y", "P"
   frmNominalLedger.Form_Activated = False
   Exit Sub
ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   Set adoRst = Nothing
End Sub

Private Sub UpdateBudgetByFinancialYear(conConn As ADODB.Connection)
   Dim adoRst As New ADODB.Recordset

   adoRst.Open "SELECT * FROM FinancialYear WHERE ClientID = '" & lblClientName.Tag & "'", conConn, adOpenStatic, adLockReadOnly

   If adoRst.RecordCount = 1 Then
      conConn.Execute "UPDATE Property SET CBY = '" & FinancialYearID & "' WHERE ClientID = '" & lblClientName.Tag & "';"
      conConn.Execute "UPDATE GlobalSC AS G, Property AS P " & _
                      "SET FinancialYear = '" & FinancialYearID & "' " & _
                      "WHERE P.ClientID = '" & lblClientName.Tag & "' AND " & _
                            "G.PropertyID = P.PropertyID;"
      conConn.Execute "UPDATE GlobalInsurance AS G, Property AS P " & _
                      "SET FinancialYear = '" & FinancialYearID & "' " & _
                      "WHERE P.ClientID = '" & lblClientName.Tag & "' AND " & _
                            "G.PropertyID = P.PropertyID;"
      conConn.Execute "UPDATE GlobalRC AS G, Property AS P " & _
                      "SET FinancialYear = '" & FinancialYearID & "' " & _
                      "WHERE P.ClientID = '" & lblClientName.Tag & "' AND " & _
                            "G.PropertyID = P.PropertyID;"
   End If

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub flxPeriods_DblClick()
   If Not optCustome.Value Then
      ShowMsgInTaskBar "If you wish to amend periods, please select Custom basis option", "Y", "N"
      Exit Sub
   End If
   If flxPeriods.col < 2 Or flxPeriods.col > 5 Then Exit Sub
   If flxPeriods.TextMatrix(flxPeriods.row, 0) = "" Then Exit Sub
   If flxPeriods.TextMatrix(flxPeriods.row, 5) = "Closed" Then
      ShowMsgInTaskBar "You cannot amend the closed period", "Y", "N"
      Exit Sub
   End If
   If AnyPostedTrans(flxPeriods.TextMatrix(flxPeriods.row, 3)) Then
      MsgBox "Transactions have been found on or after " & _
             Format(flxPeriods.TextMatrix(flxPeriods.row, 3), "dd/mm/yyyy") & "." & Chr(13) & _
             "You cannot amend the period", vbOKOnly & vbInformation, "Editing periods"
      Exit Sub
   End If

   txtGridInput.Top = flxPeriods.CellTop + flxPeriods.Top
   txtGridInput.Left = flxPeriods.CellLeft + flxPeriods.Left
   txtGridInput.Width = flxPeriods.ColWidth(flxPeriods.col)
   txtGridInput.Height = flxPeriods.RowHeight(flxPeriods.row) - 15
   txtGridInput.text = flxPeriods.TextMatrix(flxPeriods.row, flxPeriods.col)
   szUndoText = txtGridInput.text
   txtGridInput.Visible = True
   flxPeriods.ScrollBars = flexScrollBarNone
   txtGridInput.SetFocus
   SelTxtInCtrl txtGridInput

   iDataEntryRow = flxPeriods.row
   iDataEntryCol = flxPeriods.col
End Sub

Private Function AnyPostedTrans(dDate As Date) As Boolean
   AnyPostedTrans = False

   Dim conConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String

   conConn.Open getConnectionString

'Checking SIs & SCs
   szSQL = "SELECT Q.DemandID " & _
           "FROM LeaseDetails AS L, Units AS U, Property AS P, (" & _
                "SELECT D.DemandID, SUM(S.TotalAmount) AS TSum, D.LeaseRef " & _
                "FROM DemandRecords AS D INNER JOIN DemandSplitRecords AS S ON D.DemandID = S.DemandID " & _
                "WHERE D.IssueDate >= #" & Format(dDate, "dd mmmm yyyy") & "# " & _
                "GROUP BY D.DemandID, D.LeaseRef) AS Q " & _
            "WHERE Q.TSum > 0 AND L.Status AND Q.LeaseRef = L.LeaseID AND " & _
                  "L.UnitNumber = U.UnitNumber AND U.PropertyID = P.PropertyID AND " & _
                  "P.ClientID = '" & lblClientName.Tag & "';"
'Debug.Print szSQL
   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If
   adoRst.Close

'Checking JDs & JCs
'Resolved By BOSL. Issue: 0000484
'Modified by Asif. 26-10-2014
      
'   szSQL = "SELECT D.RecordID " & _
'           "FROM NJ_Header AS D INNER JOIN NJ_HeaderTotal AS S ON D.RecordID = S.RecordID " & _
'           "WHERE D.NJDate >= #" & Format(dDate, "dd mmmm yyyy") & "# AND " & _
'                 "D.ClientID = '" & lblClientName.tag & "' AND " & _
'                 "S.TotalAmt > 0;"

 szSQL = "SELECT D.RecordID " & _
           "FROM NJ_Header AS D " & _
           "WHERE CDATE(D.NJDate) >= #" & Format(dDate, "dd mmmm yyyy") & "# AND " & _
                 "D.ClientID = '" & lblClientName.Tag & "';"

'''''''' End of Modification
'Debug.Print szSQL
   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If
   adoRst.Close

'Checking PIs & PCs
   szSQL = "SELECT MY_ID " & _
           "FROM   tblPurInv " & _
           "WHERE  TRAN_DATE >= #" & Format(dDate, "dd mmmm yyyy") & "# AND TOTAL_AMOUNT >0 AND " & _
                  "CL_ID = '" & lblClientName.Tag & "';"

   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If
   adoRst.Close

'  Checking PPs, PAs, PPRs
   szSQL = "SELECT TransactionID " & _
           "FROM   tlbPayment " & _
           "WHERE  PDate >= #" & Format(dDate, "dd mmmm yyyy") & "# AND Amount >0 AND " & _
                  "(Type = 8 OR Type = 9 OR Type = 24) AND " & _
                  "ClientID = '" & lblClientName.Tag & "';"

   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If
   adoRst.Close

'Checking BRs & BPs
   szSQL = "SELECT MY_ID " & _
           "FROM   tlbBankPayment " & _
           "WHERE  TRAN_DATE >= #" & Format(dDate, "dd mmmm yyyy") & "# AND (NET_AMOUNT + VAT) >0 AND " & _
                  "ClientID = '" & lblClientName.Tag & "';"

   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If
   adoRst.Close

'  Checking SRs, SAs, SRRs

'Resolved By BOSL. Issue: 0000484
'Modified by Asif. 26-10-2014

'   szSQL = "SELECT R.TransactionID " & _
'           "FROM tlbReceipt AS R, Units AS U, Property AS P " & _
'           "WHERE R.RDate >= '" & dDate & "' AND R.Amount >0 AND " & _
'                 "(R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
'                 "R.UnitID = U.UnitNumber AND " & _
'                 "U.PropertyID = P.PropertyID AND " & _
'                 "P.ClientID = '" & lblClientName.tag & "';"


   szSQL = "SELECT R.TransactionID " & _
           "FROM tlbReceipt AS R, Units AS U, Property AS P " & _
           "WHERE R.RDate >= #" & Format(dDate, "dd mmmm yyyy") & "# AND R.Amount >0 AND " & _
                 "(R.Type = 3 OR R.Type = 4 OR R.Type = 23) AND " & _
                 "R.UnitID = U.UnitNumber AND " & _
                 "U.PropertyID = P.PropertyID AND " & _
                 "P.ClientID = '" & lblClientName.Tag & "';"
                 
'End of Modification

   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If Not adoRst.EOF Then
      AnyPostedTrans = True
      GoTo ExitFunction
   End If

ExitFunction:
   adoRst.Close
   Set adoRst = Nothing
   conConn.Close
   Set conConn = Nothing
End Function

Private Sub flxPeriods_KeyPress(KeyAscii As Integer)
   With flxPeriods
      If .TextMatrix(.row, 0) <> "" And KeyAscii = 13 And .col >= 2 And .col <= 4 And optCustome.Value Then
         txtGridInput.Top = .CellTop + .Top
         txtGridInput.Left = .CellLeft + .Left
         txtGridInput.Width = .ColWidth(.col)
         txtGridInput.Height = .RowHeight(.row) - 15
         txtGridInput.text = .TextMatrix(.row, .col)
         txtGridInput.Visible = True
         txtGridInput.SetFocus

         SelTxtInCtrl txtGridInput
         iDataEntryRow = .row
         iDataEntryCol = .col

         .ScrollBars = flexScrollBarNone
      End If
   End With
End Sub

Private Sub flxPeriods_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   flxPeriods.ToolTipText = flxPeriods.TextMatrix(flxPeriods.MouseRow, flxPeriods.MouseCol)
End Sub

Public Sub LoadFlxPeriods()
   Dim conConn As New ADODB.Connection
   Dim adoRst  As New ADODB.Recordset
   Dim szSQL   As String
   Dim iRow    As Integer

   On Error GoTo ErrorHandler

'  szHeader$ = "PeriodID|<Period|<Period_Descp|<P_StDate|<P_EndDate|<Status|FYrID"
   szSQL = "SELECT PeriodID, Period, Period_Descp, P_StDate, " & _
                  "P_EndDate, Status, FYrID, Transactions " & _
           "FROM Periods " & _
           "WHERE FYrID = '" & FinancialYearID & "' " & _
           "ORDER BY P_StDate;"

   conConn.Open getConnectionString
   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   iRow = 1
   While Not adoRst.EOF
      flxPeriods.TextMatrix(iRow, 0) = adoRst.Fields.Item("PeriodID").Value
      flxPeriods.TextMatrix(iRow, 1) = adoRst.Fields.Item("Period").Value
      flxPeriods.TextMatrix(iRow, 2) = adoRst.Fields.Item("Period_Descp").Value
      flxPeriods.TextMatrix(iRow, 3) = adoRst.Fields.Item("P_StDate").Value
      flxPeriods.TextMatrix(iRow, 4) = adoRst.Fields.Item("P_EndDate").Value
      flxPeriods.TextMatrix(iRow, 5) = IIf(adoRst.Fields.Item("Status").Value, "Open", "Closed")
      flxPeriods.TextMatrix(iRow, 6) = adoRst.Fields.Item("FYrID").Value
      flxPeriods.TextMatrix(iRow, 7) = adoRst.Fields.Item("Transactions").Value
      flxPeriods.TextMatrix(iRow, 8) = ""
      adoRst.MoveNext
      If Not adoRst.EOF Then flxPeriods.AddItem ""
      iRow = iRow + 1
   Wend

   adoRst.Close

   szSQL = "SELECT * " & _
           "FROM   FinancialYear " & _
           "WHERE  FYrID = '" & FinancialYearID & "';"
   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   If adoRst.Fields.Item("Basis").Value = "C" Then
      optCustome.Value = True
      cmdDelete.Enabled = True
   Else
      optPreDefined.Value = True
      cmdDelete.Enabled = False
   End If
   cboPredefined.Value = adoRst.Fields.Item("Freq").Value
   cboPredefined.Enabled = False
   txtPeriods.text = adoRst.Fields.Item("PeriodsCount").Value
   txtYearDesc.text = adoRst.Fields.Item("FY_Description").Value
   txtStDate.text = adoRst.Fields.Item("FY_StDate").Value
   txtEndDate.text = adoRst.Fields.Item("FY_EndDate").Value

   adoRst.Close
   Set adoRst = Nothing

NoRes:
   conConn.Close
   Set conConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   conConn.Close
   Set adoRst = Nothing
   Set conConn = Nothing
End Sub

Private Sub Form_Activate()
   If txtStDate.Locked Then txtEndDate.SetFocus
End Sub

Private Sub Form_Load()
   Me.Width = 7650
   Me.Height = 7875 '
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor
   Frame2.BackColor = Me.BackColor
   optPreDefined.BackColor = Me.BackColor
   optCustome.BackColor = Me.BackColor

   'lblClientName.Caption = frmFinancialYear.txtClientList.text

   ConfigureFlxPeriods
'   LoadYear
   LoadCboPredefined
   Call WheelHook(Me.hWnd)
End Sub

Private Sub ConfigureFlxPeriods()
   Dim szHeader As String, iCol As Integer

   flxPeriods.Clear
   flxPeriods.Cols = 9
   flxPeriods.Rows = 2
   flxPeriods.RowHeight(0) = 0
   szHeader$ = "PeriodID|<Period|<Period_Descp|<P_StDate|<P_EndDate|<Status|FYrID|Transactions"
   flxPeriods.FormatString = szHeader$

   flxPeriods.ColWidth(0) = 0                                    'PeriodID
   flxPeriods.ColWidth(1) = Label0(1).Left - Label0(0).Left      'Period
   flxPeriods.ColWidth(2) = Label0(2).Left - Label0(1).Left      'Period_Descp
   flxPeriods.ColWidth(3) = Label0(3).Left - Label0(2).Left      'P_StDate
   flxPeriods.ColWidth(4) = Label0(4).Left - Label0(3).Left      'P_EndDate
   flxPeriods.ColWidth(5) = flxPeriods.Width + flxPeriods.Left - Label0(4).Left - 300        'Status
   flxPeriods.ColWidth(6) = 0                                    'FYrID
   flxPeriods.ColWidth(7) = 0                                    'Transactions
   flxPeriods.ColWidth(8) = 0                                    'Modification
End Sub

Private Sub LoadCboPredefined()
   Dim conConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   Dim TotalRow As Integer, TotalCol As Integer
   Dim Data() As String
   Dim i As Integer, j As Integer

   On Error GoTo ErrorHandler

   szSQL = "SELECT S.Code, S.Value " & _
           "FROM SecondaryCode AS S " & _
           "WHERE S.PrimaryCode = 'FREQ' " & _
           "ORDER BY VAL(S.Description) DESC;"

   conConn.Open getConnectionString

   adoRst.Open szSQL, conConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   TotalRow = adoRst.RecordCount - 1
   TotalCol = adoRst.Fields.Count - 1

   cboPredefined.ColumnCount = TotalCol + 1

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow
      For j = 0 To TotalCol
         Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
      Next j
      adoRst.MoveNext
      If adoRst.EOF Then Exit For
   Next i

   cboPredefined.Column() = Data()
   'Modified by anol 07 Sep 2015
   'issue 571
   If cboPredefined.ListCount > 1 Then
        cboPredefined.ListIndex = 1
   End If
   adoRst.Close

NoRes:
   Set adoRst = Nothing
   conConn.Close
   Set conConn = Nothing
   Exit Sub

ErrorHandler:
   MsgBox ERR.description & "::" & ERR.Number

   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim iRow As Integer
   Dim szFy1ID As String

   For iRow = 1 To flxPeriods.Rows - 1
      If flxPeriods.TextMatrix(iRow, 8) = "M" Then Exit For
   Next iRow

   If iRow < flxPeriods.Rows Then
      Dim conConn          As New ADODB.Connection

      conConn.Open getConnectionString
      szFy1ID = FinancialYearID

      For iRow = 1 To flxPeriods.Rows - 1
         If flxPeriods.TextMatrix(iRow, 8) = "M" Then
            conConn.Execute "UPDATE Periods " & _
                            "SET Period_Descp = '" & flxPeriods.TextMatrix(iRow, 2) & "', " & _
                                "P_StDate = #" & Format(flxPeriods.TextMatrix(iRow, 3), "dd mmmm yyyy") & "#, " & _
                                "P_EndDate = #" & Format(flxPeriods.TextMatrix(iRow, 4), "dd mmmm yyyy") & "# " & _
                            "WHERE PeriodID = '" & flxPeriods.TextMatrix(iRow, 0) & "';"
            flxPeriods.TextMatrix(iRow, 8) = ""
         End If
      Next iRow

      conConn.Close
      Set conConn = Nothing
   End If
   If IsLoadedAndVisible("frmFinancialYear") Then
        frmFinancialYear.RefreshGrid
        frmFinancialYear.Show
   End If
   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub optCustome_Click()
   If optCustome.Value Then
      cmdGenrateFinancialYear.Caption = "Add a Period"
      cmdDelete.Enabled = True
   End If
End Sub

Private Sub optPreDefined_Click()
   If optPreDefined.Value Then
      cmdGenrateFinancialYear.Caption = "Generate Financial Year"
      cmdDelete.Enabled = False
   End If
End Sub

Private Sub txtEndDate_Change()
   TextBoxChangeDate txtEndDate
End Sub

Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl txtYearDesc
   End If
   TextBoxKeyPrsDate txtEndDate, KeyAscii
End Sub

Private Sub txtEndDate_LostFocus()
   If Not TextBoxFormatDate(txtEndDate) Then Exit Sub

   If DateDiff("d", CDate(txtStDate.text), CDate(txtEndDate.text)) < 1 Then
      MsgBox "End date cannot be earlier than start date", vbInformation, "Warning"
      txtEndDate.text = ""
      FocusControl txtEndDate
      Exit Sub
   End If

    txtYearDesc.text = txtStDate.text + "-" + txtEndDate.text
End Sub

Private Sub txtGridInput_Change()
   If iDataEntryCol = 3 Or iDataEntryCol = 4 Then TextBoxChangeDate txtGridInput
End Sub

Private Sub txtGridInput_KeyPress(KeyAscii As Integer)
   Dim i As Integer

   With flxPeriods
      If KeyAscii = 13 Then
        If iDataEntryCol = 4 And (flxPeriods.Rows - 1) = (iDataEntryRow) And txtEndDate.text <> txtGridInput.text Then
                   ShowMsgInTaskBar "The last Period End date entered must be same as the Financial Year end date.", "Y", "N"
                   txtGridInput.text = szUndoText
                   flxPeriods.TextMatrix(flxPeriods.Rows - 1, 4) = szUndoText
        Else
            .SetFocus
        End If
      Else
         If iDataEntryCol = 3 Or iDataEntryCol = 4 Then TextBoxKeyPrsDate txtGridInput, KeyAscii
      End If
   End With
End Sub

Private Sub txtGridInput_LostFocus()
   If iDataEntryCol = 3 Or iDataEntryCol = 4 Then
      If TextBoxFormatDate(txtGridInput) Then
         If iDataEntryRow = 1 And iDataEntryCol = 3 Then      'first row start date = Financial year start dt
            If txtStDate.text <> txtGridInput.text Then
               ShowMsgInTaskBar "First period start date must be same as financial year start date.", "Y", "N"
               txtGridInput.text = txtStDate.text
               Exit Sub
            End If
         End If
'         If iDataEntryRow = flxPeriods.Rows - 1 And iDataEntryCol = 4 Then    'last row end date = Financial year end dt
'            If txtStDate.text <> txtGridInput.text Then
'               ShowMsgInTaskBar "Last period end date must be same as financial year end date.", "Y", "N"
'               txtGridInput.text = txtEndDate.text
'               Exit Sub
'            End If
'         End If

         flxPeriods.TextMatrix(iDataEntryRow, iDataEntryCol) = txtGridInput.text
         If Right(Me.Caption, 4) = "Edit" Then flxPeriods.TextMatrix(iDataEntryRow, 8) = "M"

         If iDataEntryCol = 3 Then
            flxPeriods.TextMatrix(iDataEntryRow - 1, 4) = DateAdd("d", -1, flxPeriods.TextMatrix(iDataEntryRow, 3))
            If Right(Me.Caption, 4) = "Edit" Then flxPeriods.TextMatrix(iDataEntryRow - 1, 8) = "M"
         End If

         If iDataEntryCol = 4 And (flxPeriods.Rows - 1) <> (iDataEntryRow) Then
            flxPeriods.TextMatrix(iDataEntryRow + 1, 3) = DateAdd("d", 1, flxPeriods.TextMatrix(iDataEntryRow, 4))
            If Right(Me.Caption, 4) = "Edit" Then flxPeriods.TextMatrix(iDataEntryRow + 1, 8) = "M"
         End If
         If iDataEntryCol = 4 And (flxPeriods.Rows - 1) = (iDataEntryRow) Then
            If txtEndDate.text <> flxPeriods.TextMatrix(flxPeriods.Rows - 1, 4) Then
               ShowMsgInTaskBar "The last Period End date entered must be same as the Financial Year end date.", "Y", "N"
               txtGridInput.text = szUndoText
               flxPeriods.TextMatrix(flxPeriods.Rows - 1, 4) = szUndoText
               Exit Sub
            End If
         End If
      End If
   Else
      flxPeriods.TextMatrix(iDataEntryRow, iDataEntryCol) = txtGridInput.text
   End If

   txtGridInput.Visible = False
End Sub

Private Sub txtPeriods_Change()
   OnlyNumberedText txtPeriods
End Sub

Private Sub txtPeriods_KeyPress(KeyAscii As Integer)
   If KeyAscii < 48 Or KeyAscii > 57 Then Exit Sub
End Sub

Private Sub txtPeriods_KeyUp(KeyCode As Integer, Shift As Integer)
   If Val(txtPeriods.text) > 365 Then
      txtPeriods.text = Left(txtPeriods.text, Len(txtPeriods.text) - 1)

      txtPeriods.SelStart = Len(txtPeriods.text)
      txtPeriods.SelLength = Len(txtPeriods.text)
   End If
End Sub

Private Sub txtStDate_Change()
   TextBoxChangeDate txtStDate
End Sub

Private Sub txtStDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        FocusControl txtEndDate
   End If
   TextBoxKeyPrsDate txtStDate, KeyAscii
End Sub

Private Sub txtStDate_LostFocus()
   TextBoxFormatDate txtStDate
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
'          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
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

Private Sub txtYearDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FocusControl cmdGenrateFinancialYear
    End If
End Sub
