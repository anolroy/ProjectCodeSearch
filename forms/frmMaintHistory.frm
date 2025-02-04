VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintHistory 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Maintanance History"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   Icon            =   "frmMaintHistory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   11295
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmMaintHistory.frx":0442
      Left            =   1640
      List            =   "frmMaintHistory.frx":044F
      TabIndex        =   21
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add:"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   4080
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "&Payment"
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Credit Note"
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Invoice"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00AAAAAF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5760
      TabIndex        =   12
      Top             =   4080
      Width           =   5415
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add Notes"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCost 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10320
      TabIndex        =   11
      Top             =   3600
      Width           =   800
   End
   Begin VB.TextBox txtPerson 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9150
      TabIndex        =   10
      Top             =   3600
      Width           =   1140
   End
   Begin VB.TextBox txtCompDate 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7620
      TabIndex        =   9
      Top             =   3600
      Width           =   1520
   End
   Begin VB.TextBox txtDes 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3140
      TabIndex        =   8
      Top             =   3600
      Width           =   4480
   End
   Begin VB.TextBox txtType 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   1480
   End
   Begin VB.TextBox txtType1 
      Height          =   375
      Left            =   16603
      TabIndex        =   6
      Top             =   3600
      Width           =   1480
   End
   Begin VB.TextBox txtRepDate 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   1480
   End
   Begin VB.TextBox txtInsuRwnewalDt 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ListView lsvUnit 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblUnitID 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Renewal Date:"
      Height          =   195
      Index           =   1
      Left            =   7080
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMaintHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbConn As New ADODB.Connection
Dim rstRecordSet As New ADODB.Recordset
Dim SQLStr As String
Dim inListView As Integer

Public unitid As String

Private Sub EnableTextBoxes()
    txtRepDate.Enabled = True
'    txtType.Enabled = True
    cmbType.Enabled = True
    txtDes.Enabled = True
    txtCompDate.Enabled = True
    txtPerson.Enabled = True
    txtCost.Enabled = True
    txtRepDate.text = ""
'    txtType.Text = ""
    cmbType.text = ""
    txtDes.text = ""
    txtCompDate.text = ""
    txtPerson.text = ""
    txtCost.text = ""
End Sub

Private Sub DisableTextBoxes()
    txtRepDate.Enabled = False
'    txtType.Enabled = False
    cmbType.Enabled = False
    txtDes.Enabled = False
    txtCompDate.Enabled = False
    txtPerson.Enabled = False
    txtCost.Enabled = False
    txtRepDate.text = ""
'    txtType.Text = ""
    cmbType.text = ""
    txtDes.text = ""
    txtCompDate.text = ""
    txtPerson.text = ""
    txtCost.text = ""
End Sub

Private Sub cmdAddNew_Click()
    EnableTextBoxes
    
    cmdAddNew.Enabled = False
    cmdEdit.Enabled = False
    cmdUpdate.Enabled = True
    txtRepDate.text = Format(Now, "dd/mm/yyyy")
    txtRepDate.SetFocus
    SelTxtInCtrl txtRepDate
End Sub

Private Sub cmdClose_Click()
    Unload Me
'    Load frmUnit
'    frmUnit.Show
End Sub

Private Sub cmdEdit_Click()
    cmdUpdate.Enabled = True
    EnableTextBoxes
    inListView = lsvUnit.SelectedItem.Index

    txtRepDate.text = lsvUnit.ListItems(inListView)
'    txtType.Text = lsvUnit.SelectedItem.SubItems(1)
    cmbType.text = lsvUnit.SelectedItem.SubItems(1)
    txtDes.text = lsvUnit.SelectedItem.SubItems(2)
    txtCompDate.text = lsvUnit.SelectedItem.SubItems(3)
    txtPerson.text = lsvUnit.SelectedItem.SubItems(4)
    txtCost.text = lsvUnit.SelectedItem.SubItems(5)
    txtRepDate.SetFocus
End Sub

Private Sub cmdUpdate_Click()
    If cmdAddNew.Enabled = False Then
'        MsgBox lsvUnit.ListItems.Count
        Dim iLstVw As ListItem
        Set iLstVw = lsvUnit.ListItems.Add(lsvUnit.ListItems.count + 1, , txtRepDate.text)
'        iLstVw.SubItems(1) = txtType.Text
        iLstVw.SubItems(1) = cmbType.text
        iLstVw.SubItems(2) = txtDes.text
        iLstVw.SubItems(3) = txtCompDate.text
        iLstVw.SubItems(4) = txtPerson.text
        iLstVw.SubItems(5) = txtCost.text
        
        dbConn.Open getConnectionString
        
        SQLStr = "SELECT ID, UnitNumber, MaintenanceType, " & _
                  "ReportedDate, Description, DateCompleted, TaskOwner, " & _
                  "Contact, RemindDate, Alarm, EstimateCost " & _
                 "FROM UnitMaintHistory"
        rstRecordSet.Open SQLStr, dbConn, adOpenDynamic, adLockOptimistic

        rstRecordSet.AddNew

        rstRecordSet!UnitNumber = lblUnitID.Caption
'        rstRecordSet!MaintType = txtType.Text
        rstRecordSet!MaintType = cmbType.text
        rstRecordSet!RepDate = CDate(Format(txtRepDate.text, "dd/mm/yyyy"))
        rstRecordSet!description = txtDes.text
        rstRecordSet!WorkerName = txtPerson.text
        If txtCompDate.text <> "" Then
            rstRecordSet!ComplitedDate = CDate(Format(txtCompDate.text, "dd/mm/yyyy"))
        End If
        If txtCost.text <> "" Then
            rstRecordSet!Cost = CCur(txtCost.text)
        End If

        rstRecordSet.Update
        rstRecordSet.Close
        dbConn.Close

        cmdUpdate.Enabled = False
        cmdAddNew.Enabled = True
        DisableTextBoxes
        ShowMsgInTaskBar "New Data has updated successfully"
        Exit Sub
    End If
    
    lsvUnit.ListItems(inListView).text = txtRepDate.text
'    lsvUnit.SelectedItem.SubItems(1) = txtType.Text
    lsvUnit.SelectedItem.SubItems(1) = cmbType.text
    lsvUnit.SelectedItem.SubItems(2) = txtDes.text
    lsvUnit.SelectedItem.SubItems(3) = txtCompDate.text
    lsvUnit.SelectedItem.SubItems(4) = txtPerson.text
    lsvUnit.SelectedItem.SubItems(5) = txtCost.text
    
    dbConn.Open getConnectionString
    
    SQLStr = "SELECT ID, UnitNumber, MaintenanceType, " & _
                  "ReportedDate, Description, DateCompleted, TaskOwner, " & _
                  "Contact, RemindDate, Alarm, EstimateCost  " & _
             "FROM UnitMaintHistory " & _
             "WHERE UnitNumber = '" & lblUnitID.Caption & "'"
             
    rstRecordSet.Open SQLStr, dbConn, adOpenDynamic, adLockOptimistic
    
    rstRecordSet!MaintType = cmbType.text
    rstRecordSet!RepDate = txtRepDate.text
    rstRecordSet!description = txtDes.text
    rstRecordSet!WorkerName = txtPerson.text
    rstRecordSet!ComplitedDate = txtCompDate.text
    rstRecordSet!Cost = txtCost.text
    
    rstRecordSet.Update
    rstRecordSet.Close
    dbConn.Close
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then ShowMsgInTaskBar "hello"
End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50
    
    lsvUnit.ColumnHeaders.Clear
    lsvUnit.ColumnHeaders.Add 1, , "Reported Date", 1500
    lsvUnit.ColumnHeaders.Add 2, , "Type", 1500
    lsvUnit.ColumnHeaders.Add 3, , "Description", 4500
    lsvUnit.ColumnHeaders.Add 4, , "Completed Date", 1500
    lsvUnit.ColumnHeaders.Add 5, , "Person", 1100
    lsvUnit.ColumnHeaders.Add 6, , "Cost", 800

    dbConn.Open getConnectionString

    rstRecordSet.Open SQLStr, dbConn, adOpenStatic, adLockReadOnly

    If rstRecordSet.EOF Then ShowMsgInTaskBar "There is no maintainance history recoded for this UNIT", , "N"

    While Not rstRecordSet.EOF
        Dim iLstVw As ListItem
        Set iLstVw = lsvUnit.ListItems.Add(, , rstRecordSet!RepDate)
        iLstVw.SubItems(1) = rstRecordSet!MaintType
        iLstVw.SubItems(2) = rstRecordSet!description
        iLstVw.SubItems(3) = rstRecordSet!ComplitedDate & ""
        iLstVw.SubItems(4) = rstRecordSet!WorkerName
        iLstVw.SubItems(5) = rstRecordSet!Cost
        
        rstRecordSet.MoveNext
    Wend

    rstRecordSet.Close
    dbConn.Close
End Sub

Private Sub lsvUnit_Click()
    cmdEdit.Enabled = True
End Sub

Private Sub txtAddNew_Click()
End Sub

Private Sub lsvUnit_DblClick()
    ShowMsgInTaskBar lsvUnit.ListItems(lsvUnit.SelectedItem.Index)
    ShowMsgInTaskBar lsvUnit.SelectedItem.SubItems(5)
End Sub
