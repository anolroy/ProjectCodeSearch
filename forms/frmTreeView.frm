VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tree Contol"
   ClientHeight    =   9255
   ClientLeft      =   2775
   ClientTop       =   3645
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvwLandLord 
      Height          =   9255
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   16325
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeView.frx":1A8E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
   tvwLandLord.Top = 0
   tvwLandLord.Left = 0
   tvwLandLord.Width = Me.Width - 40
   tvwLandLord.Height = Me.Height - 340
End Sub

Private Sub Form_Load()
   Dim Conn1 As New RDO.rdoConnection
   Dim Rst1 As rdoResultset
   
   Conn1.Connect = "DSN=" & Adsn & ";UID=;PWD="
   Conn1.CursorDriver = rdUseIfNeeded
   Conn1.EstablishConnection rdDriverNoPrompt
   Set Rst1 = Conn1.OpenResultset("SELECT ClientID FROM Client", rdOpenStatic, rdConcurReadOnly)
   While Not Rst1.EOF
      DrawLandLordTree tvwLandLord, imgList, Rst1!ClientID, False
      Rst1.MoveNext
   Wend
   Rst1.Close
   Conn1.Close
   
   frmTreeView.Width = 1995
   frmMMain.ChildFormLeft = Me.Width
End Sub

Private Sub Form_Resize()
   tvwLandLord.Width = Me.Width - 40
   
   frmMMain.ChildFormLeft = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmMMain.mnuTreeView.Checked = False
   frmMMain.ChildFormLeft = 20
End Sub

Private Sub tvwLandLord_BeforeLabelEdit(Cancel As Integer)

End Sub
