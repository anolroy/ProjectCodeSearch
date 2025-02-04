VERSION 5.00
Begin VB.Form frmAdiComment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAdiComment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAdiComment 
      Appearance      =   0  'Flat
      Height          =   2355
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7425
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6450
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmAdiComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public szAdiComm As String

Private Sub cmdCancel_Click()
   frmDemands3.txtAdiComment.text = IIf(IsNull(szAdiComm), "", szAdiComm)

   Me.Hide
   frmDemands3.Enabled = True
   frmDemands3.txtAmount.SetFocus
End Sub

Private Sub cmdOK_Click()
   frmDemands3.SetComment txtAdiComment.text
   frmDemands3.txtAdiComment.text = Left(txtAdiComment.text, 20) & "..."

   Me.Hide
   frmDemands3.Enabled = True
   frmDemands3.txtAmount.SetFocus
End Sub

Private Sub Form_Activate()
   txtAdiComment.SetFocus
End Sub

Private Sub Form_Load()
   txtAdiComment.text = IIf(IsNull(szAdiComm), "", szAdiComm)
   Me.BackColor = MODULEBACKCOLOR
End Sub
