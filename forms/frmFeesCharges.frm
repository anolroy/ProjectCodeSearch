VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFeesCharges 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fees & Charges"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFeesCharges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin MSForms.CommandButton cmdRP 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   2295
      Caption         =   "Rent Payable"
      Size            =   "4048;873"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdMF 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Caption         =   "Management Fees"
      Size            =   "4048;873"
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmFeesCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMF_Click()
   Load frmFees
   frmFees.Show
   Unload Me
End Sub

Private Sub cmdRP_Click()
   Load frmRentPayable
   frmRentPayable.Show
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Height = 1440
   Me.Width = 5475
   Me.BackColor = MODULEBACKCOLOR

   Me.Top = 500
   Me.Left = 500
End Sub
