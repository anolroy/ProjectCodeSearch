VERSION 5.00
Begin VB.Form frmEventType 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Type"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   Icon            =   "frmEventType.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Demand Type"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Changes"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New Demand Type"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Demand Type"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdsaveNew 
      Caption         =   "&Save New Demand Type"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelNew 
      Caption         =   "&Cancel New Demand Type"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboDemand 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox txtType 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lable1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFDFC0&
      Caption         =   "Select Demand Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFDFC0&
      Caption         =   "Demand Type:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFDFC0&
      Caption         =   "ID:"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "frmEventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50

End Sub
