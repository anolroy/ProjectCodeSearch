VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRentReceipt 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipts Entry Form"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   Icon            =   "frmRentReceipt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11070
   Begin VB.Frame fraPayee 
      Height          =   2655
      Left            =   6000
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdPayeeClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   2200
         Width           =   855
      End
      Begin VB.CommandButton cmdPayeeOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   2200
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPayee 
         Height          =   1935
         Left            =   75
         TabIndex        =   3
         Top             =   180
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   3413
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdPayee 
      Caption         =   "v"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPay 
      Height          =   4935
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8705
      _Version        =   393216
      FixedCols       =   0
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmRentReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
