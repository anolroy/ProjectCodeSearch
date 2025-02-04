VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSelPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin MSForms.ComboBox cmbPrinter 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3285
      VariousPropertyBits=   1753237531
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5794;503"
      TextColumn      =   1
      ColumnCount     =   2
      ListRows        =   20
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Myriad Web"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "3527;0"
   End
End
Attribute VB_Name = "frmSelPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function PopulateListControlWithPrinters(ListControl As Object) As Boolean
   On Error GoTo errHandler:
   Dim l As Long
   Dim lCount As Long
   ListControl.Clear

   lCount = Printers.count

   If lCount = 0 Then
       ListControl.AddItem "(No Printer Installed)"
   Else
       For l = 0 To lCount - 1
           ListControl.AddItem Printers(l).DeviceName
       Next
   End If

   PopulateListControlWithPrinters = True

   Exit Function
errHandler:
   PopulateListControlWithPrinters = False
   Exit Function
End Function

Private Sub Form_Load()
   PopulateListControlWithPrinters cmbPrinter
End Sub
