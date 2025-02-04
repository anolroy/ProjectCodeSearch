VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   ClientHeight    =   8745
   ClientLeft      =   135
   ClientTop       =   510
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      lastProp        =   500
      _cx             =   26908
      _cy             =   19500
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  frmReport
'  Shows the crystal report (invoice)
'=========================================================================================
'  Created By: Tania Sanchez - PCM Consulting
'  Published Date: 29/07/2004
'  WebSite: www.pcmconsulting.com
'  Legal Copyright: Tania Sanchez � 29/07/2004
'=========================================================================================
Option Explicit

'Private Sub CRViewer91_PrintButtonClicked(UseDefault As Boolean)
'        frmListAll.Report.PaperOrientation = crPortrait
'        frmListAll.Report.PrinterSetup Me.hWnd
'End Sub

'=========================================================================================
Private Sub Form_Resize()
On Error GoTo ErrorTrap
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.WhatsThisHelpID = ScaleWidth

Exit Sub

ErrorTrap:
MsgBox "Error Number: " & err.Number & vbCrLf & err.Description & vbCrLf & vbCrLf & "Debug Information:" & vbCrLf & _
"BDLInvoice.frmReport.Form_Resize" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"

End Sub 'Form_Resize()
'=========================================================================================
Public Sub LoadReportViewer(reportSource As CRAXDRT.Report)
    'Set the report source for the Report Viewer to the Report
On Error GoTo ErrorTrap
    CRViewer91.reportSource = reportSource
    CRViewer91.ViewReport

    CRViewer91.EnableExportButton = True
    CRViewer91.EnableGroupTree = False
    CRViewer91.DisplayTabs = False
    CRViewer91.EnableCloseButton = False
    CRViewer91.EnableProgressControl = True
    Me.Caption = "Goldshield Cheque"
    Me.Show
Exit Sub

ErrorTrap:
MsgBox "Error Number: " & err.Number & vbCrLf & err.Description & vbCrLf & vbCrLf & "Debug Information:" & vbCrLf & _
"BDLInvoice.frmReport.LoadReportViewer" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"

End Sub      'LoadReportViewer(reportSource As CRAXDRT.Report)
'=========================================================================================


