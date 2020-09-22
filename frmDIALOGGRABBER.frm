VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmDIALOGGRABBER 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCOLOR 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
End
Attribute VB_Name = "frmDIALOGGRABBER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CloseDialogGrabber()
    On Error Resume Next
    Unload Me
End Sub
