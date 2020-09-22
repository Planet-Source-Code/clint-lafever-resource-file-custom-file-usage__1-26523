VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDIALOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**********************************************************************************
'This class requires a form named frmDIALOGGRABBER with a dialog box on it called dlgCOLOR
'**********************************************************************************
Private dlg As CommonDialog
Private frm As frmDIALOGGRABBER
Friend Sub StartForm()
    Set frm = New frmDIALOGGRABBER
    Set dlg = frm.dlgCOLOR
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  To get printer information from the user.
' Date: April,19 1999 @ 09:54:47
'------------------------------------------------------------
Public Function GetPrinter() As Long
    On Error GoTo ErrorGetPrinter
    StartForm
    With dlg
        .ShowPrinter
        GetPrinter = .PrinterDefault
    End With
    EndForm
    Exit Function
ErrorGetPrinter:
    MsgBox Err & ":Error in GetPrinter.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
Public Function GetFile(Optional Title As String = "", Optional filter As String = "*.* (All Files)|*.*", Optional defaultPath As String = "C:\", Optional FileNameText As String = "*.*")
    On Error Resume Next
    StartForm
    With dlg
        .filter = filter
        .FileName = FileNameText
        .Flags = cdlCCFullOpen
        .InitDir = defaultPath
        .DialogTitle = Title
        .Action = 1
        GetFile = .FileName
    End With
    EndForm
End Function
Public Function SaveFile(Optional Title As String = "", Optional filter As String = "*.* (All Files)|*.*", Optional defaultPath As String = "", Optional FileNameText As String = "*.*")
    On Error Resume Next
    If defaultPath = "" Then defaultPath = CurDir
    StartForm
    With dlg
        .filter = filter
        .FileName = FileNameText
        .InitDir = defaultPath
        .DialogTitle = Title
        .ShowSave
        SaveFile = .FileName
    End With
    EndForm
End Function
Public Function GetColor(OriginalColor As Long)
    On Error GoTo ErrHandler
    StartForm
    dlg.CancelError = True
    dlg.Flags = cdlCCRGBInit
    dlg.ShowColor
    GetColor = dlg.Color
    EndForm
    Exit Function
ErrHandler:
    GetColor = OriginalColor
    EndForm
End Function
Friend Sub EndForm()
    frm.CloseDialogGrabber
    Set frm = Nothing
End Sub

