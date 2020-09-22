VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSOUND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Used to play .WAV files.
' Parameters:
' Example:
' Date: July,21 1998 @ 19:27:13
'------------------------------------------------------------
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private mstrSOUND As String
Public Enum SoundPlayingOptions
    SND_SYNC = &H0         '  play synchronously (default)
    SND_ASYNC = &H1         '  play asynchronously
    SND_LOOP = &H8         '  loop the sound until next sndPlaySound
End Enum
Public Property Get SoundFile() As String
    SoundFile = mstrSOUND
End Property
Public Property Let SoundFile(pstr As String)
    mstrSOUND = pstr
End Property
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Method to play a passed .WAV file or the objects .SoundFile Property.
' Parameters:
' Example:
' Date: July,21 1998 @ 19:27:48
'------------------------------------------------------------
Public Function Play(Optional sOPTION As SoundPlayingOptions = SND_ASYNC, Optional sNAME = "") As Long
    On Error Resume Next
    Dim ret As Long
    If IsNumeric(sNAME) Then
        Me.SoundFile = App.Path & LoadResString(sNAME)
        sNAME = ""
    End If
    If sNAME = "" Then sNAME = Me.SoundFile
    If sNAME <> "" Then
        ret = sndPlaySound(sNAME, sOPTION)
    End If
    Play = ret
End Function




