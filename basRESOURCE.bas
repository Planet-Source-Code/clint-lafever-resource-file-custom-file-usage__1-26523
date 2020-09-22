Attribute VB_Name = "basRESOURCE"
'------------------------------------------------------------
' Note:  Most of this code can be found in my code
' depot add-in at http://lafever.iscool.net
'------------------------------------------------------------

Option Explicit
'------------------------------------------------------------
' These are only needed for playing WAV files.
'------------------------------------------------------------
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Play synchronously (default).
Private Const SND_NODEFAULT = &H2    ' Do not use default sound.
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8         ' Loop the sound until next
Private Const SND_NOSTOP = &H10      ' Do not stop any currently
Private Const SND_ASYNC = &H1          '  play asynchronously
Private bytSound() As Byte ' Always store binary data in byte arrays!
Public Enum SoundFlags
    soundSYNC = SND_SYNC
    soundNO_DEFAULT = SND_NODEFAULT
    soundMEMORY = SND_MEMORY
    soundLOOP = SND_LOOP
    soundNO_STOP = SND_NOSTOP
    soundASYNC = SND_ASYNC
End Enum
'------------------------------------------------------------
' This is the ENUM list of the files contained
' inside the Resource File.
'------------------------------------------------------------
Public Enum AppResource
    appresBLANK_MDB = 101
    appresMY_DAUGHTER = 102
    appresNT_LOGON_SOUND = 103
    appresMY_WIFE = 104
End Enum
'------------------------------------------------------------
' Enum list for icons in the resource file
'------------------------------------------------------------
Public Enum AppIcons
    appICON_INFO = 101
End Enum
'------------------------------------------------------------
' This is the ENUM list of .WAV files that are
' in the resource file to play.
'------------------------------------------------------------
Public Enum AppSounds
    appsoundNT_LOGON_WAVE = 103
End Enum
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@usa.net]
' Purpose:  Used to set a given form's Icon property to an icon from the Resource File.  Note the use of AppIcons
' Parameters:
' Example:
' Date: July,21 1998 @ 19:25:18
'------------------------------------------------------------
Public Sub SetFormIcon(frm As Form, lngICON As AppIcons)
    On Error Resume Next
    frm.Icon = LoadResPicture(lngICON, vbResIcon)
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [lafeverc@saic.com]
' Purpose:  To play .WAV files contained within a resource
'                file
' Parameters:  ID of .WAV file to play.  Flag of how to play .WAV file
' Date: October,18 1999 @ 11:45:29
'------------------------------------------------------------
Public Sub PlayWaveRes(vntResourceID As AppSounds, Optional vntFlags As SoundFlags = soundASYNC)
    bytSound = LoadResData(vntResourceID, "WAVE")
    If IsMissing(vntFlags) Then
        vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
    End If
    If (vntFlags And SND_MEMORY) = 0 Then
        vntFlags = vntFlags Or SND_MEMORY
    End If
    sndPlaySound bytSound(0), vntFlags
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [lafeverc@saic.com]
' Purpose:  Extracts a file from the resource file and save
'                the file to the destination passed in.
' Date: October,18 1999 @ 11:45:53
'------------------------------------------------------------
Public Function BuildFileFromResource(destFILE As String, resID As AppResource, Optional resTITLE As String = "CUSTOM") As String
    On Error GoTo ErrorBuildFileFromResource
    Dim resBYTE() As Byte
    resBYTE = LoadResData(resID, resTITLE)
    Open destFILE For Binary Access Write As #1
    Put #1, , resBYTE
    Close #1
    BuildFileFromResource = destFILE
    Exit Function
ErrorBuildFileFromResource:
    BuildFileFromResource = ""
    MsgBox Err & ":Error in BuildFileFromResource.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
