VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RES File"
   ClientHeight    =   3960
   ClientLeft      =   2190
   ClientTop       =   1380
   ClientWidth     =   1425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   1425
   Begin VB.CommandButton cmdSOUNDHD 
      Caption         =   "HD Sound"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Play .WAV from hard drive."
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdWIFE 
      Caption         =   "Wife"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDAUGHTER 
      Caption         =   "Daughter"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdWAV 
      Caption         =   "Extract &WAV"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSOUND 
      Caption         =   "&Play .WAV"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdJPGG 
      Caption         =   "&JPG Extract"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNEW 
      Caption         =   "&New DB"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Note, all remarks in this application were built
' with my Remark Builder Add-In which you can download
' at:  http://24.7.173.85/lafever
'------------------------------------------------------------

Option Explicit
Private Sub cmdCLOSE_Click()
    On Error Resume Next
    '------------------------------------------------------------
    ' Unload the form.
    '------------------------------------------------------------
    Unload Me
End Sub
Private Sub cmdDAUGHTER_Click()
    On Error Resume Next
    Dim frm As frmPIC
    '------------------------------------------------------------
    ' Instanciate a frmPIC form
    '------------------------------------------------------------
    Set frm = New frmPIC
    '------------------------------------------------------------
    ' Call it`s method to load a picture from the resource
    ' file embedded in the .EXE
    '------------------------------------------------------------
    frm.SetPicture appresMY_DAUGHTER
    '------------------------------------------------------------
    ' Show the form.
    '------------------------------------------------------------
    frm.Show
End Sub
Private Sub cmdWIFE_Click()
    On Error Resume Next
    Dim frm As frmPIC
    '------------------------------------------------------------
    ' Instanciate a frmPIC form.
    '------------------------------------------------------------
    Set frm = New frmPIC
    '------------------------------------------------------------
    ' Call it`s method to load a picture from the resource
    ' file embedded in the .EXE
    '------------------------------------------------------------
    frm.SetPicture appresMY_WIFE
    '------------------------------------------------------------
    ' Show the form.
    '------------------------------------------------------------
    frm.Show
End Sub
Private Sub cmdJPGG_Click()
    On Error Resume Next
    Dim fNAME As String
    Dim obj As CDIALOG
    Set obj = New CDIALOG
    '------------------------------------------------------------
    ' Get the save location
    '------------------------------------------------------------
    fNAME = obj.SaveFile("Extract JPG to", "*.jpg (JPG Image)|*.jpg", CurDir(), "*.jpg")
    '------------------------------------------------------------
    ' If a valid file name then
    '------------------------------------------------------------
    If InStr(fNAME, "*") = 0 Then
        If UCase(Right(fNAME, 4)) <> ".JPG" Then
            fNAME = fNAME & ".jpg"
        End If
        '------------------------------------------------------------
        ' Extract the file to the location given
        '------------------------------------------------------------
        BuildFileFromResource fNAME, appresMY_DAUGHTER, "JPG"
    End If
End Sub
Private Sub cmdNEW_Click()
    '------------------------------------------------------------
    ' Note that this embeded MDB file has a default
    ' table in it.  Now you can see how you can embed
    ' an .MDB for an application where you will be
    ' allowing the user to create new database files
    ' and know that each .MDB will contain all default
    ' objects and data you expect it to have without
    ' haveing to keeps some file on the users machine
    ' that you copy or having to write code to create
    ' tables and what not.
    '------------------------------------------------------------
    On Error Resume Next
    Dim fNAME As String
    Dim obj As CDIALOG
    Set obj = New CDIALOG
    '------------------------------------------------------------
    ' Get the save location
    '------------------------------------------------------------
    fNAME = obj.SaveFile("Create new MDB file", "*.mdb (Access Databases)|*.mdb", CurDir(), "*.mdb")
    '------------------------------------------------------------
    ' If a valid file name then
    '------------------------------------------------------------
    If InStr(fNAME, "*") = 0 Then
        If UCase(Right(fNAME, 4)) <> ".MDB" Then
            fNAME = fNAME & ".mdb"
        End If
        '------------------------------------------------------------
        ' Extract the .MDB out of the .EXE to the location
        ' given.
        '------------------------------------------------------------
        BuildFileFromResource fNAME, appresBLANK_MDB, "MDB"
    End If
End Sub
Private Sub cmdSOUND_Click()
    On Error Resume Next
    '------------------------------------------------------------
    ' Call function to play a .WAV from a resource
    ' file without having to extract it.
    '------------------------------------------------------------
    PlayWaveRes appsoundNT_LOGON_WAVE, soundASYNC
    '------------------------------------------------------------
    ' Watch out for large .WAV files or you will end
    ' up with a BIG EXE file.
    '------------------------------------------------------------
End Sub
Private Sub cmdSOUNDHD_Click()
    On Error Resume Next
    Dim obj As CDIALOG, objSOUND As CSOUND, fNAME As String
    Set obj = New CDIALOG
    '------------------------------------------------------------
    ' Get the full path and file name for .WAV to play.
    '------------------------------------------------------------
    fNAME = obj.GetFile("Play .WAV", "*.wav (WAV files)|*.wav", CurDir, "*.wav")
    '------------------------------------------------------------
    ' If a valid file name then
    '------------------------------------------------------------
    If InStr(fNAME, "*") = 0 And fNAME <> "" Then
        '------------------------------------------------------------
        ' Instanciate my CSOUND class pass it the filename
        ' and call the method to play.
        '------------------------------------------------------------
        Set objSOUND = New CSOUND
        objSOUND.SoundFile = fNAME
        objSOUND.Play SND_ASYNC
    End If
End Sub
Private Sub cmdWAV_Click()
    On Error Resume Next
    Dim fNAME As String
    Dim obj As CDIALOG
    Set obj = New CDIALOG
    '------------------------------------------------------------
    ' Get the save location
    '------------------------------------------------------------
    fNAME = obj.SaveFile("Extract .WAV File", "*.wav (.WAV Sound)|*.wav", CurDir(), "*.wav")
    '------------------------------------------------------------
    ' If a valid file name
    '------------------------------------------------------------
    If InStr(fNAME, "*") = 0 Then
        If UCase(Right(fNAME, 4)) <> ".WAV" Then
            fNAME = fNAME & ".wav"
        End If
        '------------------------------------------------------------
        ' Extract the .WAV to the location given.
        '------------------------------------------------------------
        BuildFileFromResource fNAME, appresNT_LOGON_SOUND, "WAVE"
    End If
End Sub

