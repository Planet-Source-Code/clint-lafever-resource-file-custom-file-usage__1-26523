VERSION 5.00
Begin VB.Form frmPIC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPG Picture"
   ClientHeight    =   3195
   ClientLeft      =   6630
   ClientTop       =   1035
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox picIMAGE 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmPIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' In case you did not know, when you use the Picture
' Property while developing your applicaiton and
' you choose a picture, reguardless of file format,
' VB will embed it as a .BMP in your program.
' Imagine how big some programs get with all those
' .BMP images in it.  Now this demo shows how you
' can embed a .JPG literally in the EXE and extract
' it at runtime to the local drive, assign it to
' the Picture Property of some object then kill
' it from the local drive.  This method allows
' you to embed your graphics into your EXE without
' causing you file size to grow at such a large
' rate.
'------------------------------------------------------------



'------------------------------------------------------------
' Author:  Clint LaFever - [lafeverc@saic.com]
' Date: December,29 1999 @ 15:16:08
'------------------------------------------------------------
Public Sub SetPicture(pSEL As AppResource)
    On Error GoTo ErrorSetPicture
    Dim fNAME As String
    fNAME = BuildFileFromResource(App.Path & "\temp.JPG", pSEL, "JPG")
    Me.picIMAGE.Picture = LoadPicture(fNAME)
    Kill fNAME
    Me.Width = Me.picIMAGE.Width + 80
    Me.Height = Me.picIMAGE.Height + 240
    Exit Sub
ErrorSetPicture:
    MsgBox Err & ":Error in SetPicture.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
Private Sub Form_Load()
    On Error Resume Next
    '------------------------------------------------------------
    ' This is where I load the icon from the resource
    ' file for this form.
    '------------------------------------------------------------
    SetFormIcon Me, appICON_INFO
End Sub
