VERSION 5.00
Begin VB.Form fCallback 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "fCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

  Dim i As Integer

    'Find what sound we are being notified about
    For i = 1 To UBound(Sounds)
        If Sounds(i).Notification = eventid Then
            Exit For
        End If
    Next i

    Sounds(i).Playing = False

End Sub

':) Ulli's VB Code Formatter V2.12.7 (24/05/2002 13:56:18) 3 + 17 = 20 Lines
