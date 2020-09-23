Attribute VB_Name = "modSoundStruct"
Option Explicit
Option Base 1

Private Type DSOUND
    DSBuffer             As DirectSoundSecondaryBuffer8
    Notification         As Long
    ChannelVolume        As Long
    Playing              As Boolean
    Path                 As String
End Type

Public Sounds() As DSOUND

':) Ulli's VB Code Formatter V2.12.7 (24/05/2002 13:56:17) 14 + 0 = 14 Lines
