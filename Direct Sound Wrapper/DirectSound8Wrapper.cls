VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSound8Wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Enum FX
    [Effect Chorus] = 1
    [Effect Compress] = 2
    [Effect Distort] = 3
    [Effect Echo] = 4
    [Effect Flange] = 5
    [Effect Gargle] = 6
    [Effect Reverb] = 7
    [Effect EQ] = 8
End Enum

Private DX As DirectX8
Private DS As DirectSound8

Private Sub Class_Initialize()

    On Error GoTo ErrHandler

    Set DX = New DirectX8
    Set DS = DX.DirectSoundCreate(vbNullString)

    'Simple error check
    If DX Is Nothing Then GoTo ErrHandler
    If DS Is Nothing Then GoTo ErrHandler

    Load fCallback

    'Set cooperative level to normal
    DS.SetCooperativeLevel fCallback.hWnd, DSSCL_NORMAL

    'Add an empty "Slot"
    ReDim Sounds(1)

Exit Sub

ErrHandler:
    Err.Raise vbObjectError, "DirectSound8Wrapper.Class_Initialise", "Unable to initialise DirectSound"
    Class_Terminate

End Sub

Private Sub Class_Terminate()

  'If objects exist, delete them

    If ObjPtr(DS) Then Set DS = Nothing
    If ObjPtr(DX) Then Set DX = Nothing

    'Erase the full sound array
    Erase Sounds

    'Unload the callback form
    Unload fCallback

End Sub

'Load a sound effect into a buffer
Public Function CreateSoundBuffer(ByVal SoundName As String, Optional ByVal FromResourceFile As Boolean = False, Optional ByVal DeleteDuplicates As Boolean = True, Optional ByVal UseFrequency As Boolean = False, Optional ByVal UseSFX As Boolean = False) As Integer

  Dim BufferDesc As DSBUFFERDESC
  Dim DSPosition(1) As DSBPOSITIONNOTIFY
  Dim DSNotification As Long
  Dim i As Integer, j As Integer

    If UseSFX Then
        BufferDesc.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFX Or DSBCAPS_CTRLPAN
      ElseIf UseFrequency Then
        BufferDesc.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN
      Else
        BufferDesc.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN
    End If

    'Look for an unused "slot" and if none found, create a new one
    For i = 1 To UBound(Sounds)
        If Sounds(i).DSBuffer Is Nothing Then
            Exit For
          ElseIf i = UBound(Sounds) Then
            ReDim Preserve Sounds(UBound(Sounds) + 1)
            i = i + 1
            Exit For
        End If
    Next i

    CreateSoundBuffer = i

    If Not FromResourceFile Then
        Set Sounds(i).DSBuffer = DS.CreateSoundBufferFromFile(SoundName, BufferDesc)
      Else
        Set Sounds(i).DSBuffer = DS.CreateSoundBufferFromResource(vbNullString, SoundName, BufferDesc)
    End If

    Sounds(i).Path = SoundName

    Sounds(i).Notification = DX.CreateEvent(fCallback)
    DSPosition(1).hEventNotify = Sounds(i).Notification
    DSPosition(1).lOffset = DSBPN_OFFSETSTOP
    Sounds(i).DSBuffer.SetNotificationPositions 1, DSPosition()

    If DeleteDuplicates Then
        'If duplicate sound detected then delete the new one and return index of old one
        For j = 1 To UBound(Sounds)
            If i <> j And Sounds(i).DSBuffer Is Sounds(j).DSBuffer And Sounds(i).Path = Sounds(j).Path Then
                Set Sounds(i).DSBuffer = Nothing

                If i = UBound(Sounds) Then
                    ReDim Preserve Sounds(i - 1)
                End If

                CreateSoundBuffer = j

                Exit Function
            End If
        Next j
    End If

End Function

Public Sub PlaySound(ByVal Index As Integer)

    Sounds(Index).DSBuffer.Play DSBPLAY_DEFAULT
    Sounds(Index).Playing = True

End Sub

Public Sub StopSound(ByVal Index As Integer)

    Sounds(Index).DSBuffer.Stop
    Sounds(Index).Playing = False

End Sub

'Remove all sound effects from a buffer given the index
Public Sub RemoveAllEffects(ByVal Index As Integer)

  Dim sFX(1) As DSEFFECTDESC
  Dim Result(1) As Long

    'Remove effects by passing 0 as the effects count
    If Sounds(Index).Playing Then Sounds(Index).DSBuffer.Stop: Sounds(Index).Playing = False
    Sounds(Index).DSBuffer.SetFX 0, sFX, Result

End Sub

'Delete a sound given its index
Public Sub RemoveSound(ByVal Index As Integer)

    Set Sounds(Index).DSBuffer = Nothing

End Sub

'Apply an effect to a sound given its index
Public Sub SetEffects(ByRef Effects() As FX, ByVal Index As Integer)

  Dim i As Integer
  Dim sFX() As DSEFFECTDESC
  Dim Result() As Long

    ReDim sFX(LBound(Effects) To UBound(Effects))
    ReDim Result(LBound(Effects) To UBound(Effects))

    For i = LBound(sFX) To UBound(sFX)
        'Get standard effect GUID
        Select Case Effects(i)
          Case [Effect Chorus]
            sFX(i).guidDSFXClass = DSFX_STANDARD_CHORUS
          Case [Effect Compress]
            sFX(i).guidDSFXClass = DSFX_STANDARD_COMPRESSOR
          Case [Effect Distort]
            sFX(i).guidDSFXClass = DSFX_STANDARD_DISTORTION
          Case [Effect Echo]
            sFX(i).guidDSFXClass = DSFX_STANDARD_ECHO
          Case [Effect Flange]
            sFX(i).guidDSFXClass = DSFX_STANDARD_FLANGER
          Case [Effect Gargle]
            sFX(i).guidDSFXClass = DSFX_STANDARD_GARGLE
          Case [Effect Reverb]
            sFX(i).guidDSFXClass = DSFX_STANDARD_WAVES_REVERB
          Case [Effect EQ]
            sFX(i).guidDSFXClass = DSFX_STANDARD_PARAMEQ
        End Select

        sFX(i).lFlags = 0
    Next i

    If Sounds(Index).Playing Then Sounds(Index).DSBuffer.Stop: Sounds(Index).Playing = False
    Sounds(Index).DSBuffer.SetFX UBound(sFX) - LBound(sFX) + 1, sFX, Result

End Sub

Public Sub SetPan(ByVal Index As Integer, ByVal Value As Long)

    If Value <= DSBPAN_RIGHT And Value >= DSBPAN_LEFT Then Sounds(Index).DSBuffer.SetPan Value

End Sub

Public Sub SetVFrequency(ByVal Index As Integer, ByVal Value As Long)

    If Value <= DSBFREQUENCY_MAX And Value >= DSBFREQUENCY_MIN Then Sounds(Index).DSBuffer.SetFrequency Value

End Sub

Public Sub SetVolume(ByVal Index As Integer, ByVal Value As Long)

    If Value <= DSBVOLUME_MAX And Value >= DSBVOLUME_MIN Then Sounds(Index).DSBuffer.SetVolume Value

End Sub

':) Ulli's VB Code Formatter V2.12.7 (24/05/2002 14:27:15) 16 + 186 = 202 Lines
