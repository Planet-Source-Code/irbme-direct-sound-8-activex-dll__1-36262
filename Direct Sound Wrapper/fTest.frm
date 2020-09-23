VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddGarglr 
      Caption         =   "Play With Gargle"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEcho 
      Caption         =   "Play With Echo"
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3480
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Sound"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SW As DirectSound8Wrapper
Attribute SW.VB_VarHelpID = -1

Dim Sound1 As Integer

Private Sub cmdAdd_Click()

    CD.ShowOpen
    
    'SoundIndex = CreateSOundBuffer(Filename,From A Resource File?, _
    Delete duplicate files?, Use frequency?, Use Effects?
    
    'Note you cannot use frequency and effects at the same time. _
    It is either one or neither
    Sound1 = SW.CreateSoundBuffer(CD.FileName, False, True, False, True)

End Sub

Private Sub cmdAddGarglr_Click()

  'You can apply multiple effects at the same time by increasing the number
  'of elements in the array then setting each element to the desired effect
  Dim Effects(0) As FX

    Effects(0) = [Effect Gargle]
    
    If Sound1 > 0 Then
        SW.SetEffects Effects, Sound1
        SW.PlaySound Sound1
    End If

End Sub

Private Sub cmdEcho_Click()

  'You can apply multiple effects at the same time by increasing the number
  'of elements in the array then setting each element to the desired effect
  Dim Effects(0) As FX

    Effects(0) = [Effect Echo]
    
    If Sound1 > 0 Then
        SW.SetEffects Effects, Sound1
        SW.PlaySound Sound1
    End If

End Sub

Private Sub cmdPlay_Click()

    If Sound1 > 0 Then
        SW.PlaySound Sound1
    End If

End Sub

Private Sub cmdStop_Click()

    'Play any amount of sounds at hte same time just by giving its index.

    If Sound1 > 0 Then
        SW.StopSound Sound1
    End If
    
End Sub

Private Sub Form_Load()

    Set SW = New DirectSound8Wrapper

End Sub
