Attribute VB_Name = "modSound"
Option Explicit

Dim snd(1 To 9) As Long

Sub InitSound()
    FSOUND_Init 44100, 32, 0

    Dim i As Long, path As String
    For i = 1 To 9
        path = GetAssetPath("SFX/sound" + CStr(i) + ".wav")
        If Exists(path) Then
            snd(i) = FSOUND_Sample_Load(FSOUND_FREE, path, FSOUND_LOOP_OFF, 0, 0)
        End If
    Next i
End Sub

Sub UnloadSound()
    Dim i As Long, path As String
    For i = 1 To 9
        path = GetAssetPath("SFX/sound" + CStr(i) + ".wav")
        If Exists(path) Then
            If snd(i) > 0 Then
                FSOUND_Sample_Free snd(i)
            End If
        End If
    Next i
    FSOUND_Close
End Sub

Sub PlayWav(number As Long)
    If options.Wav = True Then
        If (number > 0 And number < 10) Then
            If snd(number) > 0 Then
                FSOUND_SetSFXMasterVolume 100
                FSOUND_PlaySound -1, snd(number)
            End If
        End If
    End If
End Sub

