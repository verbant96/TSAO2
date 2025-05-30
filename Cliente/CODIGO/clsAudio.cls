VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Modulo de clase by �Parra

Option Explicit
 
''
' Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffers As Byte = 30

Dim MP3 As Boolean
Dim DSAudio  As IBasicAudio
Dim DSEvent As IMediaEvent
Dim DSControl As IMediaControl
Dim DSPosition As IMediaPosition
 
''
' The distance at which the sound is inaudible.
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150
 
''
' The frequency variation coefficient for Doppler effect simulation.
Private Const DELTA_FQ As Integer = 600
 
''
' Custom sound buffer structure.
'
' @param    FileName    The name of the file loaded in this buffer
' @param    looping     Set to true if the sound will loop until stopped
' @param    X           The X position of the sound in map coordinates
' @param    Y           The Y position of the sound in map coordinates
' @param    normalFq    The frequency of the sound when not affected by Doppler effect.
' @param    Buffer      The DirectSoundBuffer object to wich the file is loaded
Private Type SoundBuffer
    Filename As String
    looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    Buffer As DirectSoundSecondaryBuffer8
End Type
''
' States how to set a sound's looping state.
'
' @param    Default     If wave is in memory use it's looping setting, otherwise false
' @param    Disabled    Never loop
' @param    Enabled     Loop until stopped
Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum
 
''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean
 
 
''
'Directory from which to load sound files by default
Dim SoundPath As String
 
''
'The volume for sounds (FXs)
Dim SndVolume As Long
 
''
'States wether music is supported by hardware or not
Dim Music As Boolean
 
''
'Directory from which to load music files by default
Dim MusicPath As String
 
''
'Last loaded music file. Used to be able to load it when music gets activated
Dim CurrentMusicFile As String
 
''
'States wether music is currently activated or not
Dim mActivated As Boolean
 
''
'States wether sound is currently activated or not
Dim sActivated As Boolean
 
''
'States the last position where the listener was in the X axis
Dim lastPosX As Integer
 
''
'States the last position where the listener was in the Y axis
Dim lastPosY As Integer
 
''
'States wether sound effects are currently activated or not
Dim sEffectsActivated As Boolean
 
'**** Direct Sound ********
 
Private dX As DirectX8
Private Sound_Buffer() As DirectSoundSecondaryBuffer8
Private DS As DirectSound8
Private DS_Enum As DirectSoundEnum8
 
'**** Direct Music ********
''
'DirectMusic's Performance object
Private Performance As DirectMusicPerformance8
 
''
'Currently loaded segment
Private Segment As DirectMusicSegment8
 
''
'The one and only DirectMusic Loader
Private Loader As DirectMusicLoader8
 
''
'State of the currently loaded segment
Private SegState As DirectMusicSegmentState8
 
''
'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer
 
''
'Constructor. Performs no action at all.
 
Private Sub Class_Initialize()
 
End Sub
 
''
'Destructor. Releases all created objects assuring no memory-leaks.
 
Private Sub Class_Terminate()
On Local Error Resume Next
    Dim loopc As Long
   
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To NumSoundBuffers
        Call DSBuffers(loopc).Buffer.Stop
        Set DSBuffers(loopc).Buffer = Nothing
    Next loopc
   
    'Stop music
    Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
   
    Call Performance.CloseDown
   
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
   
    inited = False
End Sub
 
''
'Initializes the clsAudio object.
Public Function Initialize(ByVal hWnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
On Error GoTo errHandler
   
    ' Prevent multiple creations
    If inited Then Exit Function
   
     'Cargamos
    Set dX = New DirectX8
   
    'Validate DX object
    If dX Is Nothing Then GoTo errHandler
   
    'Set sound path
    SoundPath = sound_path
   
    'Set default volume
    SndVolume = 0
   
    Set DS_Enum = dX.GetDSEnum
    Set DS = dX.DirectSoundCreate(DS_Enum.GetGuid(1))
 
    DS.SetCooperativeLevel hWnd, DSSCL_PRIORITY
   
    'Initialize DirectMusic
    Music = InitializedirectMusic(dX, hWnd)
   
    'Set music path
    MusicPath = music_path
   
    mActivated = True
    sActivated = True
   
    CurrentMusicFile = vbNullString
   
    'Set tempo to 0 and volume of music
    MusicTempo = 0
    Call Performance.SetMasterVolume(200)
   
    inited = True
   
    Initialize = True
Exit Function
 
errHandler:
End Function
 
''
' Creates and configures all DirectMusic objects.
 
Private Function InitializedirectMusic(ByRef dX As DirectX8, ByVal hWnd As Long) As Boolean
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo errHandler
    Dim mus_Params As DMUS_AUDIOPARAMS
    Set Loader = dX.DirectMusicLoaderCreate()
   
    Set Performance = dX.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(hWnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
   
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
   
    InitializedirectMusic = True
Exit Function
 
errHandler:
End Function
 
''
' Plays a new WAVE file.
 
Public Function PlayWave(ByVal Filename As String, Optional ByVal srcX As Byte = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
On Error GoTo errHandler
    Dim bufferIndex As Long
    
    If Not Sound Then Exit Function
   
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(Filename, LoopSound)
    If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
   
    With DSBuffers(bufferIndex)
        If .looping Then
            .Buffer.Play DSBPLAY_LOOPING
        Else
            .Buffer.Play DSBPLAY_DEFAULT
        End If
       
        'Apply volume
        Call .Buffer.SetVolume(SndVolume)
       
        'Store position
        .X = srcX
        .Y = srcY
       
        'Get standard frequency
        .normalFq = .Buffer.GetFrequency()
    End With
   
    PlayWave = bufferIndex
   
    If sEffectsActivated Then
        If srcX <> 0 And srcY <> 0 Then
            Call Update3DSound(bufferIndex, 0, 0)
        End If
    End If
Exit Function
 
errHandler:
End Function
 
Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance As Single
    Dim distanceXOld As Integer
    Dim distanceYOld As Integer
    Dim distanceXNew As Integer
    Dim distanceYNew As Integer
    Dim perceivedFq As Long
   
    With DSBuffers(bufferIndex)
        distanceXOld = .X - lastPosX
        distanceYOld = .Y - lastPosY
       
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
       
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
       
        deltaDistance = linearDistanceNew - linearDistanceOld
       
        'Update frequency emulating Doppler effect
        perceivedFq = .normalFq - deltaDistance * DELTA_FQ
        If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
       
        Call .Buffer.SetFrequency(perceivedFq)
       
        'Set volumen amortiguation according to distance
        Call .Buffer.SetVolume(SndVolume * (1 - linearDistanceNew / MAX_DISTANCE_TO_SOURCE))
       
        'Prevent division by zero
        If linearDistanceNew = 0 Then linearDistanceNew = 1
       
        'Set panning according to relative position of the source to the listener
        Call .Buffer.SetPan((distanceXNew / linearDistanceNew) * DSBPAN_RIGHT)
    End With
End Sub
 
''
' Updates 3D sounds based on the movement of the listener.
 
Public Sub MoveListener(ByVal X As Integer, ByVal Y As Integer)
    Dim i As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
   
    If Not sEffectsActivated Then Exit Sub
   
    deltaX = X - lastPosX
    deltaY = Y - lastPosY
   
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(i).X <> 0 And DSBuffers(i).Y <> 0 Then
                    Call Update3DSound(i, deltaX, deltaY)
                End If
            End If
        End If
    Next i
   
    lastPosX = X
    lastPosY = Y
End Sub
 
''
' Loads a new WAVE file.
 
Private Function LoadWave(ByVal Filename As String, ByVal looping As LoopStyle) As Long
On Error GoTo errHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
   
    Filename = UCase$(Filename)
   
    If Not FileExist(SoundPath & Filename, vbArchive) Then _
        Exit Function
   
    'Check if the buffer is in memory and not playing
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Filename = Filename Then
            If DSBuffers(i).Buffer.GetStatus = 0 Or DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = i
               
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(i).looping = (looping = LoopStyle.Enabled)
               
                Exit Function
            End If
        End If
    Next i
   
    'Not in memory, search for an empty buffer
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer Is Nothing Then
            Exit For
        End If
    Next i
   
    'No empty buffers?? Search for one not being played and use that slot
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).Buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
   
    'Still no free buffers?? Get the first non-looping sound
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If Not (DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).Buffer.Stop
                Exit For
            End If
        Next i
    End If
   
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If i > NumSoundBuffers Then
        If Not looping Then Exit Function
        i = 1
    End If
   
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
   
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 32
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
   
    With DSBuffers(i)
        Set .Buffer = Nothing   'Get rid of any previous data
       
        .Filename = Filename
        .looping = (looping = LoopStyle.Enabled)
       
        Set .Buffer = DS.CreateSoundBufferFromFile(SoundPath & Filename, bufferDesc) ' waveFormat
    End With
   
    LoadWave = i
Exit Function
 
errHandler:
End Function
 
''
' Stops a given sound or all of them.
 
Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).Buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To NumSoundBuffers
            If Not DSBuffers(i).Buffer Is Nothing Then
                If DSBuffers(i).Buffer.GetStatus() <> 0 And DSBuffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub
 
''
' Retrieves wether there are sounds currentyl playing or not.
 
Public Property Get PlayingSound() As Boolean
    Dim i As Long
   
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            PlayingSound = True
            Exit Property
        End If
    Next i
End Property
 
''
' Sets the volume of sound.
 
Public Property Let SoundVolume(ByVal Volume As Byte)
    Dim i As Long
   
    If Volume > 100 Then Exit Property
   
    'Take percentage to actual value
    SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
   
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).Buffer.SetVolume(SndVolume)
            End If
        End If
    Next i
End Property
 
''
' Gets the volume of sound.
 
Public Property Get SoundVolume() As Byte
    SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))
End Property
 
''
' Plays a new MIDI file.
 
Public Function PlayMIDI(Optional ByVal File As String = "", Optional ByVal Loops As Long = -1) As Boolean
On Error GoTo errHandler
    If LenB(File) > 0 Then _
        CurrentMusicFile = File

    If Not Sound Then Exit Function
    If Configuracion.Sound = 0 Then Exit Function
   
    If LenB(File) > 0 Then
        If Not LoadMidi(File) Then Exit Function
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
   
    'Play it
    Call Segment.SetRepeats(Loops)
   
    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
   
    PlayMIDI = True
Exit Function
 
errHandler:
End Function
 
''
' Loads a new MIDI file.
 
Private Function LoadMidi(ByVal File As String) As Boolean
On Error GoTo errHandler
    If Not FileExist(MusicPath & File, vbArchive) Then Exit Function
   
    Call StopMidi
   
    'Destroy old object
    Set Segment = Nothing
   
    Set Segment = Loader.LoadSegment(MusicPath & File)
   
    If Segment Is Nothing Then
        Exit Function
    End If
   
    Call Segment.SetStandardMidiFile
   
    LoadMidi = True
Exit Function
 
errHandler:
End Function
 
''
' Stops playing the currently loaded MIDI file.
 
Public Sub StopMidi()
On Error GoTo errHandler
    If PlayingMusic Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    End If
Exit Sub
 
errHandler:
End Sub
 
''
' Sets the music's tempo.
 
Public Property Let MusicTempo(ByVal Tempo As Single)
    If Not Music Then Exit Property
   
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
   
    Call Performance.SetMasterTempo(Tempo)
End Property
 
''
' Sets the music's volume.
 
Public Property Let MusicVolume(ByVal Volume As Long)
    If Not Music Then Exit Property
   
    If Volume < 0 Or Volume > 100 Then Exit Property
   
    ' Volume ranges from -10000 to 10000
    Call Performance.SetMasterVolume(Volume * 200 - 10000)
End Property
 
''
' Gets the music's tempo.
 
Public Property Get MusicTempo() As Single
    If Not Music Then Exit Property
   
    MusicTempo = Performance.GetMasterTempo()
End Property
 
''
' Gets the music's volume.
 
Public Property Get MusicVolume() As Long
    If Not Music Then Exit Property
   
    MusicVolume = (Performance.GetMasterVolume() + 10000) / 200
End Property
 
''
' Checks wether there is music currently playing.
 
Public Property Get PlayingMusic() As Boolean
    If Not Music Or Segment Is Nothing Then Exit Property
   
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property
 
''
' Retrieves the music's length.
 
Public Property Get MusicLenght() As Long
    If Not Music Then Exit Property
   
    MusicLenght = Segment.GetLength()
End Property
 
''
' Retrieves wether music is activated or not.
 
Public Property Get MusicActivated() As Boolean
    MusicActivated = mActivated
End Property
 
''
' Sets wether music is activated or not.
 
Public Property Let MusicActivated(ByVal nValue As Boolean)
    If mActivated = nValue Then Exit Property
   
    mActivated = nValue
   
    If nValue Then
        Call PlayMIDI(CurrentMusicFile)
    Else
        Call StopMidi
    End If
End Property
 
''
' Retrieves wether sound is activated or not.
 
Public Property Get SoundActivated() As Boolean
    SoundActivated = sActivated
End Property
 
''
' Sets wether sound is activated or not.
 
Public Property Let SoundActivated(ByVal nValue As Boolean)
    If sActivated = nValue Then Exit Property
   
    sActivated = nValue
   
    If Not nValue Then
        Call StopWave
    End If
End Property
 
''
' Sets wether sound effects are activated or not.
 
Public Property Let SoundEffectsActivated(ByVal nValue As Boolean)
    sEffectsActivated = nValue
End Property
 
''
' Retrieves wether sound effects are activated or not.
 
Public Property Get SoundEffectsActivated() As Boolean
    SoundEffectsActivated = sEffectsActivated
End Property
 

Function MP3_Load(ByVal Filename As String) As Boolean
 
If Not FileExist(App.Path & "\Data\SOUNDS\MP3\" & Filename & ".mp3", vbNormal) Then Exit Function

Set DSControl = New FilgraphManager
Call DSControl.RenderFile(App.Path & "\Data\SOUNDS\MP3\" & Filename & ".mp3")
 
Set DSAudio = DSControl

DSAudio.Volume = (-4000) + 3900
DSAudio.Balance = 0
   
Set DSEvent = DSControl
Set DSPosition = DSControl
 
DSPosition.Rate = 1#
DSPosition.CurrentPosition = 0
 
End Function
 
Public Function MP3_Play(ByVal Filename As String) As Boolean

'If Not Musica Then Exit Function
If Not FileExist(App.Path & "\Data\SOUNDS\MP3\" & Filename & ".mp3", vbNormal) Then Exit Function
If Configuracion.Music = 0 Then Exit Function
If frmMain.Visible = True And Musica = False Then Exit Function

If MP3 = True Then _
MP3_Stop
 
MP3 = MP3_Load(Filename)
 
DSControl.Run
 
MP3_Play = True
 
Exit Function
 
End Function
 
Public Sub MP3_Stop()
  If MP3 = True Then
    DSControl.Stop
    DSPosition.CurrentPosition = 0
    
    MP3 = False
  End If
Exit Sub
End Sub
 
Public Sub MP3_Destroy()
 
If MP3 = True Then _
    MP3_Stop
 
Set DSAudio = Nothing
Set DSEvent = Nothing
Set DSControl = Nothing
Set DSPosition = Nothing
End Sub


