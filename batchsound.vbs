' BatchSound by Nethuja Gunawardane

If WScript.Arguments.Count = 0 Then
    MsgBox "No file given", vbInformation, "BatchSound"
    WScript.Quit
End If

audioFile = WScript.Arguments(0)

' File extension validation
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(audioFile) Then
    WScript.Quit
End If

fileExtension = LCase(fso.GetExtensionName(audioFile))
If fileExtension <> "mp3" And fileExtension <> "wav" And fileExtension <> "wma" Then
    WScript.Quit
End If

' Create WMPlayer.OCX object
Set objPlayer = CreateObject("WMPlayer.OCX")

' Load and play the audio file
Set objMedia = objPlayer.newMedia(audioFile)
objPlayer.currentMedia = objMedia
objPlayer.controls.play

' Keep the script running in the background to allow the sound to play
Do While objPlayer.playState <> 1 ' 1 = Stopped
    WScript.Sleep 100
Loop

' Quick clean up after playing
Set objPlayer = Nothing
Set objMedia = Nothing
