Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "C:\Windows\Media\Windows Critical Stop.wav"
Dim oPlayer2
Set oPlayer2 = CreateObject("WMPlayer.OCX")
oPlayer2.URL = "C:\Windows\Media\Windows Exclamation.wav"
Dim oPlayer3
Set oPlayer3 = CreateObject("WMPlayer.OCX")
oPlayer3.URL = "C:\Windows\Media\Windows Foreground.wav"
X=MsgBox("Error while creating encrypted files. The file information has been lost.",0+16,"Crypt4 Beta")
For i = 1 To 100 Step 1
  oPlayer.controls.play
  oPlayer2.controls.play
  oPlayer3.controls.play
  While oPlayer.playState <> 1
    WScript.Sleep 1
  Wend
Next
oPlayer.close