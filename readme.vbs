Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file

Dim oFileStream, oVoice

text = InputBox("Enter the to be converted text","Somesh's Script")

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\Users\flash\Desktop\T2A\Sample.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.Voice = oVoice.GetVoices.Item(0)
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak text

oFileStream.Close